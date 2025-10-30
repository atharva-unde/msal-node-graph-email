require("dotenv").config();
const express = require("express");
const msal = require("@azure/msal-node");
const axios = require("axios");
const fs = require("fs").promises;
const path = require("path");
const cors = require("cors");

const app = express();
const port = process.env.PORT || 3001;
const clientId = process.env.AUTH_MICROSOFT_ENTRA_ID_ID;
const clientSecret = process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET;
const tenantId = "common"; // 'common' for multi-tenant
const redirectUri =
  process.env.AZURE_REDIRECT_URI ||
  "http://localhost:3000/auth/microsoft/callback";

app.use(express.json());

app.use(cors({
  origin: "http://localhost:3000",
  methods: ["GET", "POST", "PUT", "DELETE"],
  credentials: true,
}));

const msalClient = new msal.ConfidentialClientApplication({
  auth: {
    clientId: clientId,
    clientSecret: clientSecret,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
});

// Token storage configuration
const TOKEN_FILE_PATH = path.join(__dirname, 'tokens.json');

// Helper functions for error handling
const createError = (status, message) => {
  const error = new Error(message);
  error.status = status;
  return error;
};

const handleError = (error) => {
  return {
    statusCode: error.status || 500,
    body: { error: error.message || 'Internal Server Error' },
    success: false,
  };
};

// Token management functions
async function saveTokenToFile(tokenData) {
  try {
    const tokenInfo = {
      accessToken: tokenData.accessToken,
      refreshToken: tokenData.refreshToken || null,
      expiresOn: tokenData.expiresOn,
      savedAt: new Date().toISOString(),
      account: tokenData.account
    };
    
    await fs.writeFile(TOKEN_FILE_PATH, JSON.stringify(tokenInfo, null, 2));
    console.log('Token saved to file successfully');
  } catch (error) {
    console.error('Error saving token to file:', error);
  }
}

async function loadTokenFromFile() {
  try {
    const data = await fs.readFile(TOKEN_FILE_PATH, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    console.log('No existing token file found or error reading file');
    return null;
  }
}

function isTokenExpired(expiresOn) {
  const now = new Date();
  const expiry = new Date(expiresOn);
  // Add 5 minute buffer before expiry
  const bufferTime = 5 * 60 * 1000; 
  return now >= (expiry.getTime() - bufferTime);
}

async function getValidAccessToken() {
  try {
    // First, try to load existing token
    const savedToken = await loadTokenFromFile();
    
    if (savedToken && !isTokenExpired(savedToken.expiresOn)) {
      console.log('Using existing valid token');
      return savedToken.accessToken;
    }
    
    // If token is expired or doesn't exist, try to refresh
    if (savedToken && savedToken.refreshToken) {
      console.log('Token expired, attempting to refresh...');
      const newToken = await refreshAccessToken(savedToken.refreshToken);
      if (newToken) {
        return newToken;
      }
    }
    
    // If refresh fails or no refresh token, need new authorization
    throw new Error('No valid token available. Please re-authorize.');
    
  } catch (error) {
    console.error('Error getting valid access token:', error);
    throw error;
  }
}

async function refreshAccessToken(refreshToken) {
  try {
    const response = await msalClient.acquireTokenByRefreshToken({
      refreshToken: refreshToken,
      scopes: [
        "offline_access",
        "user.read",
        "openid",
        "profile",
        "email",
        "https://graph.microsoft.com/Mail.Send",
        "https://graph.microsoft.com/User.Read",
      ],
    });

    console.log("Token refreshed successfully");
    await saveTokenToFile(response);
    return response.accessToken;
  } catch (error) {
    console.error("Failed to refresh token:", error);
    return null;
  }
}

// Email sending function
async function sendEmailViaMicrosoft(emailData) {
  try {
    const accessToken = await getValidAccessToken();
    
    const emailPayload = {
      message: {
        subject: emailData.subject,
        body: {
          contentType: emailData.bodyType || "HTML",
          content: emailData.body
        },
        toRecipients: emailData.to.map(email => ({
          emailAddress: {
            address: email
          }
        })),
        ...(emailData.cc && emailData.cc.length > 0 && {
          ccRecipients: emailData.cc.map(email => ({
            emailAddress: {
              address: email
            }
          }))
        }),
        ...(emailData.bcc && emailData.bcc.length > 0 && {
          bccRecipients: emailData.bcc.map(email => ({
            emailAddress: {
              address: email
            }
          }))
        })
      }
    };

    const response = await axios.post(
      'https://graph.microsoft.com/v1.0/me/sendMail',
      emailPayload,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );

    console.log('Email sent successfully');
    return {
      success: true,
      message: 'Email sent successfully',
      messageId: response.headers['x-ms-request-id'] || 'unknown'
    };

  } catch (error) {
    console.error('Error sending email:', error.response?.data || error.message);
    
    // If it's an authentication error, the token might be invalid
    if (error.response?.status === 401) {
      throw new Error('Authentication failed. Please re-authorize the application.');
    }
    
    throw new Error(`Failed to send email: ${error.response?.data?.error?.message || error.message}`);
  }
}

async function generateAuthUrl(tracker) {
  try {
    const authCodeUrlParameters = {
      scopes: [
        "offline_access",
        "user.read",
        "openid",
        "profile",
        "email",
        "https://graph.microsoft.com/Mail.Send",
        "https://graph.microsoft.com/User.Read",
      ],
      redirectUri: redirectUri,
      prompt: "consent",
      state: JSON.stringify(tracker),
    };

    return await msalClient.getAuthCodeUrl(authCodeUrlParameters);
  } catch (error) {
    console.error("Error generating Microsoft auth URL:", error);
    throw new Error("Failed to generate Microsoft OAuth URL");
  }
}

exports.microsoftOauthRedirect = async (payload) => {
  try {
    const { code, state } = payload?.query || {};

    console.log("Microsoft OAuth Redirect Payload:", {
      code: code ? "PRESENT" : "MISSING",
      state: state ? "PRESENT" : "MISSING",
    });

    if (!code) {
      throw createError(400, "Authorization code is missing");
    }

    const tokenRequest = {
      code: code,
      scopes: [
        "offline_access",
        "user.read",
        "openid",
        "profile",
        "email",
        "https://graph.microsoft.com/Mail.Send",
        "https://graph.microsoft.com/User.Read",
      ],
      redirectUri: redirectUri,
    };

    const response = await msalClient.acquireTokenByCode(tokenRequest);

    // Save the token to file
    await saveTokenToFile(response);

    return {
      statusCode: 200,
      body: {
        message: "Authorization successful",
        expiresOn: response.expiresOn,
        account: response.account?.username
      },
      success: true,
    };
  } catch (error) {
    console.error("Error in Microsoft OAuth Redirect:", error);
    throw handleError(error);
  }
};

app.get("/api/mailer/email-servers/microsoft", async (req, res) => {
  try {
    const tracker = {
      timestamp: Date.now(),
      ip: req.ip,
    };
    const authUrl = await generateAuthUrl(tracker);
    res.redirect(authUrl);
  } catch (error) {
    console.error("Error in /api/mailer/email-servers/microsoft route:", error);
    res.status(500).send("Internal Server Error");
  }
});

app.get("/api/mailer/email-servers/microsoft/redirect", async (req, res) => {
  try {
    const result = await exports.microsoftOauthRedirect({ query: req.query });
    res.status(result.statusCode).json(result.body);
  } catch (error) {
    console.error(
      "Error in /api/mailer/email-servers/microsoft/redirect route:",
      error
    );
    res.status(500).send("Internal Server Error");
  }
});

// Route to send email
app.post("/api/send-email", async (req, res) => {
  try {
    const { to, cc, bcc, subject, body, bodyType } = req.body;
    
    // Validate required fields
    if (!to || !Array.isArray(to) || to.length === 0) {
      return res.status(400).json({ 
        success: false, 
        error: "Recipients (to) field is required and must be an array" 
      });
    }
    
    if (!subject) {
      return res.status(400).json({ 
        success: false, 
        error: "Subject is required" 
      });
    }
    
    if (!body) {
      return res.status(400).json({ 
        success: false, 
        error: "Email body is required" 
      });
    }

    const emailData = {
      to,
      cc: cc || [],
      bcc: bcc || [],
      subject,
      body,
      bodyType: bodyType || "HTML"
    };

    const result = await sendEmailViaMicrosoft(emailData);
    res.json(result);
    
  } catch (error) {
    console.error("Error sending email:", error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Route to check token status
app.get("/api/token-status", async (req, res) => {
  try {
    const savedToken = await loadTokenFromFile();
    
    if (!savedToken) {
      return res.json({
        hasToken: false,
        message: "No token found. Please authorize first."
      });
    }
    
    const isExpired = isTokenExpired(savedToken.expiresOn);
    
    res.json({
      hasToken: true,
      isExpired,
      expiresOn: savedToken.expiresOn,
      account: savedToken.account?.username || 'Unknown',
      message: isExpired ? "Token expired. Will refresh automatically on next email send." : "Token is valid"
    });
    
  } catch (error) {
    console.error("Error checking token status:", error);
    res.status(500).json({
      success: false,
      error: "Failed to check token status"
    });
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
  console.log(`Available endpoints:`);
  console.log(`- GET  /api/mailer/email-servers/microsoft - Start OAuth flow`);
  console.log(`- GET  /api/mailer/email-servers/microsoft/redirect - OAuth callback`);
  console.log(`- POST /api/send-email - Send email via Microsoft Graph`);
  console.log(`- GET  /api/token-status - Check token status`);
  console.log('');
  console.log('To test via CLI:');
  console.log('  node -e "require(\'./test-email\').checkTokenStatus()"');
  console.log('  node -e "require(\'./test-email\').testEmailSending()"');
});
