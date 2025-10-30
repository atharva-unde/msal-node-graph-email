/**
 * Test file to demonstrate how to use the Microsoft Email API
 * Make sure to:
 * 1. Set up your .env file with proper Microsoft Azure credentials
 * 2. Start the server with `node index.js`
 * 3. First authorize by visiting: http://localhost:3001/api/mailer/email-servers/microsoft
 * 4. Then use the functions below to send emails
 * node -e "require('./test-email').checkTokenStatus()"
 * node -e "require('./test-email').testEmailSending()"
 */

require("dotenv").config();
const axios = require("axios");
const BASE_URL = "http://localhost:3001";

async function checkTokenStatus() {
  try {
    const response = await axios.get(`${BASE_URL}/api/token-status`);
    console.log("Token Status:", response.data);
    return response.data;
  } catch (error) {
    console.error(
      "Error checking token status:",
      error.response?.data || error.message
    );
    return null;
  }
}

async function sendEmail(emailData) {
  try {
    const response = await axios.post(`${BASE_URL}/api/send-email`, emailData);
    console.log('Email sent successfully:', response.data);
    return response.data;
  } catch (error) {
    console.error('Error sending email:', error.response?.data || error.message);
    return null;
  }
}


async function testEmailSending() {
  console.log("Checking token status...");
  const tokenStatus = await checkTokenStatus();

  if (!tokenStatus || !tokenStatus.hasToken) {
    console.log("No token found. Please authorize first by visiting:");
    console.log(`${BASE_URL}/api/mailer/email-servers/microsoft`);
    return;
  }

  if (tokenStatus.isExpired) {
    console.log(
      "Token is expired but will be refreshed automatically when sending email."
    );
  }

  // Example email data
  const emailData = {
    to: [process.env.RECIPIENT_USER], // Replace with actual recipient
    cc: [], // Optional
    bcc: [], // Optional
    subject: "Test Email from Microsoft Graph API",
    body: "<h1>Hello!</h1><p>This is a test email sent via Microsoft Graph API.</p>",
    bodyType: "HTML", // or 'Text'
  };

  console.log("Sending email...");
  const result = await sendEmail(emailData);

  if (result && result.success) {
    console.log("✅ Email sent successfully!");
  } else {
    console.log("❌ Failed to send email");
  }
}

module.exports = {
  checkTokenStatus,
  testEmailSending,
};
