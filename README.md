# Microsoft Graph API Email Service

A Node.js application that provides automated email sending capabilities through Microsoft Graph API with token management, automatic refresh, and persistent storage.

## 🚀 Features

- **OAuth 2.0 Authentication** with Microsoft Azure
- **Automatic Token Management** - tokens are stored locally and refreshed automatically
- **Email Sending** via Microsoft Graph API
- **Token Expiration Handling** - automatic refresh with 5-minute buffer
- **Persistent Token Storage** - tokens saved to JSON file
- **RESTful API** endpoints for easy integration
- **Error Handling** and validation

## 📋 Prerequisites

Before you begin, ensure you have:

- Node.js (v14 or higher)
- npm or yarn package manager
- Microsoft Azure account
- Azure App Registration with proper permissions

## 🛠️ Setup Instructions

### 1. Clone and Install Dependencies

```bash
git clone https://github.com/atharva-unde/msal-node-graph-email.git
cd msal-node-graph-email
npm install
```

### 2. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: Your app name
   - **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
   - **Redirect URI**: `http://localhost:3001/api/mailer/email-servers/microsoft/redirect`

5. After creation, note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

6. Go to **Certificates & secrets** → **New client secret**
   - Note down the **client secret value**

7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
   - Add these permissions:
     - `Mail.Send`
     - `User.Read`
     - `offline_access`
     - `openid`
     - `profile`
     - `email`

8. Click **Grant admin consent** (if you're an admin)

### 3. Environment Configuration

Create a `.env` file in the project root:

```env
# Microsoft Azure Configuration
AUTH_MICROSOFT_ENTRA_ID_ID=your_application_client_id_here
AUTH_MICROSOFT_ENTRA_ID_SECRET=your_client_secret_here

# Server Configuration
PORT=3001
AZURE_REDIRECT_URI=http://localhost:3001/api/mailer/email-servers/microsoft/redirect
```

Replace the placeholder values with your actual Azure app credentials.

### 4. Start the Server

```bash
npm start
# or
node index.js
```

The server will start on `http://localhost:3001`

## 📚 How It Works

### Token Management Flow

1. **Initial Authorization**: User visits the OAuth endpoint to authorize the application
2. **Token Storage**: Access and refresh tokens are saved to `tokens.json`
3. **Automatic Refresh**: Before each email send, the system checks token validity
4. **Token Renewal**: If expired, the refresh token is used to get a new access token
5. **Fallback**: If refresh fails, the user needs to re-authorize

### Token Lifecycle

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   User visits   │    │  OAuth callback  │    │ Token saved to  │
│ /api/.../microsoft │──▶│ receives tokens │──▶│   tokens.json   │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                                                         │
                                                         ▼
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│  Email request  │    │  Check token     │    │ Token valid?    │
│   comes in      │──▶│   validity       │──▶│ Use existing    │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                                                         │
                                                         ▼ (if expired)
                                               ┌─────────────────┐
                                               │ Refresh token   │
                                               │ automatically   │
                                               └─────────────────┘
```

## 🔌 API Endpoints

### 1. Start OAuth Flow
```http
GET /api/mailer/email-servers/microsoft
```
Redirects user to Microsoft OAuth consent screen.

### 2. OAuth Callback (Internal)
```http
GET /api/mailer/email-servers/microsoft/redirect
```
Handles OAuth callback and stores tokens.

### 3. Send Email
```http
POST /api/send-email
Content-Type: application/json

{
  "to": ["recipient@example.com"],
  "cc": ["cc@example.com"],           // Optional
  "bcc": ["bcc@example.com"],         // Optional
  "subject": "Your Subject Here",
  "body": "<h1>Hello World!</h1>",
  "bodyType": "HTML"                  // "HTML" or "Text"
}
```

### 4. Check Token Status
```http
GET /api/token-status
```
Returns current token status and expiration info.

## 💻 Usage Examples

### Using cURL

1. **Authorize the application** (visit in browser):
```bash
http://localhost:3001/api/mailer/email-servers/microsoft
```

2. **Check token status**:
```bash
curl -X GET http://localhost:3001/api/token-status
```

3. **Send an email**:
```bash
curl -X POST http://localhost:3001/api/send-email \
  -H "Content-Type: application/json" \
  -d '{
    "to": ["recipient@example.com"],
    "subject": "Test Email",
    "body": "<h1>Hello!</h1><p>This is a test email.</p>",
    "bodyType": "HTML"
  }'
```

### Using the Test Script

The project includes a test script (`test-email.js`):

```bash
# Check token status
node -e "require('./test-email').checkTokenStatus()"

# Send a test email (modify recipients in test-email.js first)
node -e "require('./test-email').testEmailSending()"
```

### Using Node.js

```javascript
const axios = require('axios');

async function sendEmail() {
  try {
    const response = await axios.post('http://localhost:3001/api/send-email', {
      to: ['recipient@example.com'],
      subject: 'Automated Email',
      body: '<h1>Hello from Node.js!</h1>',
      bodyType: 'HTML'
    });
    
    console.log('Email sent:', response.data);
  } catch (error) {
    console.error('Error:', error.response.data);
  }
}

sendEmail();
```

## 📁 File Structure

```
masl-debug/
├── index.js              # Main server file
├── test-email.js         # Testing utilities
├── tokens.json           # Token storage (created automatically)
├── package.json          # Dependencies
├── .env                  # Environment variables (create this)
├── .env.example          # Environment template
└── README.md             # This file
```

## 🔧 Configuration Options

### Environment Variables

| Variable | Description | Required | Default |
|----------|-------------|----------|---------|
| `AUTH_MICROSOFT_ENTRA_ID_ID` | Azure App Client ID | Yes | - |
| `AUTH_MICROSOFT_ENTRA_ID_SECRET` | Azure App Client Secret | Yes | - |
| `PORT` | Server port | No | 3001 |
| `AZURE_REDIRECT_URI` | OAuth redirect URI | No | `http://localhost:3001/api/mailer/email-servers/microsoft/redirect` |

### Token Storage

Tokens are stored in `tokens.json` with the following structure:

```json
{
  "accessToken": "eyJ0eXAiOiJKV1QiLCJub...",
  "refreshToken": "M.C507_BL2.0.U.-CjuY...",
  "expiresOn": "2025-10-30T15:30:00.000Z",
  "savedAt": "2025-10-30T14:30:00.000Z",
  "account": {
    "username": "user@example.com"
  }
}
``` 
## 🐛 Troubleshooting

### Common Issues

1. **"No valid token available"**
   - Solution: Visit the OAuth endpoint to re-authorize

2. **"Authentication failed"**
   - Check Azure app permissions
   - Verify client ID and secret in `.env`
   - Ensure redirect URI matches Azure configuration

3. **"Token refresh failed"**
   - The refresh token might be expired
   - Re-authorize by visiting the OAuth endpoint

4. **Email sending fails**
   - Check if the authenticated user has permission to send emails
   - Verify recipient email addresses
   - Check Microsoft Graph service status

### Debug Mode

Enable debug logging by adding console.log statements or use the token status endpoint to diagnose issues:

```bash
curl http://localhost:3001/api/token-status
```
