const { google } = require('googleapis');
const http = require('http');
const url = require('url');
const fs = require('fs');
const path = require('path');

// OAuth2 credentials for Google Apps Script API
// You need to create these in Google Cloud Console
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const REDIRECT_URI = 'http://localhost:3000/callback';
const SCOPES = ['https://www.googleapis.com/auth/script.projects'];

const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

async function getAuthToken() {
  if (!CLIENT_ID || !CLIENT_SECRET) {
    console.log('');
    console.log('=== Google OAuth2 Token Generator ===');
    console.log('');
    console.log('Step 1: Go to Google Cloud Console');
    console.log('  https://console.cloud.google.com');
    console.log('');
    console.log('Step 2: Create or select a project');
    console.log('');
    console.log('Step 3: Enable Apps Script API');
    console.log('  https://console.cloud.google.com/apis/library/script.googleapis.com');
    console.log('');
    console.log('Step 4: Create OAuth2 credentials');
    console.log('  - Go to APIs & Services > Credentials');
    console.log('  - Click "Create Credentials" > "OAuth client ID"');
    console.log('  - Application type: "Desktop app"');
    console.log('  - Copy the Client ID and Client Secret');
    console.log('');
    console.log('Step 5: Run this script with credentials:');
    console.log('  $env:GOOGLE_CLIENT_ID="your-client-id"');
    console.log('  $env:GOOGLE_CLIENT_SECRET="your-client-secret"');
    console.log('  node get-token.js');
    console.log('');
    return;
  }

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'
  });

  console.log('');
  console.log('Opening browser for authentication...');
  console.log('If browser doesn\'t open, visit this URL:');
  console.log('');
  console.log(authUrl);
  console.log('');

  // Try to open browser
  const { exec } = require('child_process');
  exec(`start "${authUrl}"`, (err) => {
    if (err) console.log('Could not open browser automatically');
  });

  // Start local server to handle callback
  const server = http.createServer(async (req, res) => {
    const query = url.parse(req.url, true).query;
    if (query.code) {
      try {
        const { tokens } = await oauth2Client.getToken(query.code);
        oauth2Client.setCredentials(tokens);

        // Save token in clasp format
        const clasprc = {
          token: tokens,
          oauth2Credentials: {
            installed: {
              client_id: CLIENT_ID,
              client_secret: CLIENT_SECRET,
              redirect_uris: [REDIRECT_URI]
            }
          }
        };

        const rcPath = path.join(require('os').homedir(), '.clasprc.json');
        fs.writeFileSync(rcPath, JSON.stringify(clasprc, null, 2));

        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<h1>Authentication successful!</h1><p>You can close this window and return to the terminal.</p>');

        console.log('');
        console.log('=== Authentication Successful ===');
        console.log('');
        console.log('Token saved to: ' + rcPath);
        console.log('');
        console.log('You can now run: clasp push');
        console.log('');

        server.close();
        process.exit(0);
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'text/html' });
        res.end('<h1>Error</h1><p>' + error.message + '</p>');
        console.error('Error getting token:', error.message);
      }
    }
  });

  server.listen(3000, () => {
    console.log('Listening on http://localhost:3000');
  });
}

getAuthToken();
