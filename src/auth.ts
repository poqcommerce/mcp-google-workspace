#!/usr/bin/env node

import { OAuth2Client } from 'google-auth-library';
import { createServer } from 'http';
import { URL } from 'url';
import open from 'open';

const REDIRECT_URI = 'http://localhost:3000/oauth/callback';
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

async function authenticate() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    console.error('Error: GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET environment variables are required');
    console.error('Usage: GOOGLE_CLIENT_ID="your-id" GOOGLE_CLIENT_SECRET="your-secret" npm run auth');
    process.exit(1);
  }

  const oauth2Client = new OAuth2Client(clientId, clientSecret, REDIRECT_URI);

  // Generate the url that will be used for the consent dialog
  const authorizeUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent', // Force to get refresh token
  });

  console.log('Starting OAuth flow...');
  console.log('Opening browser for authorization...');

  // Create a local server to handle the callback
  const server = createServer(async (req, res) => {
    if (req.url?.startsWith('/oauth/callback')) {
      const url = new URL(req.url, `http://localhost:3000`);
      const code = url.searchParams.get('code');
      
      if (code) {
        try {
          const { tokens } = await oauth2Client.getToken(code);
          
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`
            <html>
              <body style="font-family: Arial, sans-serif; padding: 40px; text-align: center;">
                <h1 style="color: green;">✅ Authorization Successful!</h1>
                <p>You can close this window and return to your terminal.</p>
                <div style="background: #f5f5f5; padding: 20px; margin: 20px 0; border-radius: 5px;">
                  <h3>Add this to your environment:</h3>
                  <code style="background: white; padding: 10px; display: block; margin: 10px 0; border-radius: 3px;">
                    GOOGLE_REFRESH_TOKEN=${tokens.refresh_token}
                  </code>
                </div>
              </body>
            </html>
          `);
          
          console.log('\n🎉 Authorization successful!');
          console.log('\nAdd this environment variable to your Claude Desktop config:');
          console.log(`\n  "GOOGLE_REFRESH_TOKEN": "${tokens.refresh_token}"\n`);
          console.log('Then restart Claude Desktop.\n');
          
          server.close();
        } catch (error) {
          console.error('Error exchanging code for tokens:', error);
          res.writeHead(500, { 'Content-Type': 'text/plain' });
          res.end('Error during authentication');
          server.close();
        }
      } else {
        console.error('No authorization code received');
        res.writeHead(400, { 'Content-Type': 'text/plain' });
        res.end('No authorization code received');
        server.close();
      }
    } else {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not found');
    }
  });

  server.listen(3000, () => {
    console.log('Local server started on http://localhost:3000');
    console.log('Opening authorization URL in browser...\n');
    
    // Open the authorization URL in the default browser
    open(authorizeUrl).catch(() => {
      console.log('Could not open browser automatically. Please visit this URL manually:');
      console.log(authorizeUrl);
    });
  });

  // Handle process termination
  process.on('SIGINT', () => {
    console.log('\nAuthentication cancelled');
    server.close();
    process.exit(0);
  });
}

authenticate().catch(console.error);
