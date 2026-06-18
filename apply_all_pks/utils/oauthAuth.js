import { google } from 'googleapis';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import open from 'open';
import readline from 'node:readline';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SCOPES = ['https://www.googleapis.com/auth/drive'];
const TOKEN_PATH = path.join(path.dirname(__dirname), 'auth', 'token.json');
const OAUTH_CREDENTIALS_PATH = path.join(path.dirname(__dirname), 'auth', 'oauth-credentials.json');

/**
 * Load or request authorization for OAuth2
 * @returns {Promise<google.auth.OAuth2Client>}
 */
export const authorize = async () => {
  // Check if OAuth credentials exist
  if (!fs.existsSync(OAUTH_CREDENTIALS_PATH)) {
    throw new Error(
      `OAuth credentials not found: ${OAUTH_CREDENTIALS_PATH}\n` +
        'Please follow the steps in OAUTH_SETUP.md to create OAuth credentials.',
    );
  }

  const credentials = JSON.parse(fs.readFileSync(OAUTH_CREDENTIALS_PATH, 'utf8'));

  // Handle both "installed" (desktop) and "web" credential types
  const creds = credentials.installed || credentials.web || {};
  const { client_secret, client_id, redirect_uris } = creds;

  if (!client_id || !client_secret) {
    throw new Error('Invalid OAuth credentials. Missing client_id or client_secret.');
  }

  // For desktop apps, use localhost redirect URI
  // Note: For "web" type credentials, you need to add this redirect URI in Google Cloud Console:
  // http://localhost:3000/oauth2callback
  let redirectUri = redirect_uris?.[0];

  // If no redirect URI is specified, use localhost (works for desktop apps)
  // For web type, user must add it in Google Cloud Console
  if (!redirectUri) {
    redirectUri = 'http://localhost:3000/oauth2callback';
    if (credentials.web && !redirect_uris?.length) {
      console.warn('\n⚠️  Warning: No redirect URI found in web credentials.');
      console.warn('Please add "http://localhost:3000/oauth2callback" to Authorized redirect URIs');
      console.warn('in Google Cloud Console for your OAuth client.\n');
    }
  }

  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

  // Check if we have previously stored a token
  if (fs.existsSync(TOKEN_PATH)) {
    try {
      const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
      oAuth2Client.setCredentials(token);

      // Verify token is still valid by refreshing if needed
      if (oAuth2Client.isTokenExpiring()) {
        const newToken = await oAuth2Client.refreshAccessToken();
        oAuth2Client.setCredentials(newToken.credentials);
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(newToken.credentials, null, 2));
      }

      return oAuth2Client;
    } catch (err) {
      console.warn('Error loading token, will re-authenticate:', err.message);
    }
  }

  // Get new token
  return getNewToken(oAuth2Client);
};

/**
 * Get and store new token after prompting for user authorization
 * @param {google.auth.OAuth2Client} oAuth2Client The OAuth2 client to get token for
 * @returns {Promise<google.auth.OAuth2Client>}
 */
const getNewToken = async (oAuth2Client) => {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent', // Force consent to get refresh token
  });

  console.log('\n🔐 Authorization required!');
  console.log('Opening browser for authentication...\n');
  console.log('If browser does not open, visit this URL:');
  console.log(authUrl);
  console.log('\n');
  console.log('After authorizing, you will be redirected to a page.');
  console.log('Copy the "code" parameter from the URL (or the entire URL if it shows an error).');
  console.log(
    'If you see "localhost refused to connect", that\'s normal - just copy the code from the URL.\n',
  );

  // Try to open browser
  try {
    await open(authUrl);
  } catch (err) {
    console.warn('Could not open browser automatically:', err.message);
  }

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve, reject) => {
    rl.question('Enter the code from that page (or paste the full URL): ', async (input) => {
      rl.close();

      try {
        // Extract code from URL if user pasted the full URL
        let code = input.trim();

        // If it looks like a URL, extract the code parameter
        if (code.includes('?') || code.includes('code=')) {
          try {
            const url = new URL(code);
            const extractedCode = url.searchParams.get('code');
            if (extractedCode) {
              code = extractedCode;
            } else {
              // Try regex extraction as fallback
              const codeMatch = input.match(/[?&]code=([^&\s]+)/);
              if (codeMatch) {
                code = decodeURIComponent(codeMatch[1]);
              }
            }
          } catch (urlErr) {
            // Not a valid URL, try regex extraction
            const codeMatch = input.match(/[?&]code=([^&\s]+)/);
            if (codeMatch) {
              code = decodeURIComponent(codeMatch[1]);
            }
          }
        }

        if (!code) {
          throw new Error(
            'Could not extract authorization code. Please paste the code or the full URL.',
          );
        }

        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);

        // Store the token for future use
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
        console.log('✅ Token stored to', TOKEN_PATH);

        resolve(oAuth2Client);
      } catch (err) {
        reject(new Error(`Error retrieving access token: ${err.message}`));
      }
    });
  });
};

/**
 * Check if OAuth is configured
 * @returns {boolean}
 */
export const isOAuthConfigured = () => {
  return fs.existsSync(OAUTH_CREDENTIALS_PATH);
};
