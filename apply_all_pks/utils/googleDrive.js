import { google } from 'googleapis';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Readable } from 'node:stream';
import { authorize as authorizeOAuth, isOAuthConfigured } from './oauthAuth.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let driveClient = null;
let serviceAccountEmail = null;
let authMethod = null; // 'oauth' or 'service_account'

/** Ask intermediaries not to serve stale spreadsheet/binary exports (best-effort). */
const DRIVE_FETCH_HEADERS = {
  'Cache-Control': 'no-cache',
  Pragma: 'no-cache',
};

const initializeDrive = async () => {
  if (driveClient) return driveClient;

  // Check if OAuth is configured (preferred for uploads)
  if (isOAuthConfigured()) {
    try {
      authMethod = 'oauth';
      const oauth2Client = await authorizeOAuth();
      driveClient = google.drive({ version: 'v3', auth: oauth2Client });
      return driveClient;
    } catch (err) {
      console.warn('OAuth initialization failed, falling back to service account:', err.message);
    }
  }

  // Fall back to service account
  const credentialsPath = path.join(path.dirname(__dirname), 'auth', 'credentials.json');
  if (!fs.existsSync(credentialsPath)) {
    throw new Error(
      `No authentication method found!\n` +
        `- OAuth: Create oauth-credentials.json (see OAUTH_SETUP.md)\n` +
        `- Service Account: Create credentials.json`,
    );
  }

  authMethod = 'service_account';
  const credentials = JSON.parse(fs.readFileSync(credentialsPath, 'utf8'));
  serviceAccountEmail = credentials.client_email;

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/drive'],
  });

  driveClient = google.drive({ version: 'v3', auth });
  return driveClient;
};

export const getServiceAccountEmail = () => {
  if (authMethod === 'oauth') {
    return null; // OAuth doesn't use service account email
  }
  if (!serviceAccountEmail) {
    // This will throw if not initialized, but that's okay
    return null;
  }
  return serviceAccountEmail;
};

export const getAuthMethod = () => authMethod;

/**
 * List all files in a Google Drive folder
 * @param {string} folderId - Google Drive folder ID
 * @returns {Promise<Array>} Array of file objects with id, name, mimeType
 */
export const listFilesInFolder = async (folderId) => {
  const drive = await initializeDrive();
  const files = [];
  let pageToken = null;

  do {
    const response = await drive.files.list({
      q: `'${folderId}' in parents and trashed=false`,
      fields: 'nextPageToken, files(id, name, mimeType, parents, owners)',
      pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });

    if (response.data.files) {
      files.push(...response.data.files);
    }

    pageToken = response.data.nextPageToken;
  } while (pageToken);

  return files;
};

/**
 * Get current user's email address
 * @returns {Promise<string|null>} User email or null if cannot be determined
 */
export const getCurrentUserEmail = async () => {
  try {
    const drive = await initializeDrive();

    if (authMethod === 'oauth') {
      // For OAuth, try multiple methods to get user email

      // Method 1: Use Drive about API (most reliable for Drive)
      try {
        const about = await drive.about.get({
          fields: 'user',
          supportsAllDrives: true,
        });
        if (about.data && about.data.user && about.data.user.emailAddress) {
          return about.data.user.emailAddress;
        }
      } catch (aboutErr) {
        // Method 2: Try OAuth2 userinfo API
        try {
          const auth = drive.context._options.auth;
          if (auth && auth.credentials) {
            const oauth2 = google.oauth2({ version: 'v2', auth });
            const userInfo = await oauth2.userinfo.get();
            if (userInfo.data && userInfo.data.email) {
              return userInfo.data.email;
            }
          }
        } catch (oauthErr) {
          // Method 3: Try to get from token file (fallback)
          try {
            const tokenPath = path.join(path.dirname(__dirname), 'auth', 'token.json');
            if (fs.existsSync(tokenPath)) {
              const token = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
              // Token might have email in id_token (JWT)
              if (token.id_token) {
                // Simple base64 decode of JWT payload (without verification)
                const parts = token.id_token.split('.');
                if (parts.length === 3) {
                  const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
                  if (payload.email) {
                    return payload.email;
                  }
                }
              }
            }
          } catch (tokenErr) {
            // Ignore token file errors
          }
          console.warn(
            'Could not get user email from OAuth. Error:',
            aboutErr.message || oauthErr?.message,
          );
        }
      }
    } else if (authMethod === 'service_account') {
      // Service account email is already known
      return serviceAccountEmail;
    }
  } catch (err) {
    console.warn('Error getting current user email:', err.message);
  }
  return null;
};

/**
 * Recursively list all files in a folder and its subfolders
 * @param {string} folderId - Google Drive folder ID
 * @param {string} basePath - Base path for building relative paths
 * @param {string|null} excludeOwnerEmail - Email of owner to exclude (null = don't filter)
 * @returns {Promise<Array>} Array of objects with { id, name, path, mimeType, ownerEmail }
 */
export const listFilesRecursively = async (folderId, basePath = '', excludeOwnerEmail = null) => {
  const drive = await initializeDrive();
  const allFiles = [];
  let filteredCount = 0;

  const processFolder = async (currentFolderId, currentPath) => {
    const items = await listFilesInFolder(currentFolderId);

    for (const item of items) {
      const itemPath = currentPath ? `${currentPath}/${item.name}` : item.name;

      if (item.mimeType === 'application/vnd.google-apps.folder') {
        // Recursively process subfolder
        await processFolder(item.id, itemPath);
      } else if (
        item.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        item.name.toLowerCase().endsWith('.xlsx')
      ) {
        // Get owner email
        const ownerEmail =
          item.owners && item.owners.length > 0 ? item.owners[0].emailAddress : null;

        // Skip files owned by the current user if excludeOwnerEmail is set
        if (excludeOwnerEmail && ownerEmail === excludeOwnerEmail) {
          filteredCount += 1;
          continue;
        }

        // It's an Excel file
        allFiles.push({
          id: item.id,
          name: item.name,
          path: itemPath,
          mimeType: item.mimeType,
          ownerEmail,
        });
      }
    }
  };

  await processFolder(folderId, basePath);

  // Store filtered count for logging
  if (excludeOwnerEmail && filteredCount > 0) {
    listFilesRecursively._lastFilteredCount = filteredCount;
  }

  return allFiles;
};

/**
 * Download a file from Google Drive
 * @param {string} fileId - Google Drive file ID
 * @param {string} outputPath - Local path to save the file
 * @returns {Promise<string>} Path to the downloaded file
 */
export const downloadFile = async (fileId, outputPath) => {
  const drive = await initializeDrive();

  // Ensure output directory exists
  const dir = path.dirname(outputPath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  // Check if file is a Google Sheets file (needs export)
  const fileMetadata = await drive.files.get({ fileId, fields: 'mimeType, name' });
  const mimeType = fileMetadata.data.mimeType;

  let response;
  if (mimeType === 'application/vnd.google-apps.spreadsheet') {
    // Export Google Sheets as Excel
    response = await drive.files.export(
      {
        fileId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      { responseType: 'stream', headers: DRIVE_FETCH_HEADERS },
    );
  } else {
    // Download regular file
    response = await drive.files.get(
      {
        fileId,
        alt: 'media',
      },
      { responseType: 'stream', headers: DRIVE_FETCH_HEADERS },
    );
  }

  return new Promise((resolve, reject) => {
    const fileStream = fs.createWriteStream(outputPath);
    let settled = false;
    const fail = (err) => {
      if (settled) return;
      settled = true;
      reject(err);
    };
    const succeed = () => {
      if (settled) return;
      settled = true;
      resolve(outputPath);
    };

    fileStream.on('finish', succeed);
    fileStream.on('error', fail);
    response.data.on('error', fail);
    response.data.pipe(fileStream);
  });
};

/**
 * Upload a file to Google Drive with retry logic
 * @param {string} filePath - Local file path
 * @param {string} folderId - Google Drive folder ID to upload to
 * @param {string} fileName - Optional custom file name (defaults to basename of filePath)
 * @param {number} maxRetries - Maximum number of retry attempts (default: 3)
 * @returns {Promise<Object>} Uploaded file metadata
 */
export const uploadFile = async (filePath, folderId, fileName = null, maxRetries = 3) => {
  const drive = await initializeDrive();
  const name = fileName || path.basename(filePath);

  const fileMetadata = {
    name,
    parents: [folderId],
  };

  // Retry logic with exponential backoff
  let lastError;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      // Create a new read stream for each attempt (streams can only be read once)
      const media = {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: fs.createReadStream(filePath),
      };

      const response = await drive.files.create({
        requestBody: fileMetadata,
        media,
        fields: 'id, name, webViewLink',
        supportsAllDrives: true,
      });

      return response.data;
    } catch (error) {
      lastError = error;

      // Check if error is retryable
      const isRetryable =
        error.code === 'ECONNRESET' ||
        error.code === 'ETIMEDOUT' ||
        error.code === 'ENOTFOUND' ||
        error.code === 'ECONNREFUSED' ||
        (error.response && error.response.status >= 500) || // Server errors
        (error.response && error.response.status === 429); // Rate limit

      if (!isRetryable || attempt === maxRetries) {
        throw error;
      }

      // Exponential backoff: wait 2^attempt seconds (with jitter)
      const delay = Math.min(1000 * Math.pow(2, attempt) + Math.random() * 1000, 30000);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }

  throw lastError;
};

/**
 * Create a folder in Google Drive
 * @param {string} folderName - Name of the folder
 * @param {string} parentFolderId - Parent folder ID (optional)
 * @returns {Promise<Object>} Created folder metadata
 */
export const createFolder = async (folderName, parentFolderId = null) => {
  const drive = await initializeDrive();

  const fileMetadata = {
    name: folderName,
    mimeType: 'application/vnd.google-apps.folder',
  };

  if (parentFolderId) {
    fileMetadata.parents = [parentFolderId];
  }

  const response = await drive.files.create({
    requestBody: fileMetadata,
    fields: 'id, name, webViewLink',
    supportsAllDrives: true,
  });

  return response.data;
};

/**
 * Get or create a folder by name in a parent folder
 * @param {string} folderName - Name of the folder
 * @param {string} parentFolderId - Parent folder ID
 * @returns {Promise<Object>} Folder metadata with id, name, webViewLink
 */
export const getOrCreateFolder = async (folderName, parentFolderId) => {
  const drive = await initializeDrive();

  // First, try to find existing folder
  const query = `'${parentFolderId}' in parents and name='${folderName.replace(/'/g, "\\'")}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const response = await drive.files.list({
    q: query,
    fields: 'files(id, name, webViewLink)',
    spaces: 'drive',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });

  if (response.data.files && response.data.files.length > 0) {
    return response.data.files[0];
  }

  // Folder doesn't exist, create it
  return createFolder(folderName, parentFolderId);
};

/**
 * Check if a folder is in a Shared Drive
 * @param {string} folderId - Google Drive folder ID
 * @returns {Promise<Object>} Object with { isSharedDrive: boolean, driveId?: string }
 */
export const checkIfSharedDrive = async (folderId) => {
  const drive = await initializeDrive();

  try {
    const response = await drive.files.get({
      fileId: folderId,
      fields: 'driveId, capabilities',
      supportsAllDrives: true,
    });

    // If driveId exists, it's in a Shared Drive
    if (response.data.driveId) {
      return { isSharedDrive: true, driveId: response.data.driveId };
    }

    return { isSharedDrive: false };
  } catch (err) {
    // If we get an error, try to check parent folders
    try {
      const response = await drive.files.get({
        fileId: folderId,
        fields: 'parents',
        supportsAllDrives: true,
      });

      if (response.data.parents && response.data.parents.length > 0) {
        // Check parent folder
        return checkIfSharedDrive(response.data.parents[0]);
      }
    } catch (parentErr) {
      // Ignore and return false
    }

    return { isSharedDrive: false };
  }
};

/**
 * Get folder ID from a Google Drive URL
 * @param {string} url - Google Drive folder URL
 * @returns {string} Folder ID
 */
export const extractFolderIdFromUrl = (url) => {
  const match = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!match) {
    throw new Error(`Invalid Google Drive folder URL: ${url}`);
  }
  return match[1];
};

/**
 * Extract a Google Drive file ID from a variety of URL formats or a bare ID.
 * Supports:
 * - https://docs.google.com/spreadsheets/d/<ID>/...
 * - https://drive.google.com/file/d/<ID>/view
 * - https://drive.google.com/open?id=<ID>
 * - bare IDs (alphanumeric + _ or -)
 */
export const extractFileIdFromAnyUrl = (urlOrId) => {
  if (!urlOrId) return null;
  const value = String(urlOrId).trim();

  const sheetsMatch = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (sheetsMatch) return sheetsMatch[1];

  const fileMatch = value.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) return fileMatch[1];

  const openMatch = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (openMatch) return openMatch[1];

  if (/^[a-zA-Z0-9_-]+$/.test(value)) return value;

  return null;
};
