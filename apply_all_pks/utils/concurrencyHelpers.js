import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { google } from 'googleapis';
import { authorize as authorizeOAuth, isOAuthConfigured } from './oauthAuth.js';
import { logWarn } from './logger.js';

/**
 * Run processing with concurrency control
 */
export const runWithConcurrency = async (items, limit, worker) => {
  if (limit <= 0) {
    return Promise.all(items.map(worker));
  }

  const results = new Array(items.length);
  let index = 0;

  const runNext = async () => {
    const currentIndex = index;
    index += 1;
    if (currentIndex >= items.length) return;

    results[currentIndex] = await worker(items[currentIndex]);
    return runNext();
  };

  const runners = Array.from({ length: Math.min(limit, items.length) }, () => runNext());
  await Promise.all(runners);
  return results;
};

/**
 * Upload any file type to Google Drive
 * @param {string} filePath - Local file path
 * @param {string} folderId - Google Drive folder ID
 * @param {string} fileName - File name
 * @param {string} mimeType - MIME type
 * @returns {Promise<Object>} Uploaded file metadata
 */
export const uploadAnyFile = async (filePath, folderId, fileName, mimeType) => {
  let auth;
  if (isOAuthConfigured()) {
    auth = await authorizeOAuth();
  } else {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);
    const credentialsPath = path.join(__dirname, '..', 'auth', 'credentials.json');
    if (!fs.existsSync(credentialsPath)) {
      throw new Error('No authentication credentials found');
    }
    const credentials = JSON.parse(fs.readFileSync(credentialsPath, 'utf8'));
    const { GoogleAuth } = google.auth;
    auth = new GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
  }

  const drive = google.drive({ version: 'v3', auth });

  const fileMetadata = {
    name: fileName,
    parents: [folderId],
  };

  const media = {
    mimeType,
    body: fs.createReadStream(filePath),
  };

  const response = await drive.files.create({
    requestBody: fileMetadata,
    media,
    fields: 'id, name, webViewLink',
    supportsAllDrives: true,
  });

  return response.data;
};

// Create temporary directory for downloaded files
export const createTempDir = () => {
  const tempDir = path.join(
    os.tmpdir(),
    `full-validation-${Date.now()}-${Math.random().toString(36).substring(7)}`,
  );
  fs.mkdirSync(tempDir, { recursive: true });
  return tempDir;
};

// Clean up temporary directory
export const cleanupTempDir = (tempDir) => {
  try {
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  } catch (err) {
    logWarn(`Failed to cleanup temp directory: ${err.message}`);
  }
};
