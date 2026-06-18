import fs from 'node:fs';
import path from 'node:path';
import { createWriteStream, createReadStream } from 'node:fs';
import { pipeline } from 'node:stream/promises';
import { google } from 'googleapis';
import { authenticate } from '@google-cloud/local-auth';

const SCOPES = ['https://www.googleapis.com/auth/drive'];

let driveApi = null;

/** OAuth flow: oauth.json + token.json (same as lsmw_equipment_v0). */
export async function initGoogleDrive(oauthJsonPath) {
  const oauthPath = path.resolve(oauthJsonPath);
  const tokenPath = path.join(path.dirname(oauthPath), 'token.json');

  const credentials = JSON.parse(fs.readFileSync(oauthPath, 'utf8'));
  const { client_id, client_secret, redirect_uris } = credentials.installed;

  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  if (fs.existsSync(tokenPath)) {
    let token = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
    auth.setCredentials(token);

    const stale =
      !token.access_token || (token.expiry_date && token.expiry_date < Date.now() + 60_000);

    if (stale && token.refresh_token) {
      console.log('[info] Refreshing Google token...');
      const { credentials: refreshed } = await auth.refreshAccessToken();
      token = {
        ...token,
        ...refreshed,
        refresh_token: refreshed.refresh_token || token.refresh_token,
      };
      fs.writeFileSync(tokenPath, JSON.stringify(token));
      auth.setCredentials(token);
      console.log('[ok] Token refreshed.');
    }
  } else {
    const newAuth = await authenticate({
      keyfilePath: oauthPath,
      scopes: SCOPES,
    });
    const tok = newAuth.credentials;
    fs.writeFileSync(tokenPath, JSON.stringify(tok));
    auth.setCredentials(tok);
    console.log('[ok] Token saved to token.json');
  }

  driveApi = google.drive({ version: 'v3', auth });
}

export function getDrive() {
  if (!driveApi) throw new Error('initGoogleDrive() must be called first');
  return driveApi;
}

function escapeDriveName(name) {
  return String(name).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

/** Direct children only (paginated); folders and files included. */
export async function listFilesInFolder(folderId) {
  const drive = getDrive();
  const out = [];
  let pageToken;
  do {
    const res = await drive.files.list({
      q: `'${folderId}' in parents and trashed=false`,
      fields:
        'nextPageToken, files(id, name, mimeType)',
      pageSize: 1000,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      pageToken,
    });
    out.push(...(res.data.files || []));
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return out;
}

/** Recursively list all non-folder files under a Drive folder. */
export async function listFilesRecursively(folderId, relativePrefix = '') {
  const drive = getDrive();
  const out = [];

  async function walk(id, prefix) {
    let pageToken;
    do {
      const res = await drive.files.list({
        q: `'${id}' in parents and trashed=false`,
        fields: 'nextPageToken, files(id, name, mimeType, owners(emailAddress))',
        pageSize: 1000,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        pageToken,
      });

      const files = res.data.files || [];
      for (const f of files) {
        const rel = prefix ? `${prefix}/${f.name}` : f.name;
        if (f.mimeType === 'application/vnd.google-apps.folder') {
          await walk(f.id, rel);
        } else {
          out.push({
            id: f.id,
            name: f.name,
            mimeType: f.mimeType,
            path: rel,
            owners: f.owners || [],
          });
        }
      }
      pageToken = res.data.nextPageToken;
    } while (pageToken);
  }

  await walk(folderId, relativePrefix || '');
  return out;
}

export async function downloadFile(fileId, destPath) {
  const drive = getDrive();
  fs.mkdirSync(path.dirname(destPath), { recursive: true });
  const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'stream' });
  await pipeline(res.data, createWriteStream(destPath));
}

export async function getOrCreateFolder(name, parentId) {
  const drive = getDrive();
  const safe = escapeDriveName(name);
  const q = `name='${safe}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const existing = await drive.files.list({
    q,
    fields: 'files(id, name)',
    pageSize: 10,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  const found = existing.data.files?.[0];
  if (found) return { id: found.id, name: found.name };

  const created = await drive.files.create({
    resource: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId],
    },
    fields: 'id, name',
    supportsAllDrives: true,
  });
  return { id: created.data.id, name: created.data.name };
}

export async function uploadFile(localPath, fileName, parentId, mimeType) {
  const drive = getDrive();
  await drive.files.create({
    resource: {
      name: fileName,
      parents: [parentId],
    },
    media: {
      mimeType:
        mimeType ||
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: createReadStream(localPath),
    },
    fields: 'id',
    supportsAllDrives: true,
  });
}

/** Like lsmw_equipment_v0: remove same-named file(s) under parent first, then upload. */
export async function uploadXlsxReplacing(localPath, fileName, parentId, mimeType) {
  const drive = getDrive();
  const safe = escapeDriveName(fileName);
  const q = `name='${safe}' and '${parentId}' in parents and trashed=false`;
  let pageToken;
  do {
    const existing = await drive.files.list({
      q,
      fields: 'nextPageToken, files(id, name)',
      pageSize: 100,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      pageToken,
    });
    for (const f of existing.data.files || []) {
      await drive.files.delete({ fileId: f.id });
    }
    pageToken = existing.data.nextPageToken;
  } while (pageToken);

  await uploadFile(localPath, fileName, parentId, mimeType);
}

export function extractFolderIdFromUrl(url) {
  if (!url || typeof url !== 'string') return null;
  const trimmed = url.trim();
  const m =
    trimmed.match(/\/folders\/([a-zA-Z0-9_-]+)/) ||
    trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{15,}$/.test(trimmed)) return trimmed;
  return null;
}

export function extractFileIdFromAnyUrl(raw) {
  if (!raw || typeof raw !== 'string') return null;
  const s = raw.trim();
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  const m =
    s.match(/\/d\/([a-zA-Z0-9_-]+)/) ||
    s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
    s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/) ||
    s.match(/[?&#]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}
