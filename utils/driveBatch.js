import path from 'node:path';
import { createReadStream } from 'node:fs';
import { getOrCreateFolder, getDrive } from './googleDrive.js';
import { runWithConcurrency } from './concurrencyHelpers.js';

/** Timestamp string for Jakarta (folder names). */
export function getJakartaTimestamp() {
  const fmt = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  });
  const parts = Object.fromEntries(
    fmt.formatToParts(new Date()).map((p) => [p.type, p.value]),
  );
  return `${parts.year}${parts.month}${parts.day}_${parts.hour}${parts.minute}${parts.second}`;
}

/** Pre-create first path segment folders (REGIONAL …) under output root. */
export async function ensureRegionalFolders(outputFolderId, relativePaths, { logInfo, logWarn }) {
  const regionNames = new Set();
  for (const rp of relativePaths) {
    if (!rp) continue;
    const parts = String(rp).replace(/\\/g, '/').split('/').filter(Boolean);
    if (parts.length >= 2) regionNames.add(parts[0]);
  }

  const regionalFolders = new Map();
  for (const name of regionNames) {
    try {
      const folder = await getOrCreateFolder(name, outputFolderId);
      regionalFolders.set(name, folder.id);
      logInfo?.(`Regional folder ready: ${name}`);
    } catch (e) {
      logWarn?.(`Could not ensure regional folder "${name}": ${e.message}`);
    }
  }
  return regionalFolders;
}

async function uploadOneExcel(localPath, fileName, parentId) {
  const drive = getDrive();
  await drive.files.create(
    {
      resource: {
        name: fileName,
        parents: [parentId],
      },
      media: {
        mimeType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: createReadStream(localPath),
      },
      fields: 'id',
      supportsAllDrives: true,
    },
    {},
  );
}

/**
 * Upload each item under outputFolderId or its regional subfolder (from relativePath).
 * @param {Array<{ localPath: string, fileName: string, relativePath?: string }>} items
 */
export async function uploadToDriveWithRegionalFolders(items, outputFolderId, opts) {
  const { concurrency = 3, regionalFolders, logSuccess, logWarn, logError } = opts;

  return runWithConcurrency(items, concurrency, async (item) => {
    const { localPath, fileName, relativePath } = item;
    const parts = String(relativePath || path.basename(localPath))
      .replace(/\\/g, '/')
      .split('/')
      .filter(Boolean);

    let parentId = outputFolderId;
    if (parts.length >= 2 && regionalFolders?.has(parts[0])) {
      parentId = regionalFolders.get(parts[0]);
    }

    try {
      await uploadOneExcel(localPath, fileName, parentId);
      logSuccess?.(`Uploaded: ${fileName}`);
      return { success: true, fileName };
    } catch (e) {
      logError?.(`Upload failed ${fileName}: ${e.message}`);
      logWarn?.(e.stack);
      return { success: false, fileName, error: e.message };
    }
  });
}
