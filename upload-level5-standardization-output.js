import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  extractFolderIdFromUrl,
  getOrCreateFolder,
  initGoogleDrive,
  listFilesInFolder,
  uploadFile,
} from './utils/googleDrive.js';
import { runWithConcurrency } from './utils/concurrencyHelpers.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DEFAULT_OUT_DIR = path.join(__dirname, 'Output', 'BahJambi_Level5');
const DEFAULT_CONCURRENCY = 3;
const OAUTH_JSON = path.join(__dirname, 'oauth.json');

const getJakartaStamp = () => {
  const fmt = new Intl.DateTimeFormat('sv-SE', {
    timeZone: 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  });
  const parts = fmt.formatToParts(new Date());
  const map = new Map(parts.map((p) => [p.type, p.value]));
  return `${map.get('year')}${map.get('month')}${map.get('day')}_${map.get('hour')}${map.get('minute')}${map.get('second')}`;
};

const parseArgs = () => {
  const args = process.argv.slice(2);
  let sourceFolderUrl = null;
  let localOutputDir = DEFAULT_OUT_DIR;
  let concurrency = DEFAULT_CONCURRENCY;

  for (let i = 0; i < args.length; i += 1) {
    const a = args[i];
    if (a === '--source-folder-url' && args[i + 1]) {
      sourceFolderUrl = args[i + 1];
      i += 1;
    } else if (a === '--local-output-dir' && args[i + 1]) {
      localOutputDir = args[i + 1];
      i += 1;
    } else if (a === '--concurrency' && args[i + 1]) {
      const n = Number.parseInt(args[i + 1], 10);
      if (!Number.isNaN(n) && n > 0) concurrency = n;
      i += 1;
    }
  }

  if (!sourceFolderUrl) {
    throw new Error('Missing required flag: --source-folder-url <driveFolderUrl>');
  }

  return {
    sourceFolderUrl,
    localOutputDir: path.isAbsolute(localOutputDir)
      ? localOutputDir
      : path.join(process.cwd(), localOutputDir),
    concurrency,
  };
};

const collectFilesRecursive = (rootDir) => {
  const files = [];
  const walk = (dir) => {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath);
      } else if (entry.isFile()) {
        const relativePath = path.relative(rootDir, fullPath);
        files.push({ fullPath, relativePath });
      }
    }
  };
  walk(rootDir);
  return files;
};

const getMimeType = (filePath) => {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xlsx')
    return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (ext === '.html' || ext === '.htm') return 'text/html';
  if (ext === '.json') return 'application/json';
  if (ext === '.csv') return 'text/csv';
  if (ext === '.txt' || ext === '.log') return 'text/plain';
  return 'application/octet-stream';
};

const getUniqueFileName = (fileName, existingNames) => {
  if (!existingNames.has(fileName)) {
    existingNames.add(fileName);
    return { finalName: fileName, renamed: false };
  }

  const ext = path.extname(fileName);
  const base = path.basename(fileName, ext);
  let index = 1;
  while (true) {
    const candidate = `${base} (${index})${ext}`;
    if (!existingNames.has(candidate)) {
      existingNames.add(candidate);
      return { finalName: candidate, renamed: true };
    }
    index += 1;
  }
};

const ensureFolderPath = async (rootFolderId, relativeDir, folderCache) => {
  const normalized = String(relativeDir || '')
    .split(path.sep)
    .filter(Boolean)
    .join('/');
  if (!normalized || normalized === '.') return rootFolderId;
  if (folderCache.has(normalized)) return folderCache.get(normalized);

  const segments = normalized.split('/');
  let parentId = rootFolderId;
  let built = '';
  for (const seg of segments) {
    built = built ? `${built}/${seg}` : seg;
    if (folderCache.has(built)) {
      parentId = folderCache.get(built);
      continue;
    }
    const folder = await getOrCreateFolder(seg, parentId);
    folderCache.set(built, folder.id);
    parentId = folder.id;
  }
  return parentId;
};

const main = async () => {
  const { sourceFolderUrl, localOutputDir, concurrency } = parseArgs();
  await initGoogleDrive(OAUTH_JSON);
  if (!fs.existsSync(localOutputDir)) {
    throw new Error(`Local output directory not found: ${localOutputDir}`);
  }

  const folderId = extractFolderIdFromUrl(sourceFolderUrl);
  const targetFolderName = `Level 5 Standardization_${getJakartaStamp()}`;
  const outputFolder = await getOrCreateFolder(targetFolderName, folderId);

  console.log(`Upload source dir : ${localOutputDir}`);
  console.log(`Drive source      : ${sourceFolderUrl}`);
  console.log(`Target folder     : ${outputFolder.name} (${outputFolder.id})`);

  const localFiles = collectFilesRecursive(localOutputDir);
  if (!localFiles.length) {
    console.log('No local files to upload.');
    return;
  }
  const folderCache = new Map();
  const usedNamesByFolder = new Map();

  const getUsedNamesForFolder = async (folderId) => {
    if (usedNamesByFolder.has(folderId)) return usedNamesByFolder.get(folderId);
    const existing = await listFilesInFolder(folderId);
    const used = new Set(
      existing
        .filter((f) => f.mimeType !== 'application/vnd.google-apps.folder')
        .map((f) => String(f.name || '').trim())
        .filter(Boolean),
    );
    usedNamesByFolder.set(folderId, used);
    return used;
  };

  let uploadedCount = 0;
  let renamedCount = 0;
  let failedCount = 0;

  await runWithConcurrency(localFiles, concurrency, async (fileItem) => {
    const localPath = fileItem.fullPath;
    const originalName = path.basename(localPath);
    const relativeDir = path.dirname(fileItem.relativePath);

    try {
      const targetFolderId = await ensureFolderPath(outputFolder.id, relativeDir, folderCache);
      const usedNames = await getUsedNamesForFolder(targetFolderId);
      const { finalName, renamed } = getUniqueFileName(originalName, usedNames);
      await uploadFile(localPath, finalName, targetFolderId, getMimeType(localPath));
      uploadedCount += 1;
      if (renamed) renamedCount += 1;
      console.log(
        `Uploaded: ${fileItem.relativePath}${renamed ? ` -> ${path.join(relativeDir, finalName)}` : ''}`,
      );
    } catch (err) {
      failedCount += 1;
      console.error(`Failed upload: ${fileItem.relativePath} (${err.message})`);
    }
  });

  console.log('\nUpload summary');
  console.log(`- Total files : ${localFiles.length}`);
  console.log(`- Uploaded    : ${uploadedCount}`);
  console.log(`- Renamed     : ${renamedCount}`);
  console.log(`- Failed      : ${failedCount}`);
  if (outputFolder.webViewLink) {
    console.log(`- Drive folder: ${outputFolder.webViewLink}`);
  }

  if (failedCount > 0) {
    process.exitCode = 1;
  }
};

main().catch((err) => {
  console.error(`Failed: ${err.message}`);
  process.exit(1);
});
