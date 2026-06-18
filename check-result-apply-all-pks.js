/**
 * Compare FUNCTIONAL LOCATION AFTER between a master spreadsheet and
 * spreadsheets under a Google Drive --source folder or a single Drive file
 * (--single-file). Writes unavailable.json
 * with master-only and source-only rows.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';

import {
  initGoogleDrive,
  getDrive,
  extractFolderIdFromUrl,
  extractFileIdFromAnyUrl,
  listFilesRecursively,
  downloadFile,
} from './utils/googleDrive.js';
import { filterDriveSourceExcelFiles } from './utils/driveSourceFileFilter.js';
import { createTempDir, cleanupTempDir } from './utils/concurrencyHelpers.js';
import {
  findHeadersExcelJs,
  excelJsValueToPlainText,
  normalizeHeader,
  stripFormulasFromWorksheet,
  toCleanString,
} from './utils/excelHelpers.js';
import { ensureDir } from './fileSystem.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const OAUTH_JSON = path.join(__dirname, 'oauth.json');

const DEFAULT_MASTER_URL =
  'https://docs.google.com/spreadsheets/d/1u0H-GZIU0pYmrodA27uNC_zKtD8AmxLM/edit?usp=sharing';

const FUNCLOC_HEADERS = ['FUNCTIONAL LOCATION AFTER'];

const GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const printHelp = () => {
  console.log(`
Check apply-all-pks result (FUNCTIONAL LOCATION AFTER, master vs Drive source)

Usage:
  node check-result-apply-all-pks.js --source <drive folder url> [options]
  node check-result-apply-all-pks.js --single-file --source <spreadsheet url> [options]

Options:
  --source <url>         Drive folder or spreadsheet URL/id (required)
  --single-file          Treat --source as one spreadsheet file, not a folder
  --master <id|url>      Master spreadsheet (default: configured MASTER_DATA)
  --out <file>           Output JSON path (default: ./unavailable.json)
  --help, -h             Show help
`);
};

const parseArgs = () => {
  const argv = process.argv.slice(2);
  const out = {
    source: null,
    singleFile: false,
    masterUrl: DEFAULT_MASTER_URL,
    outPath: path.join(process.cwd(), 'unavailable.json'),
    help: false,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--single-file') {
      out.singleFile = true;
    } else if (a === '--source' && argv[i + 1]) {
      out.source = argv[i + 1];
      i += 1;
    } else if (a === '--master' && argv[i + 1]) {
      out.masterUrl = argv[i + 1];
      i += 1;
    } else if (a === '--out' && argv[i + 1]) {
      const raw = String(argv[i + 1]).trim();
      out.outPath = path.isAbsolute(raw) ? raw : path.join(process.cwd(), raw);
      i += 1;
    } else if (a === '--help' || a === '-h') {
      out.help = true;
    }
  }
  return out;
};

const isRegionalPath = (filePath) => {
  const first = String(filePath || '').split('/').find(Boolean);
  return /^REGIONAL/i.test(String(first || '').trim());
};

const normalizeFunclocKey = (value) => toCleanString(excelJsValueToPlainText(value)).toUpperCase();

const resolveFunclocCol = (colMap) => {
  for (const h of FUNCLOC_HEADERS) {
    const c = colMap.get(normalizeHeader(h));
    if (c) return c;
  }
  return undefined;
};

const spreadsheetUrlForFile = (file) => {
  if (file.mimeType === GOOGLE_SHEET_MIME) {
    return `https://docs.google.com/spreadsheets/d/${file.id}/edit`;
  }
  return `https://drive.google.com/file/d/${file.id}/view`;
};

const isSpreadsheetMime = (mimeType) =>
  mimeType === GOOGLE_SHEET_MIME || mimeType === XLSX_MIME;

const downloadSpreadsheetToPath = async (fileId, mimeType, destPath) => {
  const drive = getDrive();
  fs.mkdirSync(path.dirname(destPath), { recursive: true });
  if (mimeType === GOOGLE_SHEET_MIME) {
    const res = await drive.files.export(
      {
        fileId,
        mimeType: XLSX_MIME,
      },
      { responseType: 'stream' },
    );
    await new Promise((resolve, reject) => {
      const ws = fs.createWriteStream(destPath);
      res.data.pipe(ws).on('finish', resolve).on('error', reject);
    });
    return;
  }
  await downloadFile(fileId, destPath);
};

const collectFunclocRowsFromWorkbook = async (localPath, url) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(localPath);
  const rows = [];

  for (const worksheet of workbook.worksheets) {
    stripFormulasFromWorksheet(worksheet);
    const found = findHeadersExcelJs(worksheet, FUNCLOC_HEADERS, {
      maxScanRows: 10,
      minSimilarity: 0.7,
    });
    if (!found) continue;

    const { headerRowNumber, colMap } = found;
    const funclocCol = resolveFunclocCol(colMap);
    if (!funclocCol) continue;

    for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r += 1) {
      const raw = excelJsValueToPlainText(worksheet.getRow(r).getCell(funclocCol).value).trim();
      if (!raw) continue;
      const key = normalizeFunclocKey(raw);
      if (!key) continue;
      rows.push({
        url,
        sheet: worksheet.name,
        row: r,
        func_loc_after: raw,
        key,
      });
    }
  }

  return rows;
};

const loadMasterFunclocRows = async ({ masterUrl, tempDir }) => {
  const masterId = extractFileIdFromAnyUrl(masterUrl);
  if (!masterId) throw new Error(`Cannot extract master file id from URL: ${masterUrl}`);

  const localPath = path.join(tempDir, 'master.xlsx');
  const drive = getDrive();
  const meta = await drive.files.get({
    fileId: masterId,
    fields: 'id,name,mimeType',
    supportsAllDrives: true,
  });
  await downloadSpreadsheetToPath(masterId, meta.data.mimeType, localPath);
  const url = `https://docs.google.com/spreadsheets/d/${masterId}/edit`;
  console.log(`[ok] Master loaded: ${meta.data.name}`);
  return collectFunclocRowsFromWorkbook(localPath, url);
};

const loadSourceFunclocRows = async ({ sourceFolderId, tempDir }) => {
  const allFiles = await listFilesRecursively(sourceFolderId);
  const { driveFiles: xlsxFiles } = filterDriveSourceExcelFiles(allFiles, {
    allowedOwnerEmails: [],
  });
  const sheetFiles = allFiles.filter((f) => f.mimeType === GOOGLE_SHEET_MIME);

  const merged = [...xlsxFiles, ...sheetFiles];
  const regionalOnly = merged.filter((f) => isRegionalPath(f.path));
  const filtered = regionalOnly.filter((f) => {
    const segs = String(f.path || '').split('/').filter(Boolean);
    return !segs.some((s) => /^APPLIED_/i.test(s) || /^Completion_/i.test(s));
  });

  if (regionalOnly.length !== merged.length) {
    console.log(
      `[info] Skipped ${merged.length - regionalOnly.length} file(s) outside REGIONAL* subfolders`,
    );
  }

  const rows = [];
  for (const file of filtered) {
    const label = file.path || file.name;
    const safeName = String(label).replace(/[\\/]+/g, '_');
    const localPath = path.join(tempDir, 'source', safeName);
    ensureDir(path.dirname(localPath));
    const url = spreadsheetUrlForFile(file);
    console.log(`[info] Reading source: ${label}`);
    await downloadSpreadsheetToPath(file.id, file.mimeType, localPath);
    const fileRows = await collectFunclocRowsFromWorkbook(localPath, url);
    rows.push(...fileRows);
  }
  return { rows, fileCount: filtered.length };
};

const loadSingleSourceFunclocRows = async ({ sourceUrl, tempDir }) => {
  const fileId = extractFileIdFromAnyUrl(sourceUrl);
  if (!fileId) throw new Error(`Cannot extract source file id from URL: ${sourceUrl}`);

  const localPath = path.join(tempDir, 'source.xlsx');
  const drive = getDrive();
  const meta = await drive.files.get({
    fileId,
    fields: 'id,name,mimeType',
    supportsAllDrives: true,
  });
  const { id, name, mimeType } = meta.data;
  if (!isSpreadsheetMime(mimeType)) {
    throw new Error(`Source file is not a spreadsheet: ${name} (${mimeType})`);
  }

  await downloadSpreadsheetToPath(fileId, mimeType, localPath);
  const url = spreadsheetUrlForFile({ id, mimeType });
  console.log(`[ok] Source loaded: ${name}`);
  const rows = await collectFunclocRowsFromWorkbook(localPath, url);
  return { rows, fileCount: 1 };
};

const buildUnavailable = (masterRows, sourceRows) => {
  const sourceKeys = new Set(sourceRows.map((r) => r.key));
  const masterKeys = new Set(masterRows.map((r) => r.key));
  const out = [];

  for (const row of masterRows) {
    if (sourceKeys.has(row.key)) continue;
    out.push({
      url: row.url,
      sheet: row.sheet,
      row: row.row,
      func_loc_after: row.func_loc_after,
      type: 'master-only',
    });
  }

  for (const row of sourceRows) {
    if (masterKeys.has(row.key)) continue;
    out.push({
      url: row.url,
      sheet: row.sheet,
      row: row.row,
      func_loc_after: row.func_loc_after,
      type: 'source-only',
    });
  }

  return out;
};

const main = async () => {
  const args = parseArgs();
  if (args.help) {
    printHelp();
    process.exit(0);
  }
  if (!args.source) {
    console.error('[err] --source is required');
    printHelp();
    process.exit(1);
  }

  if (args.singleFile) {
    if (!extractFileIdFromAnyUrl(args.source)) {
      console.error(`[err] Invalid --source file URL or id: ${args.source}`);
      process.exit(1);
    }
  } else if (!extractFolderIdFromUrl(args.source)) {
    console.error(`[err] Invalid --source folder URL: ${args.source}`);
    process.exit(1);
  }

  await initGoogleDrive(OAUTH_JSON);
  console.log('[ok] Google Drive auth ready');

  const tempDir = createTempDir();
  try {
    const masterRows = await loadMasterFunclocRows({
      masterUrl: args.masterUrl,
      tempDir,
    });
    const { rows: sourceRows, fileCount } = args.singleFile
      ? await loadSingleSourceFunclocRows({ sourceUrl: args.source, tempDir })
      : await loadSourceFunclocRows({
          sourceFolderId: extractFolderIdFromUrl(args.source),
          tempDir,
        });

    const unavailable = buildUnavailable(masterRows, sourceRows);
    ensureDir(path.dirname(args.outPath));
    fs.writeFileSync(args.outPath, JSON.stringify(unavailable, null, 2), 'utf8');

    const masterOnly = unavailable.filter((r) => r.type === 'master-only').length;
    const sourceOnly = unavailable.filter((r) => r.type === 'source-only').length;
    console.log(
      `[ok] Compared master rows=${masterRows.length}, source files=${fileCount}, source rows=${sourceRows.length}`,
    );
    console.log(`[ok] unavailable.json: ${args.outPath} (master-only=${masterOnly}, source-only=${sourceOnly})`);
  } finally {
    cleanupTempDir(tempDir);
  }
};

main().catch((err) => {
  console.error('Failed:', err.message);
  process.exit(1);
});
