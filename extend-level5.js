/**
 * Filter FUNCTIONAL LOCATION AFTER rows with exactly 4 or 5 hyphen segments from
 * Google Drive source (--source folder or --single-file) and write only
 * FUNCTIONAL LOCATION AFTER + FUNCTLOC DESC. AFTER LEVEL 1,2,3 to extend-master.xlsx.
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
} from './utils/excelHelpers.js';
import { ensureDir } from './fileSystem.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const OAUTH_JSON = path.join(__dirname, 'oauth.json');
const DEFAULT_OUT_PATH = path.join(__dirname, 'Output', 'extend', 'extend-master.xlsx');

const FUNCLOC_HEADER = 'FUNCTIONAL LOCATION AFTER';
const DESC_HEADERS = [
  'FUNCTLOC DESC. AFTER LEVEL 1,2,3',
  'FUNCTLOC DESC AFTRER LEVEL 1,2,3',
];

const GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const printHelp = () => {
  console.log(`
Extend level 4/5 functional locations from Drive source

Usage:
  node extend-level5.js --source <drive folder url> [options]
  node extend-level5.js --single-file --source <spreadsheet url> [options]

Options:
  --source <url>         Drive folder or spreadsheet URL/id (required)
  --single-file          Treat --source as one spreadsheet file, not a folder
  --out <file>           Output workbook (default: ./Output/extend/extend-master.xlsx)
  --help, -h             Show help
`);
};

const parseArgs = () => {
  const argv = process.argv.slice(2);
  const out = {
    source: null,
    singleFile: false,
    outPath: DEFAULT_OUT_PATH,
    help: false,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--single-file') {
      out.singleFile = true;
    } else if (a === '--source' && argv[i + 1]) {
      out.source = argv[i + 1];
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

const resolveDescCol = (colMap) => {
  for (const h of DESC_HEADERS) {
    const c = colMap.get(normalizeHeader(h));
    if (c) return c;
  }
  return undefined;
};

const countFunclocSegments = (funcloc) => {
  const raw = String(funcloc ?? '').trim();
  if (!raw) return 0;
  return raw.split('-').filter(Boolean).length;
};

const isLevel4Or5Funcloc = (funcloc) => {
  const segmentCount = countFunclocSegments(funcloc);
  return segmentCount === 4 || segmentCount === 5;
};

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

const isSpreadsheetMime = (mimeType) =>
  mimeType === GOOGLE_SHEET_MIME || mimeType === XLSX_MIME;

const collectFilteredRowsFromWorkbook = async (localPath) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(localPath);
  const rows = [];

  for (const worksheet of workbook.worksheets) {
    stripFormulasFromWorksheet(worksheet);
    const found = findHeadersExcelJs(worksheet, [FUNCLOC_HEADER, ...DESC_HEADERS], {
      maxScanRows: 10,
      minSimilarity: 0.7,
    });
    if (!found) continue;

    const { headerRowNumber, colMap } = found;
    const funclocCol = colMap.get(normalizeHeader(FUNCLOC_HEADER));
    const descCol = resolveDescCol(colMap);
    if (!funclocCol || !descCol) continue;

    for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r += 1) {
      const row = worksheet.getRow(r);
      const funcloc = excelJsValueToPlainText(row.getCell(funclocCol).value).trim();
      if (!funcloc || !isLevel4Or5Funcloc(funcloc)) continue;
      const desc = excelJsValueToPlainText(row.getCell(descCol).value);
      rows.push({
        functionalLocationAfter: funcloc,
        funclocDescAfter: desc,
      });
    }
  }

  return rows;
};

const loadRowsFromSourceFolder = async ({ sourceFolderId, tempDir }) => {
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
    console.log(`[info] Reading source: ${label}`);
    await downloadSpreadsheetToPath(file.id, file.mimeType, localPath);
    const fileRows = await collectFilteredRowsFromWorkbook(localPath);
    rows.push(...fileRows);
  }
  return { rows, fileCount: filtered.length };
};

const loadRowsFromSingleSource = async ({ sourceUrl, tempDir }) => {
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
  console.log(`[ok] Source loaded: ${name}`);
  const rows = await collectFilteredRowsFromWorkbook(localPath);
  return { rows, fileCount: 1 };
};

const writeExtendMasterWorkbook = async (rows, filePath) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('extend-master', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  ws.columns = [
    { header: FUNCLOC_HEADER, key: 'functionalLocationAfter', width: 48 },
    { header: DESC_HEADERS[0], key: 'funclocDescAfter', width: 56 },
  ];
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true };
  for (const row of rows) {
    ws.addRow(row);
  }
  ensureDir(path.dirname(filePath));
  await wb.xlsx.writeFile(filePath);
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
    const { rows, fileCount } = args.singleFile
      ? await loadRowsFromSingleSource({ sourceUrl: args.source, tempDir })
      : await loadRowsFromSourceFolder({
          sourceFolderId: extractFolderIdFromUrl(args.source),
          tempDir,
        });

    await writeExtendMasterWorkbook(rows, args.outPath);
    console.log(
      `[ok] Wrote ${rows.length} row(s) from ${fileCount} source file(s) to ${args.outPath}`,
    );
  } finally {
    cleanupTempDir(tempDir);
  }
};

main().catch((err) => {
  console.error('Failed:', err.message);
  process.exit(1);
});
