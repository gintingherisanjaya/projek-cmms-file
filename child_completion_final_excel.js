/**
 * child_completion_final_excel.js
 *
 * For every .xlsx under a Google Drive --source folder, append all rows from
 * the master sheet (FUNCTIONAL LOCATION AFTER as primary key) that don't yet
 * exist in the source file, substituting the master's plant code (e.g. 2F01)
 * with the source file's own plant code in COST CENTER AFTER and FUNCTIONAL
 * LOCATION AFTER. Outputs are written into a new `Completion_{timestamp}`
 * folder created inside the source Drive folder, mirroring the original tree.
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
  getOrCreateFolder,
  uploadXlsxReplacing,
} from './utils/googleDrive.js';
import { filterDriveSourceExcelFiles } from './utils/driveSourceFileFilter.js';
import {
  createTempDir,
  cleanupTempDir,
  runWithConcurrency,
} from './utils/concurrencyHelpers.js';
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

/** Default master xlsx (BAH JAMBI). Override with --master <id|url>. */
const DEFAULT_MASTER_URL =
  'https://docs.google.com/spreadsheets/d/1yjet6YmtkY3r1ITuw7DTRuMebEAa5maq/edit?usp=sharing';

const DEFAULT_OUT_DIR = path.join(__dirname, 'Output', 'Completion');
const DEFAULT_CONCURRENCY = 4;

/** Required headers (must exist in every source file). */
const REQUIRED_HEADERS = ['COST CENTER AFTER', 'FUNCTIONAL LOCATION AFTER'];

/** First-two-column constants copied from existing source rows for new rows. */
const SOURCE_CONSTANT_HEADERS = ['REGIONAL', 'POM'];

/** Primary header used for plant-code detection. */
const MAINTENANCE_PLANT_HEADER = 'MAINTENANCE PLANT';

/**
 * AFTER / equipment-detail columns to copy from master into appended source rows.
 * Lookup is header-driven; not every source needs every header.
 */
const AFTER_COLUMNS_TO_COPY = [
  'FUNCTLOC DESC. AFTER LEVEL 1,2,3',
  'FUNCTLOC DESC AFTRER LEVEL 1,2,3',
  'EQKTU AFTER LEVEL 2 dan 3',
  'EQUIPMENT GROUP AFTER',
  'EQUIPMENT GROUP',
  'POSITION',
  'MERK',
  'TYPE',
  'CAPACITY',
  'STANDAR UMUR TEKNIS (TAHUN)',
  'STANDART UMUR TEKNIS (TAHUN)',
  'UMUR PEMAKAIAN ALAT-% Kondisi Mesin/Peralatan (AKTUAL)',
  'NILAI EKONOMIS AKTUAL Nilai Aktiva Terakhir (Rp)',
  'WORK CENTER AFTER',
  'PLANNER GROUP AFTER',
  'DESCRIPTION',
  'MAINTENANCE PLAN AFTER',
  'ABC INDICATORS',
];

/** Every header we ever need to look up. */
const ALL_WANTED_HEADERS = [
  ...SOURCE_CONSTANT_HEADERS,
  MAINTENANCE_PLANT_HEADER,
  ...REQUIRED_HEADERS,
  ...AFTER_COLUMNS_TO_COPY,
];

const PLANT_CODE_RE = /^[A-Z0-9]{4}$/i;

const printHelp = () => {
  console.log(`
Child Completion (master sync)

Usage:
  node child_completion_final_excel.js --source <drive folder url> [options]

Options:
  --source <url>         Drive folder URL containing source .xlsx files (required)
  --master <id|url>      Master spreadsheet (default: BAH JAMBI master)
  --out <dir>            Local dir for audit workbook <Completion_folder>.xlsx (default: ./Output/Completion)
  --concurrency <n>      Parallel files (default: ${DEFAULT_CONCURRENCY})
  --help, -h             Show help
`);
};

const parseArgs = () => {
  const argv = process.argv.slice(2);
  const out = {
    source: null,
    masterUrl: DEFAULT_MASTER_URL,
    outDir: DEFAULT_OUT_DIR,
    concurrency: DEFAULT_CONCURRENCY,
    help: false,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--source' && argv[i + 1]) {
      out.source = argv[i + 1];
      i += 1;
    } else if (a === '--master' && argv[i + 1]) {
      out.masterUrl = argv[i + 1];
      i += 1;
    } else if (a === '--out' && argv[i + 1]) {
      out.outDir = argv[i + 1];
      i += 1;
    } else if (a === '--concurrency' && argv[i + 1]) {
      const n = Number.parseInt(argv[i + 1], 10);
      if (Number.isFinite(n) && n > 0) out.concurrency = n;
      i += 1;
    } else if (a === '--help' || a === '-h') {
      out.help = true;
    }
  }
  out.outDir = path.isAbsolute(out.outDir) ? out.outDir : path.join(process.cwd(), out.outDir);
  return out;
};

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
  const m = new Map(parts.map((p) => [p.type, p.value]));
  return `${m.get('year')}${m.get('month')}${m.get('day')}_${m.get('hour')}${m.get('minute')}${m.get('second')}`;
};

const escapeRegex = (value) => String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Keep only files whose top-level path segment starts with "REGIONAL". */
const isRegionalPath = (filePath) => {
  const first = String(filePath || '').split('/').find(Boolean);
  return /^REGIONAL/i.test(String(first || '').trim());
};

/**
 * Whole-substring replace of a 4-char plant code (e.g. 2F01) inside any string,
 * case-insensitive. Used for COST CENTER AFTER (`2F01STAS01` → `7F06STAS01`)
 * and FUNCTIONAL LOCATION AFTER (`PALM-2F01-...` → `PALM-7F06-...`). Plant
 * codes are unique enough that a global string replace is safe.
 */
const adaptPlant = (value, fromPlant, toPlant) => {
  if (value == null) return value;
  if (!fromPlant || !toPlant || fromPlant === toPlant) return value;
  const re = new RegExp(escapeRegex(fromPlant), 'gi');
  return String(value).replace(re, toPlant);
};

/**
 * Build COST CENTER AFTER using FUNCTIONAL LOCATION AFTER:
 * - X = group 2 (plant code), e.g. PALM-7F08-... => 7F08
 * - Y = last 2 digits of group 4, e.g. ...-0007-... => 07
 * Result format: XSTASY, e.g. 7F08STAS07
 */
const buildCostCenterFromFuncLoc = (funcLocAfter) => {
  const parts = String(funcLocAfter || '')
    .trim()
    .toUpperCase()
    .split('-')
    .filter(Boolean);
  if (parts.length < 4) return null;
  const group2 = parts[1];
  const group4 = parts[3];
  if (!PLANT_CODE_RE.test(group2)) return null;
  const digits = String(group4).replace(/\D/g, '');
  if (digits.length < 2) return null;
  const y = digits.slice(-2);
  return `${group2}STAS${y}`;
};

/** Column variants for FUNCTLOC description (spreadsheet typos). */
const FUNCDESC_HEADER_VARIANTS = [
  'FUNCTLOC DESC. AFTER LEVEL 1,2,3',
  'FUNCTLOC DESC AFTRER LEVEL 1,2,3',
];

const resolveFunclocDescCol = (colMap) => {
  for (const h of FUNCDESC_HEADER_VARIANTS) {
    const c = colMap.get(normalizeHeader(h));
    if (c) return c;
  }
  return undefined;
};

/**
 * All master FUNCTIONAL LOCATION AFTER values adapted to `sourcePlant`, for “in master?” checks.
 */
const buildMasterFlSetForSourcePlant = (masterRows, masterPlant, sourcePlant) => {
  const set = new Set();
  for (const mr of masterRows) {
    const adapted = adaptPlant(mr.funcLocAfter, masterPlant, sourcePlant);
    const key = String(adapted || '').trim().toUpperCase();
    if (key) set.add(key);
  }
  return set;
};

/**
 * Write local audit workbook: rows in source files whose FUNCLOC is not in master (same plant adaptation).
 * Does not affect Drive output.
 */
const writeMasterAuditWorkbook = async (rows, filePath) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Tidak ada di master', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  ws.columns = [
    { header: 'Nama file asal', key: 'namaFile', width: 42 },
    { header: 'Baris file asal', key: 'baris', width: 16 },
    { header: 'FUNCTIONAL LOCATION AFTER', key: 'fl', width: 48 },
    { header: 'FUNCTLOC DESC. AFTER LEVEL 1,2,3', key: 'desc', width: 56 },
  ];
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true };
  for (const row of rows) {
    ws.addRow({
      namaFile: row.sourceFileName,
      baris: row.sourceRowNumber,
      fl: row.functionalLocationAfter,
      desc: row.funclocDescAfter,
    });
  }
  ensureDir(path.dirname(filePath));
  await wb.xlsx.writeFile(filePath);
};

/**
 * Detect plant code from a worksheet by scanning, in order:
 *   1) MAINTENANCE PLANT column for any 4-char alphanumeric value
 *   2) COST CENTER AFTER column for `XXXXSTAS\d+` pattern
 *   3) FUNCTIONAL LOCATION AFTER column for `PALM-XXXX-...` pattern
 */
const detectPlantFromWorksheet = (worksheet, headerRowNumber, colMap) => {
  const maintCol = colMap.get(normalizeHeader(MAINTENANCE_PLANT_HEADER));
  const ccCol = colMap.get(normalizeHeader('COST CENTER AFTER'));
  const flCol = colMap.get(normalizeHeader('FUNCTIONAL LOCATION AFTER'));
  const maxRow = worksheet.rowCount;

  if (maintCol) {
    for (let r = headerRowNumber + 1; r <= maxRow; r += 1) {
      const v = excelJsValueToPlainText(worksheet.getRow(r).getCell(maintCol).value).trim();
      if (PLANT_CODE_RE.test(v)) return v.toUpperCase();
    }
  }
  if (ccCol) {
    for (let r = headerRowNumber + 1; r <= maxRow; r += 1) {
      const v = excelJsValueToPlainText(worksheet.getRow(r).getCell(ccCol).value).trim();
      const m = v.match(/^([A-Z0-9]{4})STAS\d+/i);
      if (m) return m[1].toUpperCase();
    }
  }
  if (flCol) {
    for (let r = headerRowNumber + 1; r <= maxRow; r += 1) {
      const v = excelJsValueToPlainText(worksheet.getRow(r).getCell(flCol).value).trim();
      const m = v.match(/^PALM-([A-Z0-9]{4})-/i);
      if (m) return m[1].toUpperCase();
    }
  }
  return null;
};

const findFirstNonEmptyInColumn = (worksheet, col, fromRow) => {
  if (!col) return '';
  const maxRow = worksheet.rowCount;
  for (let r = fromRow; r <= maxRow; r += 1) {
    const v = excelJsValueToPlainText(worksheet.getRow(r).getCell(col).value).trim();
    if (v) return v;
  }
  return '';
};

/** Walk backward from worksheet.rowCount to find the last row that has any data. */
const findLastDataRow = (worksheet, headerRowNumber) => {
  for (let r = worksheet.rowCount; r > headerRowNumber; r -= 1) {
    const row = worksheet.getRow(r);
    let any = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = excelJsValueToPlainText(cell.value).trim();
      if (v) any = true;
    });
    if (any) return r;
  }
  return headerRowNumber;
};

/**
 * Master file may be a native Google Sheet OR an uploaded .xlsx blob.
 * Detect mimeType and use the appropriate API (export vs download).
 */
const downloadMasterXlsx = async (fileId, destPath) => {
  const drive = getDrive();
  const meta = await drive.files.get({
    fileId,
    fields: 'id,name,mimeType',
    supportsAllDrives: true,
  });
  fs.mkdirSync(path.dirname(destPath), { recursive: true });
  if (meta.data.mimeType === 'application/vnd.google-apps.spreadsheet') {
    const res = await drive.files.export(
      {
        fileId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      { responseType: 'stream' },
    );
    await new Promise((resolve, reject) => {
      const ws = fs.createWriteStream(destPath);
      res.data.pipe(ws).on('finish', resolve).on('error', reject);
    });
  } else {
    await downloadFile(fileId, destPath);
  }
  return meta.data;
};

const loadMasterRows = async ({ masterUrl, tempDir }) => {
  const masterId = extractFileIdFromAnyUrl(masterUrl);
  if (!masterId) throw new Error(`Cannot extract master file id from URL: ${masterUrl}`);

  const localPath = path.join(tempDir, 'master.xlsx');
  console.log(`[info] Downloading master: ${masterId}`);
  const meta = await downloadMasterXlsx(masterId, localPath);
  console.log(`[ok] Master saved: ${meta.name} (${meta.mimeType})`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(localPath);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Master file has no worksheets');

  stripFormulasFromWorksheet(worksheet);

  const found = findHeadersExcelJs(worksheet, ALL_WANTED_HEADERS, {
    maxScanRows: 10,
    minSimilarity: 0.7,
  });
  if (!found) throw new Error('Master: header row not found');

  const { headerRowNumber, colMap } = found;
  const ccCol = colMap.get(normalizeHeader('COST CENTER AFTER'));
  const flCol = colMap.get(normalizeHeader('FUNCTIONAL LOCATION AFTER'));
  if (!ccCol || !flCol) {
    throw new Error('Master: missing COST CENTER AFTER or FUNCTIONAL LOCATION AFTER column');
  }

  const masterPlant = detectPlantFromWorksheet(worksheet, headerRowNumber, colMap);
  if (!masterPlant) throw new Error('Master: could not detect plant code (e.g. 2F01)');

  const masterRows = [];
  const seenFL = new Set();
  for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const fl = excelJsValueToPlainText(row.getCell(flCol).value).trim();
    if (!fl) continue;
    const flKey = fl.toUpperCase();
    if (seenFL.has(flKey)) continue;
    seenFL.add(flKey);

    const cc = excelJsValueToPlainText(row.getCell(ccCol).value).trim();
    const after = new Map();
    for (const h of AFTER_COLUMNS_TO_COPY) {
      const norm = normalizeHeader(h);
      if (after.has(norm)) continue;
      const c = colMap.get(norm);
      if (!c) continue;
      const cellVal = row.getCell(c).value;
      const text = excelJsValueToPlainText(cellVal);
      if (text === '' || cellVal == null) continue;
      after.set(norm, cellVal);
    }
    masterRows.push({
      funcLocAfter: fl,
      costCenterAfter: cc,
      after,
    });
  }

  console.log(
    `[ok] Master parsed: plant=${masterPlant} headerRow=${headerRowNumber} rows=${masterRows.length}`,
  );
  return { masterRows, masterPlant, headerRowNumber };
};

/**
 * Resolve a Drive folder path under `rootFolderId`, creating each segment
 * on demand. The cache stores **Promises** of folder IDs (not resolved IDs)
 * so that concurrent callers reuse the same in-flight `getOrCreateFolder`
 * call instead of all racing to create duplicate folders. The synchronous
 * `folderCache.set(built, p)` runs before the first `await` inside the IIFE,
 * so any concurrent worker entering this function next will see the Promise.
 */
const ensureFolderPath = (rootFolderId, relativeDir, folderCache) => {
  const normalized = String(relativeDir || '').split('/').filter(Boolean).join('/');
  if (!normalized || normalized === '.') return Promise.resolve(rootFolderId);
  if (folderCache.has(normalized)) return folderCache.get(normalized);

  const segments = normalized.split('/');
  let parentPromise = Promise.resolve(rootFolderId);
  let built = '';
  for (const seg of segments) {
    built = built ? `${built}/${seg}` : seg;
    if (folderCache.has(built)) {
      parentPromise = folderCache.get(built);
      continue;
    }
    const prev = parentPromise;
    const p = (async () => {
      const parentId = await prev;
      const folder = await getOrCreateFolder(seg, parentId);
      return folder.id;
    })();
    folderCache.set(built, p);
    parentPromise = p;
  }
  return parentPromise;
};

const processSourceFile = async ({
  driveFile,
  tempDir,
  mirroredFolderId,
  masterRows,
  masterPlant,
  auditRows,
}) => {
  const relPath = driveFile.path.replace(/\//g, path.sep);
  const localPath = path.join(tempDir, 'in', relPath);
  fs.mkdirSync(path.dirname(localPath), { recursive: true });
  await downloadFile(driveFile.id, localPath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(localPath);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('source has no worksheets');

  stripFormulasFromWorksheet(worksheet);

  const found = findHeadersExcelJs(worksheet, ALL_WANTED_HEADERS, {
    maxScanRows: 10,
    minSimilarity: 0.7,
  });
  if (!found) throw new Error('source: header row not found');
  const { headerRowNumber, colMap } = found;

  for (const h of [...REQUIRED_HEADERS, ...SOURCE_CONSTANT_HEADERS]) {
    if (!colMap.has(normalizeHeader(h))) {
      throw new Error(`source missing required header: ${h}`);
    }
  }

  const flCol = colMap.get(normalizeHeader('FUNCTIONAL LOCATION AFTER'));
  const ccCol = colMap.get(normalizeHeader('COST CENTER AFTER'));
  const regionalCol = colMap.get(normalizeHeader('REGIONAL'));
  const pomCol = colMap.get(normalizeHeader('POM'));

  const sourcePlant = detectPlantFromWorksheet(worksheet, headerRowNumber, colMap);
  if (!sourcePlant) throw new Error('source: could not detect plant code');

  const regionalConst = findFirstNonEmptyInColumn(worksheet, regionalCol, headerRowNumber + 1);
  const pomConst = findFirstNonEmptyInColumn(worksheet, pomCol, headerRowNumber + 1);

  const masterFlSetForSource = buildMasterFlSetForSourcePlant(
    masterRows,
    masterPlant,
    sourcePlant,
  );
  const descCol = resolveFunclocDescCol(colMap);

  const existingFL = new Set();
  let existingCount = 0;
  for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const v = excelJsValueToPlainText(row.getCell(flCol).value).trim();
    if (v) {
      existingFL.add(v.toUpperCase());
      existingCount += 1;
    }
  }

  // Audit only: FUNCLOC in this PKS file that does not exist in master (adapted to this plant).
  // Runs before appending master rows so output Excel is unchanged by this list.
  for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const flRaw = excelJsValueToPlainText(row.getCell(flCol).value).trim();
    if (!flRaw) continue;
    const flKey = flRaw.toUpperCase();
    if (masterFlSetForSource.has(flKey)) continue;
    const descText = descCol
      ? excelJsValueToPlainText(row.getCell(descCol).value).trim()
      : '';
    auditRows.push({
      sourceFileName: driveFile.name,
      sourceRowNumber: r,
      functionalLocationAfter: flRaw,
      funclocDescAfter: descText,
    });
  }

  let lastRow = findLastDataRow(worksheet, headerRowNumber);
  let appended = 0;

  for (const mr of masterRows) {
    const adaptedFL = adaptPlant(mr.funcLocAfter, masterPlant, sourcePlant);
    const flKey = String(adaptedFL || '').toUpperCase();
    if (!flKey || existingFL.has(flKey)) continue;

    const adaptedCC = adaptPlant(mr.costCenterAfter, masterPlant, sourcePlant);
    const derivedCC = buildCostCenterFromFuncLoc(adaptedFL);
    if (!derivedCC) {
      throw new Error(`cannot derive COST CENTER AFTER from FUNCTIONAL LOCATION AFTER: ${adaptedFL}`);
    }
    lastRow += 1;
    const row = worksheet.getRow(lastRow);

    if (regionalCol && regionalConst) row.getCell(regionalCol).value = regionalConst;
    if (pomCol && pomConst) row.getCell(pomCol).value = pomConst;
    // Enforce required format XSTASY from FUNCTIONAL LOCATION AFTER.
    row.getCell(ccCol).value = derivedCC || adaptedCC;
    row.getCell(flCol).value = adaptedFL;

    for (const [normHdr, value] of mr.after.entries()) {
      const c = colMap.get(normHdr);
      if (!c) continue;
      row.getCell(c).value = value;
    }

    row.commit();
    existingFL.add(flKey);
    appended += 1;
  }

  const outPath = path.join(tempDir, 'out', relPath);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  await workbook.xlsx.writeFile(outPath);

  await uploadXlsxReplacing(
    outPath,
    driveFile.name,
    mirroredFolderId,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  );

  return { existingCount, appended, sourcePlant };
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

  const sourceFolderId = extractFolderIdFromUrl(args.source);
  if (!sourceFolderId) throw new Error(`Invalid --source URL: ${args.source}`);

  await initGoogleDrive(OAUTH_JSON);
  console.log('[ok] Google Drive auth ready');

  const tempDir = createTempDir();
  console.log('[info] Temp dir:', tempDir);

  try {
    const { masterRows, masterPlant } = await loadMasterRows({
      masterUrl: args.masterUrl,
      tempDir,
    });

    console.log('[info] Listing source folder...');
    const allFiles = await listFilesRecursively(sourceFolderId);
    const { driveFiles } = filterDriveSourceExcelFiles(allFiles, { allowedOwnerEmails: [] });

    // Only process files whose top-level subfolder starts with "REGIONAL".
    const regionalOnly = driveFiles.filter((f) => isRegionalPath(f.path));
    if (regionalOnly.length !== driveFiles.length) {
      console.log(
        `[info] Skipped ${driveFiles.length - regionalOnly.length} file(s) outside REGIONAL* subfolders`,
      );
    }

    // Avoid recursing into a previous Completion_* output folder if present.
    const filteredFiles = regionalOnly.filter((f) => {
      const segs = String(f.path || '').split('/').filter(Boolean);
      return !segs.some((s) => /^Completion_/i.test(s));
    });
    if (filteredFiles.length !== regionalOnly.length) {
      console.log(
        `[info] Skipped ${regionalOnly.length - filteredFiles.length} file(s) under existing Completion_*`,
      );
    }
    console.log(`[info] Source xlsx files to process: ${filteredFiles.length}`);
    if (!filteredFiles.length) {
      console.log('[warn] No source files to process.');
      return;
    }

    const completionFolderName = `Completion_${getJakartaStamp()}`;
    const completionRoot = await getOrCreateFolder(completionFolderName, sourceFolderId);
    console.log(`[ok] Created Drive folder: ${completionFolderName}`);

    /** Rows where source FUNCLOC is missing from master (audit only; not uploaded). */
    const auditRows = [];

    const folderCache = new Map();
    let okCount = 0;
    let errCount = 0;

    await runWithConcurrency(filteredFiles, args.concurrency, async (df) => {
      const label = df.path || df.name;
      try {
        const relDir = path.posix
          .dirname(String(df.path || '').replace(/\\/g, '/'));
        const targetFolderId = await ensureFolderPath(
          completionRoot.id,
          relDir === '.' ? '' : relDir,
          folderCache,
        );

        const res = await processSourceFile({
          driveFile: df,
          tempDir,
          mirroredFolderId: targetFolderId,
          masterRows,
          masterPlant,
          auditRows,
        });
        const total = res.existingCount + res.appended;
        console.log(
          `[ok] ${label} plant=${res.sourcePlant} existing=${res.existingCount} appended=${res.appended} total=${total}`,
        );
        okCount += 1;
      } catch (e) {
        console.error(`[err] ${label}: ${e.message}`);
        errCount += 1;
      }
    });

    const auditPath = path.join(args.outDir, `${completionFolderName}.xlsx`);
    await writeMasterAuditWorkbook(auditRows, auditPath);
    console.log(
      `[ok] Audit lokal (FUNCLOC tidak ada di master): ${auditPath} (${auditRows.length} baris)`,
    );

    console.log(`\nDone. ok=${okCount} err=${errCount}`);
    console.log(`Output Drive folder: ${completionFolderName}`);
    console.log(`Inside source folder: ${args.source}`);
  } finally {
    cleanupTempDir(tempDir);
  }
};

main().catch((err) => {
  console.error('Failed:', err.message);
  process.exit(1);
});
