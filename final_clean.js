import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';

import { getCWC } from './utils/extractCWC.js';
import { getPlannerGroup } from './utils/extractPlannerGroup.js';
import { ensureDir, collectExcelFiles } from './utils/fileSystem.js';
import {
  normalizeHeader,
  toCleanString,
  toDisplayString,
  findHeadersExcelJs,
  getCellStringFromRow,
  copyStyle,
  getRefCell,
  stripFormulasFromWorksheet,
  findCostCenterCol,
} from './utils/excelHelpers.js';
import { getFuncLocPrefix, isValidPlantCode } from './utils/funclocPlantAdapt.js';
import { loadAbcMap, normalizeFunclocSuffix } from './utils/abcMap.js';
import {
  initGoogleDrive,
  listFilesInFolder,
  downloadFile,
  getOrCreateFolder,
  extractFolderIdFromUrl,
  uploadXlsxReplacing,
} from './utils/googleDrive.js';
import { bold, log, logError, logInfo, logSuccess, logWarn } from './utils/logger.js';
import { parseArgs } from './utils/cliHelpers.js';
import {
  runWithConcurrency,
  createTempDir,
  cleanupTempDir,
} from './utils/concurrencyHelpers.js';
import {
  getJakartaTimestamp,
  ensureRegionalFolders,
  uploadToDriveWithRegionalFolders,
} from './utils/driveBatch.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Google Drive source folder URL — parent of REGIONAL 1 … REGIONAL 7 (same layout as lsmw_equipment_v0.cjs)
const GOOGLE_DRIVE_FOLDER_URL =
  'https://drive.google.com/drive/folders/1zERO-GM4g1THvrUeWW-JmNX2A1eHium8';

/** Top-level REGIONAL folders to traverse (exact names after trim + upper), same whitelist as lsmw_equipment_v0.cjs. */
const ALLOWED_REGIONAL = new Set([
  'REGIONAL 1',
  'REGIONAL 2',
  'REGIONAL 3',
  'REGIONAL 4',
  'REGIONAL 5',
  'REGIONAL 6',
  'REGIONAL 7',
]);

const MIME_FOLDER = 'application/vnd.google-apps.folder';
const MIME_XLSX =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const DATA_ROOT = path.join(__dirname, 'Data');
const OUTPUT_DIR = path.join(__dirname, 'Output');

const NEW_WORK_CENTER_HEADER = 'WORK CENTER AFTER';
const NEW_PLANNER_GROUP_HEADER = 'PLANNER GROUP AFTER';
const NEW_MAINT_PLAN_HEADER = 'MAINTENANCE PLAN AFTER';
const NEW_ABC_HEADER = 'ABC INDICATORS';

const DEFAULT_CONCURRENCY = Math.max(
  1,
  Math.min(8, os.availableParallelism?.() ?? os.cpus().length ?? 8),
);

const OAUTH_JSON = path.join(__dirname, 'oauth.json');

const EQUIPMENT_GROUP_RED_ARGB = 'FFFF0000';
const EQUIPMENT_GROUP_LIGHT_RED_ARGB = 'FFFFC7CE';

const HIGHLIGHT_YELLOW_ARGB = 'FFFFEB9C';

/**
 * Highlight baris kuning untuk alert validasi — tanpa menimpa fill asal di kolom yang di-skip
 * (EQUIPMENT GROUP AFTER / EQKTU AFTER: merah & merah muda dari file sumber harus tetap).
 */
const applyYellowRowFill = (row, fromCol, toCol, skipCols = new Set()) => {
  for (let c = fromCol; c <= toCol; c += 1) {
    if (skipCols.has(c)) continue;
    row.getCell(c).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: HIGHLIGHT_YELLOW_ARGB },
    };
  }
};

/** Safe filename stem for temp files alongside random download names. */
const sanitizeForFsFilename = (name) =>
  String(name).replace(/[/\\:?*"<>|]/g, '_');

/** Output `.xlsx` basename for uploads — uses Drive `name` when present (not temp `in_*` path). */
const resolvedOutputXlsxName = (excelPath, driveFileInfo) => {
  const n = driveFileInfo?.name;
  if (n && /\.xlsx$/i.test(String(n))) return String(n);
  return path.basename(excelPath);
};

/** Basename for upload/save — original Drive name (`outputUploadName`), not temp `temp_fc_*` path. */
const outputSaveBasename = (summary) => {
  if (summary.outputUploadName) return summary.outputUploadName;
  const p = summary.outputPath;
  if (!p) return 'output.xlsx';
  return path
    .basename(p)
    .replace(/^temp_fc_/, '')
    .replace(/^temp_/, '');
};

/** Bold FUNCTLOC DESC for 4- and 5-segment FUNCLOC rows. */
const boldFunclocDescOnSheet = (
  outSheet,
  funclocCol,
  funclocDescCol,
  headerRowIndex,
  dataRowCount,
) => {
  if (!funclocCol || !funclocDescCol || dataRowCount <= 0) return;
  const dataStartRow = headerRowIndex + 1;
  const dataEndRow = headerRowIndex + dataRowCount;
  for (let r = dataStartRow; r <= dataEndRow; r += 1) {
    const row = outSheet.getRow(r);
    const rawFuncloc = toCleanString(row.getCell(funclocCol).value);
    if (!rawFuncloc) continue;
    const segments = rawFuncloc.split('-').filter(Boolean);
    if (segments.length !== 4 && segments.length !== 5) continue;
    const cell = row.getCell(funclocDescCol);
    const baseStyle = cell.style || {};
    const font = { ...(baseStyle.font || {}), bold: true };
    cell.style = { ...baseStyle, font };
    row.commit();
  }
};

/** First valid 4-char plant code in a column (data rows below header), or null. */
const findFirstValidPlantInColumn = (ws, col, headerRowNumber) => {
  if (!col) return null;
  for (let r = headerRowNumber + 1; r <= ws.rowCount; r += 1) {
    const val = getCellStringFromRow(ws.getRow(r), col);
    if (!val) continue;
    const normalized = val.trim().toUpperCase();
    if (isValidPlantCode(normalized)) return normalized;
  }
  return null;
};

/** Plant from second FUNCLOC segment (e.g. PALM-2F01-… → 2F01). */
const findFirstPlantFromFunclocColumn = (ws, col, headerRowNumber) => {
  if (!col) return null;
  for (let r = headerRowNumber + 1; r <= ws.rowCount; r += 1) {
    const raw = getCellStringFromRow(ws.getRow(r), col);
    if (!raw) continue;
    const prefix = getFuncLocPrefix(raw);
    if (!prefix) continue;
    const segs = prefix.split('-');
    const plantCode = segs.length >= 2 ? segs[1] : null;
    if (isValidPlantCode(plantCode)) return String(plantCode).trim().toUpperCase();
  }
  return null;
};

/**
 * MAINTENANCE PLAN AFTER needs a 4-char plant: MAINTENANCE PLANT column, else POM, else FUNCTIONAL LOCATION AFTER.
 * @returns {{ plant: string | null, source: string | null }}
 */
const resolveMaintenancePlantForClean = (
  ws,
  headerRowNumber,
  maintPlantCol,
  pomCol,
  funclocCol,
) => {
  if (maintPlantCol) {
    const p = findFirstValidPlantInColumn(ws, maintPlantCol, headerRowNumber);
    if (p) return { plant: p, source: 'MAINTENANCE PLANT' };
  }
  if (pomCol) {
    const p = findFirstValidPlantInColumn(ws, pomCol, headerRowNumber);
    if (p) return { plant: p, source: 'POM' };
  }
  if (funclocCol) {
    const p = findFirstPlantFromFunclocColumn(ws, funclocCol, headerRowNumber);
    if (p) return { plant: p, source: 'FUNCTIONAL LOCATION AFTER' };
  }
  return { plant: null, source: null };
};

/**
 * Check if an exceljs cell has red / light-red validation fill.
 */
const isCellRedFill = (cell) => {
  if (!cell || !cell.fill) return false;
  const fill = cell.fill;
  if (fill.type !== 'pattern') return false;
  const fg = fill.fgColor;
  if (!fg) return false;
  const argb = String(fg.argb || '').toUpperCase();
  if (argb === EQUIPMENT_GROUP_RED_ARGB) return true;
  if (argb === EQUIPMENT_GROUP_LIGHT_RED_ARGB) return true;
  // Also check last 6 digits in case of different alpha
  if (argb.length === 8 && argb.slice(2) === 'FF0000') return true;
  if (argb.length === 8 && argb.slice(2) === 'FFC7CE') return true;
  return false;
};

/**
 * Find header row and column map using exceljs worksheet.
 * Scans first 10 rows for wanted header texts.
 * Returns { headerRowNumber, colMap } where colMap is Map<normalizedHeader, colNumber (1-based)>
 */
const findHeaders = (ws, wantedHeaders) => {
  const wanted = wantedHeaders.map(normalizeHeader);
  const maxRow = Math.min(ws.rowCount, 10);

  for (let r = 1; r <= maxRow; r += 1) {
    const row = ws.getRow(r);
    const colMap = new Map();

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const norm = normalizeHeader(String(cell.value ?? ''));
      if (!norm) return;

      // Exact match
      if (wanted.includes(norm)) {
        colMap.set(norm, colNumber);
      }

      // Substring match (for headers that contain our wanted text)
      for (const w of wanted) {
        if (!colMap.has(w) && (norm.includes(w) || w.includes(norm))) {
          colMap.set(w, colNumber);
        }
      }
    });

    if (wanted.some((w) => colMap.has(w))) {
      return { headerRowNumber: r, colMap };
    }
  }

  return null;
};

/**
 * Get cell value as clean string (use formula result if present).
 */
const getCellString = (row, col) => getCellStringFromRow(row, col);

/**
 * Cell value for JSON export: keep numbers, convert Date to ISO string, null/undefined -> null.
 */
const cellValueForJson = (v) => {
  if (v == null) return null;
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v.toISOString();
  return String(v);
};

/**
 * Escape a field for CSV: wrap in quotes if contains comma, quote, or newline.
 */
const escapeCsvField = (v) => {
  const s = v == null ? '' : String(v);
  if (/[,"\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
};

/**
 * Process a single Excel file using exceljs:
 * - Read with full style preservation
 * - Find required headers
 * - Rows with red EQUIPMENT GROUP / EQKTU: kept; yellow row highlight on output
 * - Append WORK CENTER AFTER, PLANNER GROUP AFTER, MAINTENANCE PLAN AFTER, ABC INDICATORS
 * - Optional bold FUNCTLOC DESC, then yellow row fill for validation-red rows (order preserves fills)
 * - Write back preserving cell styling; uploaded/saved `.xlsx` basename = Drive original name when available
 */
const processExcelFile = async (
  excelPath,
  driveFileInfo = null,
  { boldFunclocDesc = true } = {},
) => {
  const displayPath = driveFileInfo ? driveFileInfo.path : excelPath;
  const outXlsxName = resolvedOutputXlsxName(excelPath, driveFileInfo);
  log(`\n${bold('Processing:')} ${displayPath}`);

  const abcMap = loadAbcMap();

  // Read using exceljs (preserves all styles, fills, fonts, borders, etc.)
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);

  const ws = workbook.worksheets[0];
  if (!ws) {
    logWarn('No worksheet found in file');
    return { excelPath: displayPath, sheetName: null, headerFound: false, driveFileInfo };
  }

  // Final cleanup: replace all formulas with their values so we only read/write values
  stripFormulasFromWorksheet(ws);
  logInfo('Stripped formulas to values (final cleanup).');

  const sheetName = ws.name;
  logInfo(`Sheet: ${sheetName}`);

  // Find header row
  const wantedHeaders = [
    'MAINTENANCE PLANT',
    'POM',
    'FUNCTIONAL LOCATION AFTER',
    'FUNCTLOC DESC. AFTER LEVEL 1,2,3',
    'FUNCTLOC DESC AFTRER LEVEL 1,2,3',
    'COST CENTER AFTER',
    'COST CENTER',
    'EQUIPMENT GROUP AFTER',
    'EQKTU AFTER LEVEL 2 dan 3',
    'EQKTU AFTER LEVEL 2 AND 3',
  ];

  const headerResult = findHeadersExcelJs(ws, wantedHeaders);
  if (!headerResult) {
    logWarn(`Could not find header row for: ${wantedHeaders.join(', ')}`);
    return { excelPath: displayPath, sheetName, headerFound: false, driveFileInfo };
  }

  const { headerRowNumber, colMap } = headerResult;

  const maintPlantCol = colMap.get(normalizeHeader('MAINTENANCE PLANT'));
  const pomCol = colMap.get(normalizeHeader('POM'));
  const funclocCol = colMap.get(normalizeHeader('FUNCTIONAL LOCATION AFTER'));
  const funclocDescCol =
    colMap.get(normalizeHeader('FUNCTLOC DESC. AFTER LEVEL 1,2,3')) ??
    colMap.get(normalizeHeader('FUNCTLOC DESC AFTRER LEVEL 1,2,3'));
  const equipGroupCol = colMap.get(normalizeHeader('EQUIPMENT GROUP AFTER'));
  const eqktuAfterCol =
    colMap.get(normalizeHeader('EQKTU AFTER LEVEL 2 dan 3')) ??
    colMap.get(normalizeHeader('EQKTU AFTER LEVEL 2 AND 3'));
  const costCenterCol = findCostCenterCol(colMap);

  if (!funclocCol) {
    logWarn('FUNCTIONAL LOCATION AFTER column not found');
    return {
      excelPath: displayPath,
      sheetName,
      headerFound: true,
      error: 'FUNCTIONAL LOCATION AFTER column missing',
      driveFileInfo,
    };
  }
  if (!costCenterCol) {
    logWarn('COST CENTER AFTER or COST CENTER column not found');
    return {
      excelPath: displayPath,
      sheetName,
      headerFound: true,
      error: 'COST CENTER column missing',
      driveFileInfo,
    };
  }

  logInfo(
    `Found header row: ${headerRowNumber}` +
      (maintPlantCol
        ? `, MAINTENANCE PLANT: col ${maintPlantCol}`
        : ', MAINTENANCE PLANT: (no column — will use POM / FUNCLOC if available)') +
      (pomCol ? `, POM: col ${pomCol}` : '') +
      `, FUNCTIONAL LOCATION AFTER: col ${funclocCol}` +
      `, COST CENTER: col ${costCenterCol}` +
      (equipGroupCol ? `, EQUIPMENT GROUP AFTER: col ${equipGroupCol}` : '') +
      (eqktuAfterCol ? `, EQKTU AFTER: col ${eqktuAfterCol}` : '') +
      (funclocDescCol ? `, FUNCTLOC DESC AFTER: col ${funclocDescCol}` : ''),
  );

  // Find the last used column in the header row (original data boundary)
  const headerRow = ws.getRow(headerRowNumber);
  let lastOrigCol = 0;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (colNumber > lastOrigCol) lastOrigCol = colNumber;
  });

  const { plant: maintenancePlant, source: maintenancePlantSource } =
    resolveMaintenancePlantForClean(ws, headerRowNumber, maintPlantCol, pomCol, funclocCol);

  if (!maintPlantCol) {
    if (maintenancePlant) {
      logInfo(
        `MAINTENANCE PLANT column missing; using plant ${maintenancePlant} from ${maintenancePlantSource} for MAINTENANCE PLAN AFTER.`,
      );
    } else {
      logWarn(
        'MAINTENANCE PLANT column missing and no valid 4-char plant in POM or FUNCTIONAL LOCATION AFTER; MAINTENANCE PLAN AFTER will be empty.',
      );
    }
  } else if (!maintenancePlant) {
    logWarn(
      'No valid 4-char plant in MAINTENANCE PLANT, POM, or FUNCLOC; MAINTENANCE PLAN AFTER column will be empty.',
    );
  } else {
    logInfo(
      `Maintenance Plant (for MAINTENANCE PLAN AFTER): ${maintenancePlant}` +
        (maintenancePlantSource && maintenancePlantSource !== 'MAINTENANCE PLANT'
          ? ` (from ${maintenancePlantSource})`
          : ''),
    );
  }

  // === Step 1: Rows with red EQUIPMENT GROUP AFTER and/or EQKTU AFTER (yellow highlight on output) ===
  const redRows = [];
  for (let r = headerRowNumber + 1; r <= ws.rowCount; r += 1) {
    const row = ws.getRow(r);
    let flagged = false;
    if (equipGroupCol && isCellRedFill(row.getCell(equipGroupCol))) flagged = true;
    if (eqktuAfterCol && isCellRedFill(row.getCell(eqktuAfterCol))) flagged = true;
    if (flagged) redRows.push(r);
  }

  if (redRows.length > 0) {
    logInfo(
      `${redRows.length} row(s) have red fill on EQUIPMENT GROUP AFTER and/or EQKTU AFTER — ` +
        'output rows will use yellow row highlight.',
    );
  }

  const redRowSet = new Set(redRows);

  /** Data rows in sheet order (no level 5→6 column redistribution). */
  const orderedOutputRows = [];
  for (let r = headerRowNumber + 1; r <= ws.rowCount; r += 1) {
    orderedOutputRows.push({ type: 'original', rowIndex: r });
  }

  // === Step 2: Create a NEW value-only output sheet ===
  // We do NOT touch the original sheet's formulas or row structure at all.
  // Instead, we build a clean sheet that:
  // - contains only values (no formulas)
  // - keeps all data rows; red-flagged EQUIPMENT GROUP / EQKTU rows get yellow fill on export
  // - appends our AFTER columns at the end
  const finalSheetName = `${sheetName}_FINAL`;
  const finalWs = workbook.addWorksheet(finalSheetName);

  // Copy any rows that appear BEFORE the header row (e.g. titles, notes) as values only
  let targetRowIndex = 1;
  for (let r = 1; r < headerRowNumber; r += 1) {
    const srcRow = ws.getRow(r);
    const dstRow = finalWs.getRow(targetRowIndex);
    srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const dstCell = dstRow.getCell(colNumber);
      const v = cell.value;
      if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
        // Use cached result if present; otherwise empty
        dstCell.value = v.result ?? null;
      } else {
        dstCell.value = v;
      }
      copyStyle(cell, dstCell);
    });
    dstRow.commit();
    targetRowIndex += 1;
  }

  // === Step 3: Write header row in the new sheet (original headers + AFTER columns) ===
  const finalHeaderRowIndex = targetRowIndex;
  const hdrRowSrc = ws.getRow(headerRowNumber);
  const hdrRow = finalWs.getRow(finalHeaderRowIndex);
  const refHeaderCellSrc = getRefCell(hdrRowSrc, costCenterCol, lastOrigCol);

  // Copy original headers as values only
  for (let c = 1; c <= lastOrigCol; c += 1) {
    const srcCell = hdrRowSrc.getCell(c);
    const dstCell = hdrRow.getCell(c);
    const v = srcCell.value;
    if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
      dstCell.value = v.result ?? null;
    } else {
      dstCell.value = v;
    }
    copyStyle(srcCell, dstCell);
  }

  // New columns will be appended after the original boundary
  const cwcCol = lastOrigCol + 1;
  const pgCol = lastOrigCol + 2;
  const mpCol = lastOrigCol + 3;
  const abcCol = lastOrigCol + 4;

  // New headers in final sheet
  const cwcHdrCell = hdrRow.getCell(cwcCol);
  cwcHdrCell.value = NEW_WORK_CENTER_HEADER;
  if (refHeaderCellSrc) copyStyle(refHeaderCellSrc, cwcHdrCell);

  const pgHdrCell = hdrRow.getCell(pgCol);
  pgHdrCell.value = NEW_PLANNER_GROUP_HEADER;
  if (refHeaderCellSrc) copyStyle(refHeaderCellSrc, pgHdrCell);

  const mpHdrCell = hdrRow.getCell(mpCol);
  mpHdrCell.value = NEW_MAINT_PLAN_HEADER;
  if (refHeaderCellSrc) copyStyle(refHeaderCellSrc, mpHdrCell);

  const abcHdrCell = hdrRow.getCell(abcCol);
  abcHdrCell.value = NEW_ABC_HEADER;
  if (refHeaderCellSrc) copyStyle(refHeaderCellSrc, abcHdrCell);

  hdrRow.commit();
  targetRowIndex += 1;

  // === Step 4: Copy data rows into the new sheet (original order, values only) ===
  let totalRows = 0;
  let cwcFoundCount = 0;
  let cwcMissingCount = 0;
  let plannerGroupFoundCount = 0;
  let plannerGroupMissingCount = 0;
  let abcMappedCount = 0;
  let abcDefaultCount = 0;
  const yellowHighlightFinalRows = new Set();

  for (const out of orderedOutputRows) {
    const srcRow = ws.getRow(out.rowIndex);
    const funclocRaw = getCellString(srcRow, funclocCol);
    const costCenterRaw = getCellString(srcRow, costCenterCol);

    const row = finalWs.getRow(targetRowIndex);

    let anyValue = false;
    for (let c = 1; c <= lastOrigCol; c += 1) {
      const val = toCleanString(srcRow.getCell(c).value);
      if (val) {
        anyValue = true;
        break;
      }
    }
    if (!anyValue) continue;

    totalRows += 1;

    for (let c = 1; c <= lastOrigCol; c += 1) {
      const srcCell = srcRow.getCell(c);
      const dstCell = row.getCell(c);
      const v = srcCell.value;
      if (v != null && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
        dstCell.value = v.result ?? null;
      } else {
        dstCell.value = v;
      }
      copyStyle(srcCell, dstCell);
    }

    const effectiveFuncloc = funclocRaw;
    const effectiveCostCenter = costCenterRaw;

    const refCell = getRefCell(row, costCenterCol, lastOrigCol);

    const cwc = getCWC('', effectiveCostCenter);
    const cwcCell = row.getCell(cwcCol);
    cwcCell.value = cwc ?? '';
    if (refCell) copyStyle(refCell, cwcCell);
    if (cwc) cwcFoundCount += 1;
    else cwcMissingCount += 1;

    const plannerGroup = getPlannerGroup(effectiveFuncloc, effectiveCostCenter);
    const pgCell = row.getCell(pgCol);
    pgCell.value = plannerGroup ?? '';
    if (refCell) copyStyle(refCell, pgCell);
    if (plannerGroup) plannerGroupFoundCount += 1;
    else plannerGroupMissingCount += 1;

    const mpCell = row.getCell(mpCol);
    mpCell.value = maintenancePlant ?? '';
    if (refCell) copyStyle(refCell, mpCell);

    const funclocSuffix = normalizeFunclocSuffix(effectiveFuncloc);
    let abcValue = '2';
    if (funclocSuffix && abcMap.size > 0) {
      const mapped = abcMap.get(funclocSuffix);
      if (mapped && String(mapped).trim() !== '') {
        abcValue = String(mapped).trim();
        abcMappedCount += 1;
      } else {
        abcDefaultCount += 1;
      }
    } else {
      abcDefaultCount += 1;
    }

    const abcCell = row.getCell(abcCol);
    abcCell.value = abcValue;
    if (refCell) copyStyle(refCell, abcCell);

    if (redRowSet.has(out.rowIndex)) {
      yellowHighlightFinalRows.add(targetRowIndex);
    }

    row.commit();
    targetRowIndex += 1;
  }

  // === Export same data as JSON and CSV for cross-check (easy to read) ===
  // Reorder data rows A-Z by "FUNCTIONAL LOCATION AFTER" before exporting.
  const outputDir = path.dirname(excelPath);
  const exportBaseStem = path.basename(outXlsxName, path.extname(outXlsxName));
  const hdrRowFinal = finalWs.getRow(finalHeaderRowIndex);
  const headerNames = [];
  for (let c = 1; c <= abcCol; c += 1) {
    const v = hdrRowFinal.getCell(c).value;
    headerNames.push(v != null ? String(v).trim() : `Col${c}`);
  }
  const dataStartRow = finalHeaderRowIndex + 1;
  const dataEndRow = targetRowIndex - 1;
  const sortedDataRowIndices = [];
  for (let r = dataStartRow; r <= dataEndRow; r += 1) {
    const row = finalWs.getRow(r);
    const key = toCleanString(row.getCell(funclocCol).value);
    sortedDataRowIndices.push({ srcRowIndex: r, key });
  }
  sortedDataRowIndices.sort((a, b) => {
    const ak = String(a.key ?? '').toUpperCase();
    const bk = String(b.key ?? '').toUpperCase();
    const cmp = ak.localeCompare(bk);
    return cmp !== 0 ? cmp : a.srcRowIndex - b.srcRowIndex;
  });

  const outputRows = [];
  for (const item of sortedDataRowIndices) {
    const row = finalWs.getRow(item.srcRowIndex);
    const obj = {};
    for (let c = 1; c <= abcCol; c += 1) {
      const key = headerNames[c - 1] || `Col${c}`;
      obj[key] = cellValueForJson(row.getCell(c).value);
    }
    outputRows.push(obj);
  }
  const jsonPath = path.join(outputDir, `${exportBaseStem}_output.json`);
  const csvPath = path.join(outputDir, `${exportBaseStem}_output.csv`);
  ensureDir(outputDir);
  await fs.promises.writeFile(jsonPath, JSON.stringify(outputRows, null, 2), 'utf8');
  const csvLines = [headerNames.map(escapeCsvField).join(',')];
  for (const obj of outputRows) {
    csvLines.push(headerNames.map((h) => escapeCsvField(obj[h] ?? '')).join(','));
  }
  await fs.promises.writeFile(csvPath, csvLines.join('\n'), 'utf8');
  logInfo(`Cross-check exports: ${path.basename(jsonPath)}, ${path.basename(csvPath)}`);

  // === Write with ExcelJS only: fresh workbook, one sheet, values + cell styling ===
  const outWorkbook = new ExcelJS.Workbook();
  const outSheet = outWorkbook.addWorksheet(sheetName);
  const maxCol = abcCol;

  // Copy title rows and header row as-is.
  for (let r = 1; r <= finalHeaderRowIndex; r += 1) {
    const srcRow = finalWs.getRow(r);
    const dstRow = outSheet.getRow(r);
    for (let c = 1; c <= maxCol; c += 1) {
      const srcCell = srcRow.getCell(c);
      const dstCell = dstRow.getCell(c);
      const v = srcCell.value;
      if (v == null) dstCell.value = '';
      else if (typeof v === 'number' || typeof v === 'string') dstCell.value = v;
      else if (v instanceof Date) dstCell.value = v.toISOString();
      else dstCell.value = String(v);
      copyStyle(srcCell, dstCell);
    }
    dstRow.commit();
  }

  // Copy data rows (sorted); bold FUNCTLOC DESC before yellow so fills are not overwritten.
  const yellowHighlightOutRows = new Set();
  for (let i = 0; i < sortedDataRowIndices.length; i += 1) {
    const srcRowIndex = sortedDataRowIndices[i].srcRowIndex;
    const srcRow = finalWs.getRow(srcRowIndex);
    const dstRow = outSheet.getRow(finalHeaderRowIndex + 1 + i);
    for (let c = 1; c <= maxCol; c += 1) {
      const srcCell = srcRow.getCell(c);
      const dstCell = dstRow.getCell(c);
      const v = srcCell.value;
      if (v == null) dstCell.value = '';
      else if (typeof v === 'number' || typeof v === 'string') dstCell.value = v;
      else if (v instanceof Date) dstCell.value = v.toISOString();
      else dstCell.value = String(v);
      copyStyle(srcCell, dstCell);
    }
    if (yellowHighlightFinalRows.has(srcRowIndex)) {
      yellowHighlightOutRows.add(finalHeaderRowIndex + 1 + i);
    }
    dstRow.commit();
  }

  const outDataRowCount = sortedDataRowIndices.length;
  if (boldFunclocDesc) {
    boldFunclocDescOnSheet(
      outSheet,
      funclocCol,
      funclocDescCol,
      finalHeaderRowIndex,
      outDataRowCount,
    );
  }

  const yellowSkipFillCols = new Set();
  if (equipGroupCol) yellowSkipFillCols.add(equipGroupCol);
  if (eqktuAfterCol) yellowSkipFillCols.add(eqktuAfterCol);

  for (const outRowNum of yellowHighlightOutRows) {
    const yRow = outSheet.getRow(outRowNum);
    applyYellowRowFill(yRow, 1, maxCol, yellowSkipFillCols);
    yRow.commit();
  }
  if (yellowHighlightOutRows.size > 0) {
    logInfo(
      `Yellow row highlight: ${yellowHighlightOutRows.size} row(s) (source red on EQUIPMENT GROUP AFTER and/or EQKTU AFTER); ` +
        'fill asal kolom EQUIPMENT GROUP AFTER / EQKTU AFTER tidak ditimpa kuning.',
    );
  }

  // Preserve column widths from source worksheet (ExcelJS only)
  for (let c = 1; c <= maxCol; c += 1) {
    const srcCol = ws.getColumn(c);
    if (srcCol && srcCol.width != null) {
      outSheet.getColumn(c).width = srcCol.width;
    }
  }

  // Logging
  logInfo(`Rows scanned: ${totalRows}`);
  logSuccess(`CWC found: ${cwcFoundCount}`);
  if (cwcMissingCount > 0) logWarn(`CWC missing/empty: ${cwcMissingCount}`);
  logSuccess(`Planner Group found: ${plannerGroupFoundCount}`);
  if (plannerGroupMissingCount > 0)
    logWarn(`Planner Group missing/empty: ${plannerGroupMissingCount}`);
  if (abcMappedCount || abcDefaultCount) {
    logSuccess(`ABC indicators mapped: ${abcMappedCount}`);
    if (abcDefaultCount > 0) logInfo(`ABC indicators defaulted to 2: ${abcDefaultCount}`);
  }

  const tempOutputPath = path.join(
    path.dirname(excelPath),
    `temp_fc_${sanitizeForFsFilename(outXlsxName)}`,
  );
  ensureDir(path.dirname(tempOutputPath));
  await outWorkbook.xlsx.writeFile(tempOutputPath);
  logSuccess(`Processed file: ${displayPath}`);

  return {
    excelPath: displayPath,
    sheetName,
    headerFound: true,
    totalRows,
    cwcFound: cwcFoundCount,
    cwcMissing: cwcMissingCount,
    plannerGroupFound: plannerGroupFoundCount,
    plannerGroupMissing: plannerGroupMissingCount,
    maintenancePlant,
    outputPath: tempOutputPath,
    outputUploadName: outXlsxName,
    outputJsonPath: jsonPath,
    outputCsvPath: csvPath,
    driveFileInfo,
  };
};

const logFinalProcessSummary = (summaries) => {
  if (!summaries || summaries.length === 0) return;

  let filesOk = 0;
  let filesHeaderMissing = 0;
  let filesError = 0;
  let totalRows = 0;
  let totalCwcFound = 0;
  let totalCwcMissing = 0;
  let totalPlannerFound = 0;
  let totalPlannerMissing = 0;

  for (const s of summaries) {
    if (s.error) {
      filesError += 1;
      continue;
    }
    if (!s.headerFound) {
      filesHeaderMissing += 1;
      continue;
    }
    filesOk += 1;
    totalRows += s.totalRows ?? 0;
    totalCwcFound += s.cwcFound ?? 0;
    totalCwcMissing += s.cwcMissing ?? 0;
    totalPlannerFound += s.plannerGroupFound ?? 0;
    totalPlannerMissing += s.plannerGroupMissing ?? 0;
  }

  log(`\n${bold('=== Final Clean Summary ===')}`);
  log(`Files processed (headers OK): ${filesOk}/${summaries.length}`);
  if (filesHeaderMissing > 0) {
    logWarn(`Files with missing headers: ${filesHeaderMissing}`);
  }
  if (filesError > 0) {
    logError(`Files with processing errors: ${filesError}`);
  }
  logInfo(`Total data rows scanned: ${totalRows}`);
  logSuccess(`Total CWC found: ${totalCwcFound}`);
  if (totalCwcMissing > 0) {
    logWarn(`Total CWC missing/empty: ${totalCwcMissing}`);
  }
  logSuccess(`Total Planner Group found: ${totalPlannerFound}`);
  if (totalPlannerMissing > 0) {
    logWarn(`Total Planner Group missing/empty: ${totalPlannerMissing}`);
  }
};

const unlinkQuiet = (p) => {
  try {
    if (p && fs.existsSync(p)) fs.unlinkSync(p);
  } catch {
    /* ok */
  }
};

async function processExcelSafely(excelPath, driveFileInfo, opts) {
  try {
    return await processExcelFile(excelPath, driveFileInfo, opts);
  } catch (err) {
    let errorMsg = 'Unknown error';
    if (err) {
      if (typeof err === 'string') errorMsg = err;
      else if (err.message) errorMsg = String(err.message);
      else errorMsg = String(err);
    }
    const displayPath = driveFileInfo ? driveFileInfo.path : excelPath;
    logError(`Failed to process ${displayPath}: ${errorMsg}`);
    return {
      excelPath: driveFileInfo ? driveFileInfo.path : excelPath,
      error: errorMsg,
      driveFileInfo,
    };
  }
}

/**
 * Mirrors lsmw_equipment_v0.processFolder: walk one subtree, handle each `.xlsx` immediately
 * (download → Final Clean → upload or save → delete temps).
 */
async function regionalSubtreeSequentialPipeline({
  sourceFolderId,
  driveTargetFolderId,
  drivePathPrefix,
  tempDir,
  boldFunclocDesc,
  outputToLocal,
  localOutputRoot,
  driveSamples,
  sampleCounter,
}) {
  const summaries = [];
  const entries = await listFilesInFolder(sourceFolderId);

  for (const item of entries) {
    if (driveSamples != null && driveSamples > 0 && sampleCounter.count >= driveSamples) {
      break;
    }

    if (item.mimeType === MIME_FOLDER) {
      let nextDriveId = driveTargetFolderId;
      if (!outputToLocal && driveTargetFolderId) {
        const subDrive = await getOrCreateFolder(item.name, driveTargetFolderId);
        nextDriveId = subDrive.id;
      }
      const nextLocal = localOutputRoot ? path.join(localOutputRoot, item.name) : null;
      if (nextLocal) fs.mkdirSync(nextLocal, { recursive: true });

      summaries.push(
        ...(await regionalSubtreeSequentialPipeline({
          sourceFolderId: item.id,
          driveTargetFolderId: nextDriveId,
          drivePathPrefix: `${drivePathPrefix}${item.name}/`,
          tempDir,
          boldFunclocDesc,
          outputToLocal,
          localOutputRoot: nextLocal,
          driveSamples,
          sampleCounter,
        })),
      );
      continue;
    }

    if (item.mimeType !== MIME_XLSX || /^~\$/i.test(item.name)) continue;

    const relPath = `${drivePathPrefix}${item.name}`;
    const driveFileInfo = {
      id: item.id,
      name: item.name,
      mimeType: item.mimeType,
      path: relPath,
    };

    const localIn = path.join(
      tempDir,
      `in_${Date.now()}_${Math.random().toString(36).slice(2)}_${sanitizeForFsFilename(item.name)}`,
    );

    logInfo(`Downloading: ${relPath}`);
    await downloadFile(item.id, localIn);

    const summary = await processExcelSafely(localIn, driveFileInfo, {
      boldFunclocDesc,
    });

    unlinkQuiet(localIn);

    if (!summary.error && summary.outputPath && fs.existsSync(summary.outputPath)) {
      const uploadName = outputSaveBasename(summary);

      if (outputToLocal && localOutputRoot) {
        const targetPath = path.join(localOutputRoot, uploadName);
        fs.mkdirSync(path.dirname(targetPath), { recursive: true });
        fs.copyFileSync(summary.outputPath, targetPath);
        logSuccess(`Saved: ${uploadName}`);
      } else if (!outputToLocal && driveTargetFolderId) {
        await uploadXlsxReplacing(summary.outputPath, uploadName, driveTargetFolderId);
        logSuccess(`Uploaded: ${uploadName}`);
      }

      unlinkQuiet(summary.outputPath);
    }

    sampleCounter.count += 1;
    summaries.push(summary);
  }

  return summaries;
}

async function runRegionalRootPipelineLikeLsmw({
  tempDir,
  boldFunclocDesc,
  outputToLocal,
  driveSamples,
}) {
  const sourceRootId = extractFolderIdFromUrl(GOOGLE_DRIVE_FOLDER_URL);
  if (!sourceRootId) throw new Error('Invalid GOOGLE_DRIVE_FOLDER_URL folder id');

  const timestamp = getJakartaTimestamp();

  let driveResultsFolder = null;
  let localBatchRoot = null;

  if (outputToLocal) {
    localBatchRoot = path.join(OUTPUT_DIR, `Final_Clean_Results_${timestamp}`);
    fs.mkdirSync(localBatchRoot, { recursive: true });
    logSuccess(`Created local output directory: ${localBatchRoot}`);
  } else {
    driveResultsFolder = await getOrCreateFolder(`Final_Clean_Results_${timestamp}`, sourceRootId);
    logSuccess(`Created Drive output folder: ${driveResultsFolder.name}`);
  }

  const sampleCounter = { count: 0 };
  const allSummaries = [];

  const top = await listFilesInFolder(sourceRootId);

  for (const folder of top) {
    if (folder.mimeType !== MIME_FOLDER) continue;

    const nameKey = folder.name.trim().toUpperCase();
    if (!ALLOWED_REGIONAL.has(nameKey)) {
      logInfo(`Skip folder: ${folder.name}`);
      continue;
    }

    logInfo(`Processing ${folder.name}`);
    let regionalDriveId = null;
    let regionalLocalRoot = null;

    if (outputToLocal) {
      regionalLocalRoot = path.join(localBatchRoot, folder.name);
      fs.mkdirSync(regionalLocalRoot, { recursive: true });
    } else {
      const r = await getOrCreateFolder(folder.name, driveResultsFolder.id);
      regionalDriveId = r.id;
    }

    const part = await regionalSubtreeSequentialPipeline({
      sourceFolderId: folder.id,
      driveTargetFolderId: regionalDriveId,
      drivePathPrefix: `${folder.name}/`,
      tempDir,
      boldFunclocDesc,
      outputToLocal,
      localOutputRoot: regionalLocalRoot,
      driveSamples,
      sampleCounter,
    });
    allSummaries.push(...part);

    if (driveSamples != null && driveSamples > 0 && sampleCounter.count >= driveSamples) {
      break;
    }
  }

  return { summaries: allSummaries, localBatchRoot };
}

const main = async () => {
  const argv = process.argv.slice(2);

  let sourceLocal;
  let samples;
  let driveSamples;
  let singleFileId;
  let outputToLocal;
  let boldFunclocDesc = true;

  if (argv.length > 0) {
    const argResult = parseArgs(argv);
    sourceLocal = Boolean(argResult.local);
    samples = argResult.samples ?? null;
    driveSamples = argResult.driveSamples ?? null;
    singleFileId = argResult.singleFileId || argResult.driveSingleFileId || null;

    if (argv.includes('--output-local')) outputToLocal = true;
    else if (argv.includes('--output-drive')) outputToLocal = false;
    else outputToLocal = sourceLocal;

    if (argv.includes('--no-bold-functloc-desc')) {
      boldFunclocDesc = false;
    }
  } else {
    log(`\n${bold('=== Final Clean — Google Drive → Google Drive (lsmw-style) ===')}`);
    logInfo(
      'Only REGIONAL 1 … 7 under the Drive root folder; each workbook: download → clean → upload (sequential).',
    );
    logInfo(
      'Auth: oauth.json + token.json (same OAuth flow as lsmw_equipment_v0).\n',
    );
    sourceLocal = false;
    samples = null;
    driveSamples = null;
    singleFileId = null;
    outputToLocal = false;
  }

  const needsDriveAuth =
    !sourceLocal || Boolean(singleFileId) || !outputToLocal;

  const tempDir = createTempDir();

  try {
    if (needsDriveAuth) {
      await initGoogleDrive(OAUTH_JSON);
    }

    let summaries = [];

    if (sourceLocal) {
      let excelFiles = [];
      logInfo('Source: Local mode, processing files from Data/ directory...');
      excelFiles = collectExcelFiles(DATA_ROOT).sort();

      if (!excelFiles.length) {
        logWarn('No Excel files found in Data/ directory.');
        return;
      }

      if (samples !== null) {
        excelFiles = excelFiles.slice(0, samples);
        logInfo(
          `Local mode: Processing ${excelFiles.length} sample file(s) (--samples ${samples}):`,
        );
        excelFiles.forEach((file, idx) => {
          logInfo(`  ${idx + 1}. ${file}`);
        });
      } else {
        logInfo(`Local mode: Processing all ${excelFiles.length} file(s) found:`);
        excelFiles.forEach((file, idx) => {
          logInfo(`  ${idx + 1}. ${file}`);
        });
      }

      logInfo('Processing files (final clean)...');
      summaries = await runWithConcurrency(excelFiles, DEFAULT_CONCURRENCY, async (file) => {
        const excelPath = typeof file === 'string' ? file : file.path || file;
        const driveFileInfo = typeof file === 'object' && file && file.driveFileInfo ? file.driveFileInfo : null;
        return processExcelFile(excelPath, driveFileInfo, {
          boldFunclocDesc,
        }).catch((err) => {
          let errorMsg = 'Unknown error';
          if (err) {
            if (typeof err === 'string') {
              errorMsg = err;
            } else if (err.message) {
              errorMsg = String(err.message);
            } else if (err.toString && typeof err.toString === 'function') {
              try {
                errorMsg = err.toString();
              } catch {
                errorMsg = 'Error (could not convert to string)';
              }
            } else {
              errorMsg = String(err);
            }
          }
          const displayPath = driveFileInfo ? driveFileInfo.path : excelPath;
          logError(`Failed to process ${displayPath}: ${errorMsg}`);
          return {
            excelPath: driveFileInfo ? driveFileInfo.path : excelPath,
            error: errorMsg,
            driveFileInfo,
          };
        });
      });

      logFinalProcessSummary(summaries);

      if (outputToLocal) {
        logInfo('Saving files to local Output/ directory...');

        const timestamp = getJakartaTimestamp();
        const outputFolderName = `Final_Clean_Results_${timestamp}`;
        const localOutputDir = path.join(OUTPUT_DIR, outputFolderName);

        fs.mkdirSync(localOutputDir, { recursive: true });
        logSuccess(`Created local output directory: ${localOutputDir}`);

        const filesToSave = summaries.filter((s) => s.outputPath && fs.existsSync(s.outputPath));

        let savedCount = 0;
        for (const result of filesToSave) {
          try {
            const fileName = outputSaveBasename(result);

            let targetDir = localOutputDir;
            const relativePath = result.driveFileInfo?.path
              ? result.driveFileInfo.path
              : path.relative(DATA_ROOT, result.excelPath);
            const pathParts = String(relativePath).split(/[\\/]/);

            if (pathParts.length > 1 && !pathParts[0].startsWith('.')) {
              const regionalDir = path.join(localOutputDir, pathParts[0]);
              fs.mkdirSync(regionalDir, { recursive: true });
              targetDir = regionalDir;
            }

            const targetPath = path.join(targetDir, fileName);
            fs.copyFileSync(result.outputPath, targetPath);
            logSuccess(`Saved: ${fileName} (to ${targetDir})`);
            savedCount += 1;
          } catch (err) {
            logError(`Failed to save ${result.excelPath}: ${err.message}`);
          }
        }

        const successCount = summaries.filter((s) => !s.error && s.outputPath).length;
        const errorCount = summaries.filter((s) => s.error).length;

        log(`\n${bold('=== Final Clean Local Summary ===')}`);
        log(`Files saved to: ${localOutputDir}`);
        log(`Files saved: ${savedCount}/${filesToSave.length}`);
        log(`Processed successfully: ${successCount}`);
        log(`Errors: ${errorCount}`);
        log('\nReview the files before uploading to Google Drive.');

        return;
      }

      logInfo('Creating output folder in Google Drive...');
      const timestamp = getJakartaTimestamp();
      const outputFolderName = `Final_Clean_Results_${timestamp}`;
      const sourceFolderId = extractFolderIdFromUrl(GOOGLE_DRIVE_FOLDER_URL);

      let outputFolder;
      try {
        outputFolder = await getOrCreateFolder(outputFolderName, sourceFolderId);
        logSuccess(`Created output folder: ${outputFolder.name}`);
      } catch (err) {
        logWarn(`Could not create folder: ${err.message}`);
        logInfo('Will attempt to upload directly to source folder...');
        outputFolder = { id: sourceFolderId, name: 'Source Folder' };
      }

      logInfo('Pre-creating regional folders...');
      const filesToUpload = summaries.filter((s) => s.outputPath && fs.existsSync(s.outputPath));
      const relativePaths = filesToUpload.map((r) => r.driveFileInfo?.path || r.excelPath);
      const regionalFolders = await ensureRegionalFolders(outputFolder.id, relativePaths, {
        logInfo,
        logWarn,
      });
      logInfo(`Pre-created ${regionalFolders.size} regional folder(s).`);

      logInfo('Uploading processed files to Google Drive...');
      const uploadConcurrency = Math.max(1, Math.min(3, DEFAULT_CONCURRENCY));
      if (uploadConcurrency < DEFAULT_CONCURRENCY) {
        logInfo(`Reduced upload concurrency to ${uploadConcurrency} to improve reliability`);
      }

      const uploadItems = filesToUpload.map((r) => ({
        localPath: r.outputPath,
        fileName: outputSaveBasename(r),
        relativePath: r.driveFileInfo?.path || r.excelPath,
      }));
      const uploadResults = await uploadToDriveWithRegionalFolders(uploadItems, outputFolder.id, {
        concurrency: uploadConcurrency,
        regionalFolders,
        logSuccess,
        logWarn,
        logError,
      });

      const successCount = uploadResults.filter((r) => r.success).length;
      const errorCount = uploadResults.filter((r) => !r.success).length;

      logSuccess(
        `\nUploaded ${successCount}/${uploadResults.length} processed file(s) to Google Drive.`,
      );
      if (errorCount > 0) {
        logWarn(`Upload errors: ${errorCount}`);
      }

      return;
    }

    if (singleFileId) {
      logInfo('Google Drive mode: Processing a single file from provided URL/ID...');
      const singleLocalPath = path.join(tempDir, `single_${Date.now()}.xlsx`);
      await downloadFile(singleFileId, singleLocalPath);
      const pseudoInfo = {
        id: singleFileId,
        name: path.basename(singleLocalPath),
        path: path.basename(singleLocalPath),
      };

      const summary = await processExcelSafely(singleLocalPath, pseudoInfo, {
        boldFunclocDesc,
      });
      unlinkQuiet(singleLocalPath);

      if (!summary.error && summary.outputPath && fs.existsSync(summary.outputPath)) {
        const uploadName = outputSaveBasename(summary);
        if (outputToLocal) {
          const ts = getJakartaTimestamp();
          const localRoot = path.join(OUTPUT_DIR, `Final_Clean_Results_${ts}`);
          fs.mkdirSync(localRoot, { recursive: true });
          const targetPath = path.join(localRoot, uploadName);
          fs.copyFileSync(summary.outputPath, targetPath);
          logSuccess(`Saved: ${uploadName}`);
        } else {
          const sourceRootId = extractFolderIdFromUrl(GOOGLE_DRIVE_FOLDER_URL);
          const outF = await getOrCreateFolder(
            `Final_Clean_Results_${getJakartaTimestamp()}`,
            sourceRootId,
          );
          await uploadXlsxReplacing(summary.outputPath, uploadName, outF.id);
          logSuccess(`Uploaded: ${uploadName}`);
        }
        unlinkQuiet(summary.outputPath);
      }

      logFinalProcessSummary([summary]);
      return;
    }

    logInfo(
      'Google Drive (lsmw-style): only REGIONAL 1–7; download → process → upload/save per workbook (sequential).',
    );
    const { summaries: regSummaries } = await runRegionalRootPipelineLikeLsmw({
      tempDir,
      boldFunclocDesc,
      outputToLocal,
      driveSamples,
    });
    logFinalProcessSummary(regSummaries);
  } catch (err) {
    logError(`Error: ${err.message}`);
    throw err;
  } finally {
    if (!outputToLocal) {
      cleanupTempDir(tempDir);
    } else {
      logInfo(`Temp directory kept for review: ${tempDir}`);
    }
  }
};

main().catch((err) => {
  logError(err);
  process.exitCode = 1;
});
