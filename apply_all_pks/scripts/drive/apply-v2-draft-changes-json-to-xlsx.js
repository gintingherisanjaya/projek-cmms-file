/**
 * Apply dendrogram / draft-mapping JSON (FUNCLOC moves) to Excel workbook(s) — v2.
 *
 * Same sheet shape as full-validation: header row via normalizeHeader match.
 * Default apply path matches the dendrogram JSON import: normalize + dedupe mappings like the webapp,
 * run the same compact-gap base pass, then `applyDraftMappingsToRows` (logical state), then write
 * FUNCLOC/DESC back to the sheet (with Excel seg1–2 + JSON tail merge). Use `--legacy-step-replay`
 * for the older per-step row-matching replay (ordered by JSON `step`).
 *
 * v2 extras: preserve full sheet row layout and header styling on write; after draft
 * replay, derive COST CENTER AFTER (10 chars: seg2 + STAS + last 2 digits of seg4 from
 * FUNCLOC AFTER) with a light gray fill; truncate EQUIPMENT GROUP AFTER to 10 chars;
 * FUNCLOC AFTER from JSON: keep hyphen segments 1–2 from the Excel row, segments 3+ from JSON
 * (site/plant in Excel e.g. PALM-7F08 vs dendrogram JSON PALM-2F01 preserved from file);
 * FUNCTLOC DESC → EQUIPMENT GROUP AFTER when that column is empty (≤10 chars, no spaces): `drive unit …` →
 * DRVUNIT (inactive red fill inherited from direct child EG when applicable); leading GEARMOTOR/GEARBOX/…/PIPE
 * → matching prefix; else conveyor/elevator/pump/air compressor → CONVEYOR/ELEVATOR/PUMP/AIRCOMP;
 * station rows (FUNCLOC AFTER with exactly 4 hyphen segments) get subtle gray on row cells
 * (skipping inactive red fills); REGIONAL / POM blanks filled when all non-empty values match.
 *
 * Actions:
 * - rename / move / move+rename / compact-gap: update matching FUNCLOC AFTER rows (and optional DESC).
 * - create: append a new row only when that FUNCLOC AFTER is not already on the sheet (dendrogram preserve-existing;
 *   otherwise skip create). New row: FUNCLOC AFTER + DESC; other columns blank (styled like template).
 * - delete: same as legacy apply — clone row to sheet bottom (archive), clear FUNCLOC on archive, tag DESC with
 *   `[DELETED — step N]`, clear the original row. After all steps, non-archive rows are packed & sorted by
 *   FUNCLOC AFTER (parent before children); delete-archive rows stay grouped below actives.
 * - duplicate FUNCLOC on sheet (same AFTER twice): first row is canonical for the draft; non-empty cells from the
 *   later row(s) are merged into empty cells on the first row (e.g. FUNCTIONAL LOCATION BEFORE), then duplicate
 *   rows are cleared — no duplicate archive row. Pack still drops legacy `[DUPLICATE FUNCLOC — dendrogram-parity]`
 *   DESC rows if present from older runs.
 *
 * Usage:
 *   node scripts/drive/apply-draft-changes-json-to-xlsx.js
 *     → Interactive (TTY only): JSON path (defaults to Reference/master-data-draft-mapping-1776163197277.json),
 *       Local vs Drive, samples/selection, output layout, options.
 *
 *   node scripts/drive/apply-draft-changes-json-to-xlsx.js \
 *     --json path/to/changes.json --input path/to/file.xlsx [--output path/out.xlsx] [--also-before] [--dry-run]
 *
 *   node scripts/drive/apply-draft-changes-json-to-xlsx.js \
 *     --json path/to/changes.json --local [--samples N] [--output-dir dir] [--beside-input]
 * 
 *   node scripts/drive/apply-draft-changes-json-to-xlsx.js \
 *     --json path/to/changes.json --drive [--drive-samples N] [--output-dir dir]
 *
 * Default batch output: Output/Draft-applied/… (mirrors Data/ or Drive relative paths).
 * Output filename matches the source workbook basename (e.g. foo.xlsx); --beside-input
 * writes next to the source (same folder + same name — overwrites that file when not --dry-run).
 * Exception: --drive-single-file temp download uses <name>_draft-applied.xlsx unless --output is set.
 *
 * Pipeline: by default each input workbook is first run through `check-v2-bah-jambi-level5-children.js`
 * (`preprocessBahJambiLevel5ForDraftApply` — fixed master + sidecar reports under a temp `level5-pre/` folder),
 * then draft JSON is applied to that fixed file; the final .xlsx is written to the chosen output layout.
 * Use `--skip-bah-jambi-preprocess` to read the original workbook directly (old behavior).
 */

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { select, checkbox, input, confirm } from '@inquirer/prompts';

import {
  findCostCenterCol,
  findHeaderRowAndColumns,
  normalizeHeader,
  toCleanString,
} from '../../utils/excelHelpers.js';
import { fillIsReddishPreserveFromXlsx, writeWorkbookWithStyles } from '../../utils/excelWriteWithStyles.js';
import { ensureDir, collectExcelFiles } from '../../utils/fileSystem.js';
import {
  createTempDir,
  cleanupTempDir,
  runWithConcurrency,
} from '../../utils/concurrencyHelpers.js';
import {
  listFilesRecursively,
  downloadFile,
  extractFileIdFromAnyUrl,
  extractFolderIdFromUrl,
} from '../../utils/googleDrive.js';
import { filterDriveSourceExcelFiles } from '../../utils/driveSourceFileFilter.js';
import {
  clearRowColumns,
  colToLetters,
  decodeRange,
  encodeCell,
  getCellValue,
  getCleanCellValue,
  readExcelSheetCompat,
  rowHasOriginalData as worksheetRowHasOriginalData,
  sanitizeInactiveRedOnEmptyTextCell,
  setCellString,
} from '../../utils/worksheetAdapter.js';
import { normalizeWhitespace } from '../../utils/normalization.js';
import { bold, log, logError, logInfo, logSuccess, logWarn } from '../../utils/logger.js';
import {
  normalizeImportedDraftChanges,
  compactGapFuncLocDraft,
  applyDraftMappingsToRows as applyDraftMappingsToRowsDendrogramParity,
} from '../../utils/draftMappingDendrogramParity.js';
import { preprocessBahJambiLevel5ForDraftApply } from './check-v2-bah-jambi-level5-children.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/** Default Bah Jambi fix map for level-5 preprocess (same folder as this script). */
const DEFAULT_BAH_JAMBI_FIX_MAP = path.join(__dirname, 'mappings', 'bah-jambi-fix-mapping.json');

const slugifyLevel5JobKey = (k) =>
  String(k ?? 'job')
    .trim()
    .replace(/[\\/]+/g, '_')
    .replace(/[^\w.-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 120) || 'job';

/** True when this file is executed directly (not imported). */
const isExecutedAsCli = () => path.resolve(process.argv[1] || '') === __filename;

/** Repo root (Compare-LPP): used to resolve relative `--json` paths from project root. */
const PROJECT_ROOT = path.join(__dirname, '..', '..');
const DATA_ROOT = path.join(PROJECT_ROOT, 'Data');
const REFERENCE_ROOT = path.join(PROJECT_ROOT, 'Reference');
/** Pinned default for interactive JSON picker (see Reference/master-data-draft-mapping-1776163197277.json). */
const DEFAULT_DRAFT_MAPPING_JSON = path.join(
  REFERENCE_ROOT,
  'master-data-draft-mapping-1776163197277.json',
);
const OUTPUT_DRAFT_ROOT = path.join(PROJECT_ROOT, 'Output', 'Draft-applied');
const GOOGLE_DRIVE_FOLDER_URL =
  'https://drive.google.com/drive/folders/1zERO-GM4g1THvrUeWW-JmNX2A1eHium8';
const SOURCE_DRIVE_OWNER_EMAIL = 'fahrurrozy4214@gmail.com';

const DEFAULT_CONCURRENCY = Math.max(
  1,
  Math.min(20, os.availableParallelism?.() ?? os.cpus().length ?? 8),
);

const SELECT_OPTS = { pageSize: 15 };
const CHECKBOX_OPTS = { pageSize: 20 };

const WANTED_HEADERS = [
  'FUNCTIONAL LOCATION AFTER',
  'FUNCTIONAL LOCATION BEFORE',
  'FUNCTLOC DESC. AFTER LEVEL 1,2,3',
  'FUNCTLOC DESC AFTRER LEVEL 1,2,3',
  'EQKTU AFTER LEVEL 2 dan 3',
  'COST CENTER AFTER',
  'COST CENTER',
  'EQUIPMENT GROUP AFTER',
  'REGIONAL',
  'POM',
];

/** Light gray fill (RGB) for derived COST CENTER, station-row shading, etc. */
const COST_CENTER_DERIVED_FILL_RGB = 'EDEDED';

/** FUNCLOC AFTER with exactly this many hyphen segments = stasiun row (e.g. PALM-2F01-0005-0001). */
const STATION_FUNCLOC_SEGMENT_COUNT = 4;

/** Max length for EQUIPMENT GROUP AFTER when writing v2 output. */
const EQUIPMENT_GROUP_AFTER_MAX_LEN = 10;

/** Inactive / warning EQUIPMENT GROUP AFTER fill (aligned with fillIsReddishPreserveFromXlsx / Excel light red). */
const INACTIVE_EQUIPMENT_GROUP_FILL_RGB = 'FFC7CE';

const EQUIPMENT_GROUP_DRVUNIT_VALUE = 'DRVUNIT';

/**
 * True when FUNCTLOC DESC (already uppercased) starts with the word PANEL (not e.g. PANELING).
 * Leading spaces ignored; after `PANEL` must be end or a non letter/digit.
 */
const isPanelLedDescUpper = (descUpper) => {
  const t = String(descUpper ?? '').trimStart();
  if (!t.startsWith('PANEL')) return false;
  if (t.length === 5) return true;
  return !/[A-Z0-9]/.test(t.charAt(5));
};

/** Rule 1: description starts with "drive unit" then space or end (case-insensitive). */
const isRule1DriveUnitLedDesc = (descUpperTrimStart) =>
  /^DRIVE UNIT(\s|$)/.test(String(descUpperTrimStart ?? ''));

/** Rule 2 prefixes (longest first); value = token without spaces, max 10 chars. */
const EQUIPMENT_GROUP_RULE2_PREFIXES = [
  'GEARMOTOR',
  'GEARBOX',
  'ELECTROMOTOR',
  'ACCESSORIES',
  'STRUCTURE',
  'PLATFORM',
  'PANEL',
  'BODY',
  'PIPE',
].sort((a, b) => b.length - a.length || a.localeCompare(b));

const startsWithRule2Prefix = (d) => {
  const desc = String(d ?? '');
  for (const token of EQUIPMENT_GROUP_RULE2_PREFIXES) {
    if (token === 'PANEL') {
      if (isPanelLedDescUpper(desc)) return 'PANEL';
      continue;
    }
    if (!desc.startsWith(token)) continue;
    if (desc.length === token.length) return token;
    const next = desc.charAt(token.length);
    if (!/[A-Z0-9]/.test(next)) return token;
  }
  return null;
};

/** Rule 3: word conveyor / elevator / pump / air compressor (only if rules 1–2 did not apply). */
const matchRule3EquipmentKeyword = (descUpper) => {
  const x = String(descUpper ?? '');
  if (/\bCONVEYOR\b/i.test(x)) return 'CONVEYOR';
  if (/\bELEVATOR\b/i.test(x)) return 'ELEVATOR';
  if (/\bPUMP\b/i.test(x)) return 'PUMP';
  if (/AIR\s*COMPRESSOR/i.test(x)) return 'AIRCOMP';
  return null;
};

const normalizeEquipmentGroupValue = (v) =>
  String(v ?? '')
    .replace(/\s+/g, '')
    .slice(0, EQUIPMENT_GROUP_AFTER_MAX_LEN);

const equipmentGroupCellIsSemanticallyEmpty = (worksheet, rowIndex, equipmentGroupCol) => {
  if (equipmentGroupCol === undefined) return false;
  const raw = getCellValue(worksheet, rowIndex, equipmentGroupCol);
  return String(raw ?? '')
    .replace(/\s+/g, '')
    .trim()
    .length === 0;
};

/**
 * From FUNCLOC AFTER like `PALM-1F02-0005-0001-0001`: take hyphen segments #2 and #4,
 * build fixed 10 chars: [seg2 up to 4 alnum, pad with 0] + `STAS` + last 2 digits from seg4.
 * Example → `1F02STAS01`.
 */
const deriveCostCenterFromFunclocAfter = (rawFuncloc) => {
  const s = toCleanString(rawFuncloc).toUpperCase();
  if (!s) return '';
  const parts = s.split('-').filter(Boolean);
  if (parts.length < 4) return '';
  const seg2 = parts[1] ?? '';
  const seg4 = parts[3] ?? '';
  const left = String(seg2)
    .replace(/[^A-Z0-9]/g, '')
    .slice(0, 4)
    .padEnd(4, '0');
  const digits = String(seg4).replace(/\D/g, '');
  const tail = (digits.slice(-2) || '00').padStart(2, '0').slice(-2);
  return `${left}STAS${tail}`.slice(0, 10);
};

const mergeCostCenterDerivedStyle = (existingS) => {
  const out = {};
  if (existingS?.font) out.font = { ...existingS.font };
  if (existingS?.alignment) out.alignment = { ...existingS.alignment };
  if (existingS?.border) out.border = { ...existingS.border };
  out.fill = { patternType: 'solid', fgColor: { rgb: COST_CENTER_DERIVED_FILL_RGB } };
  return out;
};

const setCostCenterDerivedCell = (worksheet, rowIndex, costCenterCol, value) => {
  if (costCenterCol === undefined) return;
  const addr = encodeCell({ r: rowIndex, c: costCenterCol });
  const prev = worksheet[addr] || {};
  worksheet[addr] = {
    ...prev,
    v: value,
    t: 's',
    s: mergeCostCenterDerivedStyle(prev.s),
  };
};

/** Subtle gray on existing row cells; skip reddish inactive fills (EQUIPMENT GROUP AFTER). */
const mergeStationRowSubtleStyle = (existingS) => {
  if (existingS?.fill && fillIsReddishPreserveFromXlsx(existingS.fill)) {
    return existingS;
  }
  const out = {};
  if (existingS?.font) out.font = { ...existingS.font };
  if (existingS?.alignment) out.alignment = { ...existingS.alignment };
  if (existingS?.border) out.border = { ...existingS.border };
  out.fill = { patternType: 'solid', fgColor: { rgb: COST_CENTER_DERIVED_FILL_RGB } };
  return out;
};

const paintStationRowSubtleGray = (worksheet, rowIndex, range) => {
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = encodeCell({ r: rowIndex, c });
    const prev = worksheet[addr];
    if (!prev) continue;
    if (prev.s?.fill && fillIsReddishPreserveFromXlsx(prev.s.fill)) continue;
    worksheet[addr] = {
      ...prev,
      s: mergeStationRowSubtleStyle(prev.s),
    };
  }
};

/** True when FUNCLOC AFTER is exactly 4 hyphen groups (stasiun), e.g. PALM-2F01-0005-0001. */
const isStationDepthFuncloc = (flNorm) => {
  const parts = String(flNorm ?? '')
    .trim()
    .toUpperCase()
    .split('-')
    .filter(Boolean);
  return parts.length === STATION_FUNCLOC_SEGMENT_COUNT;
};

/**
 * If every non-empty value in a column is identical, return that value; else null.
 * (All empty → null — do not invent a default.)
 */
const computeUniformNonEmptyColumnValue = (
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  colIndex,
) => {
  if (colIndex === undefined) return null;
  const distinct = new Set();
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
    const t = String(getCellValue(worksheet, r, colIndex) ?? '').trim();
    if (t) distinct.add(t);
  }
  if (distinct.size !== 1) return null;
  return [...distinct][0];
};

const fillBlankRegionalPomWhenUniform = (
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  regionalCol,
  pomCol,
) => {
  let regionalFilled = 0;
  let pomFilled = 0;
  const regCanon = computeUniformNonEmptyColumnValue(
    worksheet,
    headerRowIndex,
    range,
    originalLastCol,
    regionalCol,
  );
  const pomCanon = computeUniformNonEmptyColumnValue(
    worksheet,
    headerRowIndex,
    range,
    originalLastCol,
    pomCol,
  );
  if (!regCanon && !pomCanon) return { regionalFilled: 0, pomFilled: 0 };

  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
    if (regionalCol !== undefined && regCanon) {
      const cur = String(getCellValue(worksheet, r, regionalCol) ?? '').trim();
      if (!cur) {
        setCellString(worksheet, r, regionalCol, regCanon);
        regionalFilled += 1;
      }
    }
    if (pomCol !== undefined && pomCanon) {
      const cur = String(getCellValue(worksheet, r, pomCol) ?? '').trim();
      if (!cur) {
        setCellString(worksheet, r, pomCol, pomCanon);
        pomFilled += 1;
      }
    }
  }
  return { regionalFilled, pomFilled };
};

const applyV2DerivedColumns = ({
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  funclocAfterCol,
  costCenterCol,
  equipmentGroupCol,
}) => {
  let costCenterFilled = 0;
  let equipmentTruncated = 0;
  let stationRowsMarked = 0;
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;

    const fl = getCleanCellValue(worksheet, r, funclocAfterCol);
    if (fl && isStationDepthFuncloc(fl)) {
      paintStationRowSubtleGray(worksheet, r, range);
      stationRowsMarked += 1;
    }
    if (fl && costCenterCol !== undefined) {
      const cc = deriveCostCenterFromFunclocAfter(fl);
      if (cc) {
        setCostCenterDerivedCell(worksheet, r, costCenterCol, cc);
        costCenterFilled += 1;
      }
    }

    if (equipmentGroupCol !== undefined) {
      const egRaw = getCellValue(worksheet, r, equipmentGroupCol);
      const egStr = egRaw == null ? '' : String(egRaw);
      if (egStr.length > EQUIPMENT_GROUP_AFTER_MAX_LEN) {
        setCellString(worksheet, r, equipmentGroupCol, egStr.slice(0, EQUIPMENT_GROUP_AFTER_MAX_LEN));
        equipmentTruncated += 1;
      }
    }
  }
  return { costCenterFilled, equipmentTruncated, stationRowsMarked };
};

/** True when `childNorm` is exactly one hyphen segment deeper than `parentNorm` (same prefix). */
const isDirectChildFuncLocNorm = (parentNorm, childNorm) => {
  const pp = String(parentNorm ?? '')
    .split('-')
    .filter(Boolean);
  const cp = String(childNorm ?? '')
    .split('-')
    .filter(Boolean);
  if (!pp.length || cp.length !== pp.length + 1) return false;
  for (let i = 0; i < pp.length; i += 1) {
    if (pp[i] !== cp[i]) return false;
  }
  return true;
};

const equipmentGroupCellHasInactiveRedFill = (worksheet, rowIndex, equipmentGroupCol) => {
  if (equipmentGroupCol === undefined) return false;
  const addr = encodeCell({ r: rowIndex, c: equipmentGroupCol });
  return fillIsReddishPreserveFromXlsx(worksheet[addr]?.s?.fill);
};

const mergeInactiveEquipmentGroupFillPreservingTextStyle = (existingS) => {
  const out = {};
  if (existingS?.font) out.font = { ...existingS.font };
  if (existingS?.alignment) out.alignment = { ...existingS.alignment };
  if (existingS?.border) out.border = { ...existingS.border };
  out.fill = { patternType: 'solid', fgColor: { rgb: INACTIVE_EQUIPMENT_GROUP_FILL_RGB } };
  return out;
};

/**
 * FUNCTLOC DESC → EQUIPMENT GROUP AFTER (only when EG is empty: whitespace-only counts as empty).
 * Order: (1) leading `drive unit` → DRVUNIT + optional inactive red from direct children;
 * (2) leading GEARMOTOR / GEARBOX / … / PIPE → that token, max 10 chars no spaces;
 * (3) else contains conveyor / elevator / pump / air compressor → CONVEYOR | ELEVATOR | PUMP | AIRCOMP.
 */
const applyDriveUnitEquipmentGroupRules = ({
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  funclocAfterCol,
  funclocDescCol,
  equipmentGroupCol,
}) => {
  if (equipmentGroupCol === undefined || funclocDescCol === undefined) {
    return {
      driveUnitRowsSet: 0,
      driveUnitRowsRedInherit: 0,
      prefixRule2RowsSet: 0,
      keywordRule3RowsSet: 0,
    };
  }

  const rowMetas = [];
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
    const flNorm = normFuncloc(getCleanCellValue(worksheet, r, funclocAfterCol));
    const descRaw = String(getCellValue(worksheet, r, funclocDescCol) ?? '');
    rowMetas.push({ r, flNorm, descUpper: descRaw.toUpperCase() });
  }

  let driveUnitRowsSet = 0;
  let driveUnitRowsRedInherit = 0;
  let prefixRule2RowsSet = 0;
  let keywordRule3RowsSet = 0;

  for (const { r, flNorm, descUpper } of rowMetas) {
    if (!equipmentGroupCellIsSemanticallyEmpty(worksheet, r, equipmentGroupCol)) continue;

    const trimDesc = descUpper.trimStart();

    if (isRule1DriveUnitLedDesc(trimDesc)) {
      let childEgInactive = false;
      if (flNorm) {
        for (const { r: cr, flNorm: cfl } of rowMetas) {
          if (cr === r) continue;
          if (!isDirectChildFuncLocNorm(flNorm, cfl)) continue;
          if (equipmentGroupCellHasInactiveRedFill(worksheet, cr, equipmentGroupCol)) {
            childEgInactive = true;
            break;
          }
        }
      }

      driveUnitRowsSet += 1;
      const addr = encodeCell({ r, c: equipmentGroupCol });
      const prev = worksheet[addr] || {};
      const val = normalizeEquipmentGroupValue(EQUIPMENT_GROUP_DRVUNIT_VALUE);

      if (childEgInactive) {
        driveUnitRowsRedInherit += 1;
        worksheet[addr] = {
          ...prev,
          v: val,
          t: 's',
          s: mergeInactiveEquipmentGroupFillPreservingTextStyle(prev.s),
        };
      } else {
        setCellString(worksheet, r, equipmentGroupCol, val);
      }
      continue;
    }

    const rule2 = startsWithRule2Prefix(trimDesc);
    if (rule2) {
      const val = normalizeEquipmentGroupValue(rule2);
      setCellString(worksheet, r, equipmentGroupCol, val);
      prefixRule2RowsSet += 1;
      continue;
    }

    const rule3 = matchRule3EquipmentKeyword(descUpper);
    if (rule3) {
      setCellString(worksheet, r, equipmentGroupCol, normalizeEquipmentGroupValue(rule3));
      keywordRule3RowsSet += 1;
    }
  }

  return {
    driveUnitRowsSet,
    driveUnitRowsRedInherit,
    prefixRule2RowsSet,
    keywordRule3RowsSet,
  };
};

/** Re-apply empty EQUIPMENT GROUP cells so inactive red fill cannot linger (misleading vs moved rows). */
const sweepEmptyEquipmentGroupStripInactiveFill = (worksheet, headerRowIndex, range, equipmentGroupCol) => {
  if (equipmentGroupCol === undefined) return 0;
  let n = 0;
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    const raw = getCellValue(worksheet, r, equipmentGroupCol);
    if (raw != null && String(raw).trim() !== '') continue;
    setCellString(worksheet, r, equipmentGroupCol, '');
    n += 1;
  }
  return n;
};

const normFuncloc = (v) => toCleanString(v).toUpperCase();

/**
 * Build FUNCLOC AFTER for this workbook: first two hyphen groups from Excel, rest from JSON draft.
 * Example: Excel `PALM-7F08-0005-0002-…` + JSON `PALM-2F01-0005-0002-…` → `PALM-7F08-0005-0002-…`.
 * If Excel has fewer than 2 segments or JSON fewer than 3, returns `jsonAfter` unchanged.
 */
const mergeAfterFunclocKeepExcelFirstTwoSegments = (excelFlRaw, jsonAfterRaw) => {
  const jsonAfter = String(jsonAfterRaw ?? '').trim();
  const excelFl = String(excelFlRaw ?? '').trim();
  const excelParts = excelFl.split('-').filter(Boolean);
  const jsonParts = jsonAfter.split('-').filter(Boolean);
  if (excelParts.length >= 2 && jsonParts.length >= 3) {
    return [...excelParts.slice(0, 2), ...jsonParts.slice(2)].join('-');
  }
  return jsonAfter;
};

/** Hyphen segments for FUNCLOC ordering (matches webapp draft `compareFuncLoc`). */
const parseFuncLocPartsForSort = (value) =>
  String(value ?? '')
    .trim()
    .toUpperCase()
    .split('-')
    .filter(Boolean);

const compareFuncLocSegments = (a, b) => {
  const aa = parseFuncLocPartsForSort(a);
  const bb = parseFuncLocPartsForSort(b);
  const len = Math.max(aa.length, bb.length);
  for (let i = 0; i < len; i += 1) {
    const av = aa[i];
    const bv = bb[i];
    if (av == null) return -1;
    if (bv == null) return 1;
    const an = Number.parseInt(av, 10);
    const bn = Number.parseInt(bv, 10);
    const bothNumeric = Number.isFinite(an) && Number.isFinite(bn);
    if (bothNumeric) {
      if (an !== bn) return an - bn;
      continue;
    }
    const cmp = av.localeCompare(bv);
    if (cmp !== 0) return cmp;
  }
  return 0;
};

/** Empty FUNCLOC keys sort last so orphan rows stay below real FUNCLOC rows. */
const compareFuncLocSortKeys = (aRaw, bRaw) => {
  const a = String(aRaw ?? '').trim();
  const b = String(bRaw ?? '').trim();
  if (!a && !b) return 0;
  if (!a) return 1;
  if (!b) return -1;
  return compareFuncLocSegments(a, b);
};

const normDesc = (v) => normalizeWhitespace(String(v ?? ''));
const DESC_HEADERS = [
  'FUNCTLOC DESC. AFTER LEVEL 1,2,3',
  'FUNCTLOC DESC AFTRER LEVEL 1,2,3',
  'EQKTU AFTER LEVEL 2 dan 3',
];

/** Looser match for draft JSON vs Excel (e.g. `NO. 1` in UI vs `NO.1` in sheet). */
const normDescLoose = (v) => {
  let t = normDesc(v);
  if (!t) return '';
  t = t.replace(/\s+/g, '');
  return t.replace(/NO\.?(\d+)/gi, (_, d) => `NO.${d}`);
};
const getLooseDescAt = (worksheet, rowIndex, colIndex) =>
  normDescLoose(getCellValue(worksheet, rowIndex, colIndex));

const getParentId = (id) => {
  const parts = String(id).split('-').filter(Boolean);
  if (parts.length <= 1) return null;
  return parts.slice(0, -1).join('-');
};

const depthOf = (id) => String(id).split('-').filter(Boolean).length;
const isLoryDescException = (desc) => /\blory/i.test(String(desc).trim());
const toSuffix4 = (n) => String(n).padStart(4, '0');

const replaceFuncLocPrefix = (full, oldRoot, newRoot) => {
  if (full === oldRoot) return newRoot;
  if (full.startsWith(`${oldRoot}-`)) return `${newRoot}${full.slice(oldRoot.length)}`;
  return full;
};

const lastSegmentNum = (id) => {
  const seg = String(id).split('-').filter(Boolean).at(-1) ?? '';
  const n = Number.parseInt(seg, 10);
  return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
};

/**
 * Dendrogram-parity compact pass: renumber sibling suffixes from level 4 onward
 * (with LORY rows fixed) and return compacted FUNCLOC ids.
 */
const compactFuncLocRows = (rows) => {
  const working = rows.map((r) => ({ ...r }));
  const maxDepth = Math.max(0, ...working.map((r) => depthOf(r.funcLocAfter)));
  for (let childDepth = 4; childDepth <= maxDepth; childDepth += 1) {
    const parentDepth = childDepth - 1;
    if (parentDepth < 3) continue;

    const childrenByParent = new Map();
    for (const row of working) {
      const p = getParentId(row.funcLocAfter);
      if (!p) continue;
      if (depthOf(p) !== parentDepth) continue;
      const arr = childrenByParent.get(p);
      if (arr) arr.push(row);
      else childrenByParent.set(p, [row]);
    }

    const sortedParents = [...childrenByParent.keys()].sort((a, b) => {
      const da = depthOf(a);
      const db = depthOf(b);
      if (da !== db) return da - db;
      return a.localeCompare(b);
    });

    for (const parentId of sortedParents) {
      const children = childrenByParent.get(parentId);
      if (!children || children.length <= 1) continue;

      const fixedKids = children.filter((r) => isLoryDescException(r.desc));
      const rest = children.filter((r) => !isLoryDescException(r.desc));
      if (rest.length === 0) continue;
      rest.sort((a, b) => lastSegmentNum(a.funcLocAfter) - lastSegmentNum(b.funcLocAfter) || a.funcLocAfter.localeCompare(b.funcLocAfter));

      const reserved = new Set();
      for (const fixed of fixedKids) {
        const seg = String(fixed.funcLocAfter).split('-').filter(Boolean).at(-1) ?? '';
        const n = Number.parseInt(seg, 10);
        if (Number.isFinite(n)) reserved.add(n);
      }

      const suffixes = [];
      let n = 1;
      while (suffixes.length < rest.length) {
        if (!reserved.has(n)) suffixes.push(n);
        n += 1;
      }

      const remap = new Map();
      for (let i = 0; i < rest.length; i += 1) {
        const row = rest[i];
        const oldId = row.funcLocAfter;
        const newId = `${parentId}-${toSuffix4(suffixes[i])}`;
        if (oldId !== newId) remap.set(oldId, newId);
      }
      if (remap.size === 0) continue;

      // Two-pass remap to avoid collisions on overlapping prefixes.
      const toTemp = new Map();
      const tempToNew = new Map();
      let tmp = 0;
      for (const [oldId, newId] of remap.entries()) {
        const tempId = `${parentId}-TMP-COMPACT-${String(tmp).padStart(4, '0')}`;
        tmp += 1;
        toTemp.set(oldId, tempId);
        tempToNew.set(tempId, newId);
      }

      const oldRoots = [...toTemp.keys()];
      const tempRoots = [...tempToNew.keys()];
      for (const row of working) {
        let id = row.funcLocAfter;
        let touched = false;
        for (const oldId of oldRoots) {
          if (id === oldId || id.startsWith(`${oldId}-`)) {
            touched = true;
            break;
          }
        }
        if (!touched) continue;
        for (const [oldId, tempId] of toTemp.entries()) {
          id = replaceFuncLocPrefix(id, oldId, tempId);
        }
        row.funcLocAfter = id;
      }
      for (const row of working) {
        let id = row.funcLocAfter;
        let touched = false;
        for (const tempId of tempRoots) {
          if (id === tempId || id.startsWith(`${tempId}-`)) {
            touched = true;
            break;
          }
        }
        if (!touched) continue;
        for (const [tempId, newId] of tempToNew.entries()) {
          id = replaceFuncLocPrefix(id, tempId, newId);
        }
        row.funcLocAfter = id;
      }
    }
  }
  return working;
};

const buildCompactedAliasByRawFuncLoc = ({
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  funclocAfterCol,
  descCol,
}) => {
  const rawRows = [];
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
    rawRows.push({
      rowIndex: r,
      funcLocAfter: normFuncloc(getCleanCellValue(worksheet, r, funclocAfterCol)),
      desc: descCol !== undefined ? String(getCleanCellValue(worksheet, r, descCol) ?? '') : '',
    });
  }
  const compacted = compactFuncLocRows(rawRows);
  const alias = new Map();
  for (let i = 0; i < rawRows.length; i += 1) {
    const rawId = rawRows[i]?.funcLocAfter;
    const compactedId = compacted[i]?.funcLocAfter;
    if (!rawId || !compactedId) continue;
    if (rawId === compactedId) continue;
    if (!alias.has(compactedId)) alias.set(compactedId, new Set());
    alias.get(compactedId).add(rawId);
  }
  return alias;
};

const collectDescColumnIndexes = (colByHeader) => {
  const cols = [];
  for (const header of DESC_HEADERS) {
    const col = colByHeader.get(normalizeHeader(header));
    if (col === undefined) continue;
    if (!cols.includes(col)) cols.push(col);
  }
  return cols;
};

function updateWorksheetRef(worksheet, range) {
  worksheet['!ref'] = `${colToLetters(range.s.c)}${range.s.r + 1}:${colToLetters(range.e.c)}${range.e.r + 1}`;
}

const cloneCellForSnapshot = (cell) => {
  if (!cell) return null;
  const out = { ...cell };
  if (cell.s) out.s = { ...cell.s };
  return out;
};

/** True when primary cell has no meaningful text/number for PK-duplicate merge (dendrogram: keep first row, fill gaps from duplicate). */
const isEmptyishCellForPkMerge = (worksheet, r, c) => {
  const raw = getCellValue(worksheet, r, c);
  if (raw == null) return true;
  if (typeof raw === 'number') return false;
  return String(raw).trim() === '';
};

/**
 * Copy non-empty cells from duplicate row into primary where primary is empty (styles preserved).
 * Matches dendrogram "preserve existing" — one row per FUNCLOC; duplicate sheet rows contribute BEFORE/other columns.
 */
function mergeEmptyPrimaryCellsFromDuplicateRow(worksheet, primaryRow, dupRow, range) {
  let mergedCells = 0;
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    if (!isEmptyishCellForPkMerge(worksheet, primaryRow, c)) continue;
    if (isEmptyishCellForPkMerge(worksheet, dupRow, c)) continue;
    const dupAddr = encodeCell({ r: dupRow, c });
    const src = worksheet[dupAddr];
    if (!src) continue;
    const priAddr = encodeCell({ r: primaryRow, c });
    const packed = sanitizeInactiveRedOnEmptyTextCell(cloneCellForSnapshot(src));
    if (packed) worksheet[priAddr] = packed;
    mergedCells += 1;
  }
  return mergedCells;
}

function snapshotRowCells(worksheet, r, range) {
  const cells = new Map();
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = encodeCell({ r, c });
    const cell = worksheet[addr];
    if (cell) cells.set(c, cloneCellForSnapshot(cell));
  }
  return cells;
}

function clearRowCellsInRange(worksheet, r, range) {
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    delete worksheet[encodeCell({ r, c })];
  }
}

/** Matches DESC tagged by legacy delete archive (`[DELETED — step 12]`) or dendrogram-parity delete. */
const DELETED_ARCHIVE_DESC_RE =
  /\[DELETED\s*[—\-]\s*(?:step\s*\d+|dendrogram-parity)\]/i;

/** Matches DESC tagged when extra sheet rows share the same FUNCLOC AFTER as an earlier row (PK duplicate). */
const DUPLICATE_FUNCLOC_ARCHIVE_DESC_RE =
  /\[DUPLICATE\s+FUNCLOC\s*[—\-]\s*dendrogram-parity\]/i;

function rowIsDeleteArchiveRow(worksheet, r, descCols) {
  for (const c of descCols) {
    const v = String(getCellValue(worksheet, r, c) ?? '');
    if (DELETED_ARCHIVE_DESC_RE.test(v) || DUPLICATE_FUNCLOC_ARCHIVE_DESC_RE.test(v)) return true;
  }
  return false;
}

/** PK-duplicate archive rows (tagged in DESC) — omitted from final workbook; first active row wins. */
function rowIsDuplicateFuncLocArchiveRow(worksheet, r, descCols) {
  if (descCols.length === 0) return false;
  for (const c of descCols) {
    const v = String(getCellValue(worksheet, r, c) ?? '');
    if (DUPLICATE_FUNCLOC_ARCHIVE_DESC_RE.test(v)) return true;
  }
  return false;
}

/**
 * After replay: sort & pack **active** rows by FUNCLOC AFTER under the header; keep **delete archive** rows
 * (DESC contains `[DELETED — …]` / duplicate-FUNCLOC archive) grouped below, preserving archive append order.
 * Collapses multiple **active** rows with the same normalized FUNCLOC (keeps first by sort order).
 * Legacy rows still tagged `[DUPLICATE FUNCLOC — dendrogram-parity]` in DESC are omitted from output
 * (re-apply without that tag when source sheet is fixed).
 */
function packSortActiveRowsAndArchivesBelow({
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  funclocAfterCol,
  descCols = [],
}) {
  const activeEntries = [];
  const archiveEntries = [];
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    const fl = normFuncloc(getCleanCellValue(worksheet, r, funclocAfterCol));
    const hasData = worksheetRowHasOriginalData(worksheet, r, range, originalLastCol);
    if (!fl && !hasData) continue;
    const isArchive = descCols.length > 0 && rowIsDeleteArchiveRow(worksheet, r, descCols);
    if (isArchive) archiveEntries.push({ r, fl });
    else activeEntries.push({ r, fl });
  }

  const oldEnd = range.e.r;

  activeEntries.sort((a, b) => {
    const cmp = compareFuncLocSortKeys(a.fl, b.fl);
    if (cmp !== 0) return cmp;
    return a.r - b.r;
  });
  const seenActiveFl = new Set();
  const dedupedActive = [];
  let suppressedDuplicateActives = 0;
  for (const e of activeEntries) {
    if (e.fl) {
      if (seenActiveFl.has(e.fl)) {
        suppressedDuplicateActives += 1;
        continue;
      }
      seenActiveFl.add(e.fl);
    }
    dedupedActive.push(e);
  }

  const archiveEntriesKept = archiveEntries.filter(
    (e) => !rowIsDuplicateFuncLocArchiveRow(worksheet, e.r, descCols),
  );
  const droppedDuplicateArchiveRows = archiveEntries.length - archiveEntriesKept.length;

  const ordered = [];
  archiveEntriesKept.sort((a, b) => a.r - b.r);
  ordered.push(...dedupedActive, ...archiveEntriesKept);

  if (ordered.length === 0) {
    for (let r = headerRowIndex + 1; r <= oldEnd; r += 1) {
      clearRowCellsInRange(worksheet, r, range);
    }
    return {
      packedRows: 0,
      activeRows: 0,
      archiveRows: 0,
      suppressedDuplicateActives: 0,
      droppedDuplicateArchiveRows,
    };
  }

  const snapshots = ordered.map(({ r, fl }) => ({
    fl,
    cells: snapshotRowCells(worksheet, r, range),
    sourceRowIndex: r,
    rowDef: worksheet['!rows']?.[r],
  }));

  for (let r = headerRowIndex + 1; r <= oldEnd; r += 1) {
    clearRowCellsInRange(worksheet, r, range);
  }

  const rowDefs = worksheet['!rows'] ? [...worksheet['!rows']] : [];
  for (let i = 0; i < snapshots.length; i += 1) {
    const targetR = headerRowIndex + 1 + i;
    const snap = snapshots[i];
    for (const [c, cell] of snap.cells.entries()) {
      const packed = sanitizeInactiveRedOnEmptyTextCell(cloneCellForSnapshot(cell));
      if (packed) worksheet[encodeCell({ r: targetR, c })] = packed;
    }
    const def = snap.rowDef;
    if (def) rowDefs[targetR] = { ...def };
    else delete rowDefs[targetR];
  }
  const firstTail = headerRowIndex + 1 + snapshots.length;
  for (let r = firstTail; r <= oldEnd; r += 1) {
    delete rowDefs[r];
  }
  worksheet['!rows'] = rowDefs;

  range.e.r = oldEnd;
  updateWorksheetRef(worksheet, range);
  return {
    packedRows: snapshots.length,
    activeRows: dedupedActive.length,
    archiveRows: archiveEntriesKept.length,
    suppressedDuplicateActives,
    droppedDuplicateArchiveRows,
  };
}

/**
 * Append a new row at the bottom by cloning `sourceRowIndex` cell-by-cell (styles preserved).
 * @returns {number} 0-based row index of the new row
 */
function appendRowClonedFromSource(worksheet, range, sourceRowIndex) {
  const newRowIndex = range.e.r + 1;
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const srcAddr = encodeCell({ r: sourceRowIndex, c });
    const destAddr = encodeCell({ r: newRowIndex, c });
    const src = worksheet[srcAddr];
    if (src) worksheet[destAddr] = { ...src };
    else delete worksheet[destAddr];
  }
  range.e.r = newRowIndex;
  updateWorksheetRef(worksheet, range);
  return newRowIndex;
}

/**
 * Append a blank row (styles copied from template, values cleared) for `create` actions.
 * @returns {number} 0-based row index of the new row
 */
function appendBlankRowFromTemplate(worksheet, range, templateRowIndex) {
  const newRowIndex = range.e.r + 1;
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const srcAddr = encodeCell({ r: templateRowIndex, c });
    const destAddr = encodeCell({ r: newRowIndex, c });
    const src = worksheet[srcAddr];
    if (src) {
      const blank = { ...src, v: '', t: 's' };
      worksheet[destAddr] = sanitizeInactiveRedOnEmptyTextCell(blank) ?? blank;
    } else {
      delete worksheet[destAddr];
    }
  }
  range.e.r = newRowIndex;
  updateWorksheetRef(worksheet, range);
  return newRowIndex;
}

function findLastDataRowIndex(worksheet, headerRowIndex, range, originalLastCol) {
  for (let r = range.e.r; r > headerRowIndex; r -= 1) {
    if (worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) return r;
  }
  return headerRowIndex + 1;
}

function getTopFolderName(file) {
  const logicalPath = file?.path ?? file?.name ?? '';
  const top = String(logicalPath).split('/')[0] || '';
  return top.trim();
}

async function pickDriveFilesInteractive(driveFiles) {
  const mode = await select({
    message: 'Select Drive Excel files to process:',
    choices: [
      { name: 'All files', value: 'all' },
      { name: 'First N files (quick)', value: 'first-n' },
      { name: 'Pick folder(s) (top-level)', value: 'folders' },
      { name: 'Pick file(s)', value: 'files' },
    ],
    default: 'first-n',
    ...SELECT_OPTS,
  });

  if (mode === 'all') {
    return { files: driveFiles.slice(), summary: `All files (${driveFiles.length})` };
  }

  if (mode === 'first-n') {
    const limitChoice = await select({
      message: 'How many files to process?',
      choices: [
        { name: '1 (single file)', value: '1' },
        { name: '3', value: '3' },
        { name: '5', value: '5' },
        { name: '10', value: '10' },
        { name: '20', value: '20' },
        { name: '50', value: '50' },
      ],
      default: '5',
      ...SELECT_OPTS,
    });
    const limitN = Number(limitChoice);
    return {
      files: driveFiles.slice(0, limitN),
      summary: `First ${Math.min(limitN, driveFiles.length)} file(s)`,
    };
  }

  if (mode === 'folders') {
    const uniqueFolders = [...new Set(driveFiles.map(getTopFolderName).filter(Boolean))].sort(
      (a, b) => a.localeCompare(b),
    );
    const selectedFolderNames = await checkbox({
      message: 'Select folder(s) to process',
      choices: uniqueFolders.map((name) => ({ name, value: name })),
      ...CHECKBOX_OPTS,
    });
    const selectedSet = new Set((selectedFolderNames ?? []).filter(Boolean));
    const picked = selectedSet.size
      ? driveFiles.filter((f) => selectedSet.has(getTopFolderName(f)))
      : driveFiles.slice();
    return { files: picked, summary: `Folders (${selectedSet.size || 'all'})` };
  }

  let filtered = driveFiles.slice();
  if (filtered.length > 250) {
    const q = await input({
      message: `There are ${filtered.length} files. Filter by path keyword (empty = no filter):`,
      default: '',
    });
    const query = String(q ?? '')
      .trim()
      .toLowerCase();
    if (query) {
      filtered = filtered.filter((f) =>
        String(f.path ?? f.name ?? '')
          .toLowerCase()
          .includes(query),
      );
    }
  }

  filtered.sort((a, b) => {
    const nameA = String(a.path ?? a.name ?? a.id ?? '').toLowerCase();
    const nameB = String(b.path ?? b.name ?? b.id ?? '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  const selectedFileIds = await checkbox({
    message: 'Select file(s) to process',
    choices: filtered.map((f) => ({
      name: f.path ?? f.name ?? f.id,
      value: f.id,
    })),
    ...CHECKBOX_OPTS,
  });

  const selectedSet = new Set((selectedFileIds ?? []).filter(Boolean));
  const picked = selectedSet.size
    ? driveFiles.filter((f) => selectedSet.has(f.id))
    : driveFiles.slice();
  return { files: picked, summary: `Files (${selectedSet.size || driveFiles.length})` };
}

function collectJsonCandidates(max = 50) {
  const found = [];
  const walk = (dir, depth) => {
    if (found.length >= max || depth > 6 || !fs.existsSync(dir)) return;
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      if (found.length >= max) return;
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p, depth + 1);
      else if (e.isFile() && e.name.toLowerCase().endsWith('.json')) found.push(p);
    }
  };
  walk(path.join(PROJECT_ROOT, 'Output'), 0);
  walk(DATA_ROOT, 0);
  walk(REFERENCE_ROOT, 0);
  let sorted = found
    .map((p) => ({ p, m: fs.statSync(p).mtimeMs }))
    .sort((a, b) => b.m - a.m)
    .map(({ p }) => p);
  if (fs.existsSync(DEFAULT_DRAFT_MAPPING_JSON)) {
    const absRef = path.resolve(DEFAULT_DRAFT_MAPPING_JSON);
    sorted = sorted.filter((p) => path.resolve(p) !== absRef);
    sorted.unshift(absRef);
  }
  return sorted.slice(0, max);
}

const parseArgs = (argv) => {
  const out = {
    jsonPath: null,
    inputPath: null,
    outputPath: null,
    outputDir: null,
    besideInput: false,
    local: false,
    drive: false,
    driveSingleFile: false,
    source: null,
    samples: null,
    driveSamples: null,
    alsoBefore: false,
    precompactMatch: true,
    dryRun: false,
    help: false,
    legacyStepReplay: false,
    skipBahJambiPreprocess: false,
    bahJambiFixMap: null,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--json' && argv[i + 1]) {
      const raw = String(argv[i + 1]).trim();
      out.jsonPath = path.isAbsolute(raw) ? path.resolve(raw) : path.resolve(PROJECT_ROOT, raw);
      i += 1;
    } else if (a === '--input' && argv[i + 1]) {
      out.inputPath = path.resolve(argv[i + 1]);
      i += 1;
    } else if (a === '--output' && argv[i + 1]) {
      out.outputPath = path.resolve(argv[i + 1]);
      i += 1;
    } else if (a === '--output-dir' && argv[i + 1]) {
      out.outputDir = path.resolve(argv[i + 1]);
      i += 1;
    } else if (a === '--beside-input') {
      out.besideInput = true;
    } else if (a === '--local' || a === '-l') {
      out.local = true;
    } else if (a === '--drive') {
      out.drive = true;
    } else if (a === '--drive-single-file') {
      out.driveSingleFile = true;
    } else if (a === '--source' && argv[i + 1]) {
      out.source = String(argv[i + 1]).trim();
      i += 1;
    } else if (a === '--samples' && argv[i + 1]) {
      const n = parseInt(argv[i + 1], 10);
      if (!Number.isNaN(n) && n > 0) out.samples = n;
      i += 1;
    } else if (a === '--drive-samples' && argv[i + 1]) {
      const n = parseInt(argv[i + 1], 10);
      if (!Number.isNaN(n) && n > 0) out.driveSamples = n;
      i += 1;
    } else if (a === '--also-before') {
      out.alsoBefore = true;
    } else if (a === '--no-precompact-match') {
      out.precompactMatch = false;
    } else if (a === '--dry-run') {
      out.dryRun = true;
    } else if (a === '--legacy-step-replay') {
      out.legacyStepReplay = true;
    } else if (a === '--help' || a === '-h') {
      out.help = true;
    }
  }
  return out;
};

const loadDraftJson = (jsonPath) => {
  const raw = fs.readFileSync(jsonPath, 'utf8');
  const data = JSON.parse(raw);
  if (!data || !Array.isArray(data.changes)) {
    throw new Error('JSON must be an object with a "changes" array');
  }
  return data;
};

const orderChanges = (changes) =>
  [...changes].map((c, idx) => ({ ...c, _ord: idx })).sort((a, b) => {
    const sa = Number(a.step);
    const sb = Number(b.step);
    const ha = Number.isFinite(sa);
    const hb = Number.isFinite(sb);
    if (ha && hb && sa !== sb) return sa - sb;
    if (ha && !hb) return -1;
    if (!ha && hb) return 1;
    return a._ord - b._ord;
  });

const loadOrderedDraft = (jsonPath) => {
  const data = loadDraftJson(jsonPath);
  return { data, ordered: orderChanges(data.changes) };
};

/**
 * Apply draft the same way the webapp does on JSON import: compact-gap base + applyDraftMappingsToRows,
 * then sync worksheet rows (archive true deletes, append missing creates).
 */
function applyChangesDendrogramParity({
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  funclocAfterCol,
  funclocBeforeCol,
  descCol,
  descCols = [],
  normalizedMappings,
  alsoBefore,
}) {
  let replacedAfter = 0;
  let replacedBefore = 0;
  let descUpdates = 0;
  let deletesArchived = 0;
  let createsAppended = 0;
  const unmatchedSteps = [];
  const unmatchedDeleteSteps = [];

  const allColIndexes = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) allColIndexes.push(c);

  const sheetRows = [];
  const seenOrig = new Set();
  /** First data row index per normalized FUNCLOC AFTER (wins draft + keeps merged "before" fields). */
  const firstRowByOrig = new Map();
  /** Later sheet rows with same FUNCLOC AFTER — merged into primary then cleared (no duplicate archive row). */
  const duplicateFuncLocRowMeta = [];
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
    const flRaw = getCleanCellValue(worksheet, r, funclocAfterCol);
    const orig = normFuncloc(flRaw);
    if (!orig) continue;
    if (seenOrig.has(orig)) {
      duplicateFuncLocRowMeta.push({ dupRow: r, orig });
      logWarn(
        `Duplicate FUNCLOC identity ${orig} (row ${r + 1}) — first row wins; merging non-empty cells into primary, then clearing duplicate row`,
      );
      continue;
    }
    seenOrig.add(orig);
    firstRowByOrig.set(orig, r);
    const desc =
      descCol !== undefined ? String(getCellValue(worksheet, r, descCol) ?? '').trim() : '';
    const initialBeforeNorm =
      funclocBeforeCol !== undefined ? normFuncloc(getCleanCellValue(worksheet, r, funclocBeforeCol)) : '';
    sheetRows.push({
      rowIndex: r,
      originalFuncLocAfter: orig,
      excelFlRaw: String(getCellValue(worksheet, r, funclocAfterCol) ?? flRaw ?? '').trim(),
      initialBeforeNorm,
      desc,
    });
  }

  const sheetOriginalSet = new Set(sheetRows.map((s) => s.originalFuncLocAfter));
  const missingBeforeKeys = [];
  for (const m of normalizedMappings) {
    if (m.action === 'create') continue;
    if (!m.beforeFuncLoc) continue;
    if (!sheetOriginalSet.has(m.beforeFuncLoc)) missingBeforeKeys.push(m.beforeFuncLoc);
  }
  if (missingBeforeKeys.length) {
    logWarn(
      `Dendrogram-parity: ${missingBeforeKeys.length} mapping beforeFuncLoc key(s) not found on sheet (first 15): ${[
        ...new Set(missingBeforeKeys),
      ]
        .slice(0, 15)
        .join(', ')}${missingBeforeKeys.length > 15 ? ' …' : ''}`,
    );
  }

  let duplicatePkCellsMerged = 0;
  for (const { dupRow, orig } of duplicateFuncLocRowMeta) {
    const primaryRow = firstRowByOrig.get(orig);
    if (primaryRow === undefined) continue;
    duplicatePkCellsMerged += mergeEmptyPrimaryCellsFromDuplicateRow(
      worksheet,
      primaryRow,
      dupRow,
      range,
    );
  }
  if (duplicatePkCellsMerged > 0) {
    logInfo(
      `Dendrogram-parity: merged ${duplicatePkCellsMerged} cell(s) from duplicate FUNCLOC row(s) into primary row(s) (empty cells only).`,
    );
  }

  if (duplicateFuncLocRowMeta.length > 0 && descCol !== undefined) {
    for (const s of sheetRows) {
      s.desc = String(getCellValue(worksheet, s.rowIndex, descCol) ?? '').trim();
    }
  }

  const draftLike = sheetRows.map((s) => ({
    funcLocAfter: s.originalFuncLocAfter,
    originalFuncLocAfter: s.originalFuncLocAfter,
    desc: s.desc,
    hiddenBySimilarNumber: false,
    equipmentGroupAfterRedBg: false,
  }));

  const base = compactGapFuncLocDraft(draftLike);
  const finalRows = applyDraftMappingsToRowsDendrogramParity(base, normalizedMappings, {
    replaceExistingOnCreate: false,
  });

  const finalByOriginal = new Map();
  for (const row of finalRows) {
    if (!row.originalFuncLocAfter) continue;
    finalByOriginal.set(normFuncloc(row.originalFuncLocAfter), row);
  }

  const createdFinal = finalRows.filter((r) => !r.originalFuncLocAfter);

  const toDelete = sheetRows.filter((s) => !finalByOriginal.has(s.originalFuncLocAfter));
  toDelete.sort((a, b) => b.rowIndex - a.rowIndex);

  for (const s of toDelete) {
    const sourceRow = s.rowIndex;
    const archiveRow = appendRowClonedFromSource(worksheet, range, sourceRow);
    setCellString(worksheet, archiveRow, funclocAfterCol, '');
    if (funclocBeforeCol !== undefined) {
      setCellString(worksheet, archiveRow, funclocBeforeCol, '');
    }
    if (descCols.length > 0) {
      const tag = '[DELETED — dendrogram-parity]';
      for (const c of descCols) {
        const priorDesc = String(getCleanCellValue(worksheet, archiveRow, c) ?? '');
        setCellString(worksheet, archiveRow, c, priorDesc ? `${tag} ${priorDesc}` : tag);
      }
    }
    clearRowColumns(worksheet, sourceRow, allColIndexes);
    deletesArchived += 1;
  }

  for (const s of sheetRows) {
    const fr = finalByOriginal.get(s.originalFuncLocAfter);
    if (!fr) continue;
    const mergedFl = mergeAfterFunclocKeepExcelFirstTwoSegments(s.excelFlRaw, fr.funcLocAfter);
    setCellString(worksheet, s.rowIndex, funclocAfterCol, mergedFl);
    replacedAfter += 1;

    if (descCols.length > 0) {
      let touched = false;
      for (const c of descCols) {
        const cur = String(getCellValue(worksheet, s.rowIndex, c) ?? '').trim();
        if (cur !== fr.desc) touched = true;
        setCellString(worksheet, s.rowIndex, c, fr.desc);
      }
      if (touched) descUpdates += 1;
    }

    if (alsoBefore && funclocBeforeCol !== undefined) {
      if (s.initialBeforeNorm === s.originalFuncLocAfter) {
        setCellString(worksheet, s.rowIndex, funclocBeforeCol, mergedFl);
        replacedBefore += 1;
      }
    }
  }

  let dedupeDuplicateFuncLocRowsCleared = 0;
  const duplicateClears = [...duplicateFuncLocRowMeta].sort((a, b) => b.dupRow - a.dupRow);
  for (const { dupRow } of duplicateClears) {
    clearRowColumns(worksheet, dupRow, allColIndexes);
    dedupeDuplicateFuncLocRowsCleared += 1;
  }

  const rangeCursor = decodeRange(worksheet['!ref'] || 'A1:A1');
  const existingFlNorms = new Set();
  for (let r = headerRowIndex + 1; r <= rangeCursor.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, rangeCursor, originalLastCol)) continue;
    const fl = normFuncloc(getCleanCellValue(worksheet, r, funclocAfterCol));
    if (fl) existingFlNorms.add(fl);
  }

  const anchorRow = findLastDataRowIndex(worksheet, headerRowIndex, rangeCursor, originalLastCol);
  const anchorFlRaw = getCleanCellValue(worksheet, anchorRow, funclocAfterCol);
  let templateRow = anchorRow;
  for (const cr of createdFinal) {
    const want = normFuncloc(cr.funcLocAfter);
    if (!want || existingFlNorms.has(want)) continue;
    const mergedCreate = mergeAfterFunclocKeepExcelFirstTwoSegments(anchorFlRaw, cr.funcLocAfter);
    const newRow = appendBlankRowFromTemplate(worksheet, rangeCursor, templateRow);
    templateRow = newRow;
    setCellString(worksheet, newRow, funclocAfterCol, mergedCreate);
    if (descCols.length > 0) {
      for (const c of descCols) {
        setCellString(worksheet, newRow, c, cr.desc ?? '');
      }
    }
    existingFlNorms.add(want);
    createsAppended += 1;
  }

  const packResult = packSortActiveRowsAndArchivesBelow({
    worksheet,
    headerRowIndex,
    range: decodeRange(worksheet['!ref'] || 'A1:A1'),
    originalLastCol,
    funclocAfterCol,
    descCols,
  });

  return {
    replacedAfter,
    replacedBefore,
    descUpdates,
    unmatchedSteps,
    deletesArchived,
    createsAppended,
    unmatchedDeleteSteps,
    rowsPacked: packResult.packedRows,
    activeRowsPacked: packResult.activeRows,
    archiveRowsPacked: packResult.archiveRows,
    dedupeDuplicateFuncLocRowsCleared,
    duplicatePkCellsMerged,
    packDroppedDuplicateArchiveRows: packResult.droppedDuplicateArchiveRows ?? 0,
    packSuppressedDuplicateActives: packResult.suppressedDuplicateActives ?? 0,
    dendrogramParity: true,
  };
}

function findTargetRowsForBeforeFuncLoc({
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  funclocAfterCol,
  descCol,
  beforeKey,
  jsonDescBefore,
  precompactMatch,
}) {
  const exactRows = [];

  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
    const curAfter = normFuncloc(getCleanCellValue(worksheet, r, funclocAfterCol));
    if (curAfter === beforeKey) exactRows.push(r);
  }

  // Fast path: exact match exists, so alias fallback is not needed.
  if (exactRows.length > 0) {
    let targetRows = exactRows;
    if (descCol !== undefined && normDesc(jsonDescBefore)) {
      const looseJson = normDescLoose(jsonDescBefore);
      const descHits = exactRows.filter(
        (r) => getLooseDescAt(worksheet, r, descCol) === looseJson,
      );
      if (descHits.length > 0) {
        targetRows = descHits;
      } else {
        targetRows = exactRows;
      }
    }
    return targetRows;
  }

  const aliasRows = [];
  if (precompactMatch) {
    const aliasByCompacted = buildCompactedAliasByRawFuncLoc({
      worksheet,
      headerRowIndex,
      range,
      originalLastCol,
      funclocAfterCol,
      descCol,
    });
    const acceptedRawIds = aliasByCompacted.get(beforeKey) ?? null;
    if (acceptedRawIds) {
      for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
        if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
        const curAfter = normFuncloc(getCleanCellValue(worksheet, r, funclocAfterCol));
        if (acceptedRawIds.has(curAfter)) aliasRows.push(r);
      }
    }
  }

  const matchingRows = aliasRows;

  let targetRows = matchingRows;
  if (descCol !== undefined && normDesc(jsonDescBefore)) {
    const looseJson = normDescLoose(jsonDescBefore);
    const descHits = matchingRows.filter(
      (r) => getLooseDescAt(worksheet, r, descCol) === looseJson,
    );
    if (descHits.length > 0) {
      targetRows = descHits;
    } else {
      targetRows = matchingRows;
    }
  }
  return targetRows;
}

const applyChanges = ({
  worksheet,
  headerRowIndex,
  range,
  originalLastCol,
  funclocAfterCol,
  funclocBeforeCol,
  descCol,
  descCols = descCol === undefined ? [] : [descCol],
  orderedChanges,
  alsoBefore,
  precompactMatch = false,
}) => {
  let replacedAfter = 0;
  let replacedBefore = 0;
  let descUpdates = 0;
  let deletesArchived = 0;
  let createsAppended = 0;
  const unmatchedSteps = [];
  const unmatchedDeleteSteps = [];
  let lastCreateTemplateRow = null;

  const allColIndexes = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) allColIndexes.push(c);

  for (const ch of orderedChanges) {
    const action = String(ch.action || '');

    if (action === 'delete') {
      const beforeKey = normFuncloc(ch.beforeFuncLoc);
      if (!beforeKey) continue;

      const targetRows = findTargetRowsForBeforeFuncLoc({
        worksheet,
        headerRowIndex,
        range,
        originalLastCol,
        funclocAfterCol,
        descCol,
        beforeKey,
        jsonDescBefore: ch.descBefore,
        precompactMatch,
      });

      let hits = 0;
      for (const sourceRow of targetRows) {
        const archiveRow = appendRowClonedFromSource(worksheet, range, sourceRow);
        setCellString(worksheet, archiveRow, funclocAfterCol, '');
        if (funclocBeforeCol !== undefined) {
          setCellString(worksheet, archiveRow, funclocBeforeCol, '');
        }
        if (descCols.length > 0) {
          const tag = `[DELETED — step ${ch.step ?? ch._ord}]`;
          for (const c of descCols) {
            const priorDesc = String(getCleanCellValue(worksheet, archiveRow, c) ?? '');
            setCellString(worksheet, archiveRow, c, priorDesc ? `${tag} ${priorDesc}` : tag);
          }
        }
        clearRowColumns(worksheet, sourceRow, allColIndexes);
        deletesArchived += 1;
        hits += 1;
      }

      if (hits === 0) unmatchedDeleteSteps.push(ch.step ?? ch._ord);
      continue;
    }

    if (action === 'create') {
      const afterVal = String(ch.afterFuncLoc ?? '').trim();
      if (!afterVal) continue;

      if (lastCreateTemplateRow === null) {
        lastCreateTemplateRow = findLastDataRowIndex(
          worksheet,
          headerRowIndex,
          range,
          originalLastCol,
        );
      }
      const templateRow = lastCreateTemplateRow;
      const newRow = appendBlankRowFromTemplate(worksheet, range, templateRow);
      lastCreateTemplateRow = newRow;
      const templateFlRaw = getCleanCellValue(worksheet, templateRow, funclocAfterCol);
      const mergedCreate = mergeAfterFunclocKeepExcelFirstTwoSegments(templateFlRaw, afterVal);
      setCellString(worksheet, newRow, funclocAfterCol, mergedCreate);
      if (descCols.length > 0) {
        for (const c of descCols) {
          setCellString(worksheet, newRow, c, String(ch.descAfter ?? ''));
        }
      }
      createsAppended += 1;
      continue;
    }

    const beforeKey = normFuncloc(ch.beforeFuncLoc);
    const afterVal = String(ch.afterFuncLoc ?? '').trim();
    if (!beforeKey || !afterVal) continue;

    const jsonDescBefore = normDesc(ch.descBefore);
    const jsonDescAfter = normDesc(ch.descAfter);
    const touchDesc = descCols.length > 0 && jsonDescBefore !== jsonDescAfter;

    const targetRows = findTargetRowsForBeforeFuncLoc({
      worksheet,
      headerRowIndex,
      range,
      originalLastCol,
      funclocAfterCol,
      descCol,
      beforeKey,
      jsonDescBefore: ch.descBefore,
      precompactMatch,
    });

    const afterNorm = normFuncloc(afterVal);
    if (targetRows.length > 0 && afterNorm !== beforeKey) {
      const keep = new Set(targetRows);
      const victims = [];
      for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
        if (!worksheetRowHasOriginalData(worksheet, r, range, originalLastCol)) continue;
        if (keep.has(r)) continue;
        const cur = normFuncloc(getCleanCellValue(worksheet, r, funclocAfterCol));
        if (cur === afterNorm) victims.push(r);
      }
      for (const sourceRow of victims) {
        const archiveRow = appendRowClonedFromSource(worksheet, range, sourceRow);
        setCellString(worksheet, archiveRow, funclocAfterCol, '');
        if (funclocBeforeCol !== undefined) setCellString(worksheet, archiveRow, funclocBeforeCol, '');
        if (descCols.length > 0) {
          const tag = `[REPLACED — step ${ch.step ?? ch._ord}]`;
          for (const c of descCols) {
            const priorDesc = String(getCleanCellValue(worksheet, archiveRow, c) ?? '');
            setCellString(worksheet, archiveRow, c, priorDesc ? `${tag} ${priorDesc}` : tag);
          }
        }
        clearRowColumns(worksheet, sourceRow, allColIndexes);
        deletesArchived += 1;
      }
    }

    let hits = 0;
    for (const r of targetRows) {
      const excelAfterRaw = getCleanCellValue(worksheet, r, funclocAfterCol);
      const mergedAfter = mergeAfterFunclocKeepExcelFirstTwoSegments(excelAfterRaw, afterVal);
      setCellString(worksheet, r, funclocAfterCol, mergedAfter);
      replacedAfter += 1;
      hits += 1;

      if (alsoBefore && funclocBeforeCol !== undefined) {
        const curBefore = normFuncloc(getCleanCellValue(worksheet, r, funclocBeforeCol));
        if (curBefore === beforeKey) {
          setCellString(worksheet, r, funclocBeforeCol, mergedAfter);
          replacedBefore += 1;
        }
      }

      if (touchDesc) {
        for (const c of descCols) {
          setCellString(worksheet, r, c, String(ch.descAfter ?? ''));
        }
        descUpdates += 1;
      }
    }

    if (hits === 0) unmatchedSteps.push(ch.step ?? ch._ord);
  }

  const packResult = packSortActiveRowsAndArchivesBelow({
    worksheet,
    headerRowIndex,
    range,
    originalLastCol,
    funclocAfterCol,
    descCols,
  });

  return {
    replacedAfter,
    replacedBefore,
    descUpdates,
    unmatchedSteps,
    deletesArchived,
    createsAppended,
    unmatchedDeleteSteps,
    rowsPacked: packResult.packedRows,
    activeRowsPacked: packResult.activeRows,
    archiveRowsPacked: packResult.archiveRows,
    dedupeDuplicateFuncLocRowsCleared: 0,
    duplicatePkCellsMerged: 0,
    packDroppedDuplicateArchiveRows: packResult.droppedDuplicateArchiveRows ?? 0,
    packSuppressedDuplicateActives: packResult.suppressedDuplicateActives ?? 0,
  };
};

/** Temp name used for --drive-single-file download before apply (not a real source basename). */
const DRIVE_SINGLE_TEMP_BASENAME = 'drive-single-input';

function resolveOutputPath({
  explicitOutputPath,
  inputPath,
  outputBasenameSourcePath,
  besideInput,
  outputDir,
  dataRoot,
  draftOutRoot,
  driveFile,
}) {
  if (explicitOutputPath) return explicitOutputPath;
  const ext = path.extname(inputPath) || '.xlsx';
  const nameSrc = outputBasenameSourcePath ?? inputPath;
  const nameExt = path.extname(nameSrc) || ext;
  const base =
    path.basename(nameSrc, nameExt) || path.basename(inputPath, ext) || path.basename(inputPath);
  const useDraftAppliedSuffix = base.toLowerCase() === DRIVE_SINGLE_TEMP_BASENAME;
  const outFileName = useDraftAppliedSuffix ? `${base}_draft-applied${ext}` : `${base}${ext}`;
  if (besideInput) {
    const besideDir = outputBasenameSourcePath ? path.dirname(outputBasenameSourcePath) : path.dirname(inputPath);
    return path.join(besideDir, outFileName);
  }
  const targetRoot = outputDir ?? draftOutRoot;
  const relSource = outputBasenameSourcePath ?? inputPath;
  const rel = driveFile
    ? driveFile.path.replace(/\//g, path.sep)
    : path.relative(dataRoot, relSource);
  const subDir = path.dirname(rel);
  const outSub = subDir && subDir !== '.' ? path.join(targetRoot, subDir) : targetRoot;
  return path.join(outSub, outFileName);
}

/**
 * Level-5 Bah Jambi preprocess into a temp folder, then apply v2 draft JSON to that workbook.
 * @param {{ sharedWorkTempDir?: string | null, level5JobKey?: string }} opts
 */
async function runApplyDraftWithLevel5Preprocess({
  jsonPath,
  inputPath,
  outputPath,
  outputBasenameSourcePath,
  alsoBefore,
  precompactMatch,
  dryRun,
  quietJsonLog,
  legacyStepReplay,
  skipBahJambiPreprocess,
  bahJambiFixMap,
  sharedWorkTempDir,
  level5JobKey,
}) {
  if (skipBahJambiPreprocess) {
    return applyDraftToOneFile({
      jsonPath,
      inputPath,
      outputPath,
      alsoBefore,
      precompactMatch,
      dryRun,
      quietJsonLog,
      legacyStepReplay,
      outputBasenameSourcePath,
    });
  }

  const ownedChainTemp = sharedWorkTempDir ? null : createTempDir();
  const chainTemp = sharedWorkTempDir ?? ownedChainTemp;
  const key = `${level5JobKey ?? ''}_${outputBasenameSourcePath ?? inputPath}`;
  const level5Out = path.join(chainTemp, 'level5-pre', slugifyLevel5JobKey(key));
  ensureDir(level5Out);

  let cleanupDl = null;
  try {
    logInfo(`Bah Jambi level-5 preprocess (temp) → ${level5Out}`);
    const prep = await preprocessBahJambiLevel5ForDraftApply({
      source: inputPath,
      tempOutDir: level5Out,
      fixMapPath: bahJambiFixMap ?? DEFAULT_BAH_JAMBI_FIX_MAP,
      outputPrefix: path.basename(
        outputBasenameSourcePath ?? inputPath,
        path.extname(outputBasenameSourcePath ?? inputPath),
      ),
    });
    cleanupDl = prep.cleanupDownloadDir ?? null;
    return await applyDraftToOneFile({
      jsonPath,
      inputPath: prep.outputXlsx,
      outputPath,
      alsoBefore,
      precompactMatch,
      dryRun,
      quietJsonLog,
      legacyStepReplay,
      outputBasenameSourcePath: outputBasenameSourcePath ?? inputPath,
    });
  } finally {
    if (cleanupDl) cleanupDl();
    if (ownedChainTemp) cleanupTempDir(ownedChainTemp);
  }
}

async function applyDraftToOneFile({
  jsonPath,
  inputPath,
  outputPath,
  alsoBefore,
  precompactMatch = true,
  dryRun,
  orderedChanges = null,
  quietJsonLog = false,
  legacyStepReplay = false,
  normalizedMappingsOverride = null,
  outputBasenameSourcePath = null,
}) {
  const logStem = () => path.basename(outputBasenameSourcePath ?? inputPath);
  const logFullLabel = () => outputBasenameSourcePath ?? inputPath;

  const data = loadDraftJson(jsonPath);
  let ordered = null;
  let normalizedMappings = null;
  if (legacyStepReplay) {
    ordered = orderedChanges ?? orderChanges(data.changes);
    if (!quietJsonLog) {
      logInfo(
        `JSON: totalChanged=${data.totalChanged ?? '?'}, linear=${data.linear ?? '?'}, changes=${ordered.length}, changedAt=${data.changedAt ?? 'n/a'} (legacy --legacy-step-replay)`,
      );
    }
  } else {
    normalizedMappings = normalizedMappingsOverride ?? normalizeImportedDraftChanges(data.changes);
    if (!quietJsonLog) {
      logInfo(
        `JSON: totalChanged=${data.totalChanged ?? '?'}, linear=${data.linear ?? '?'}, mappings=${normalizedMappings.length} (dendrogram-parity), changedAt=${data.changedAt ?? 'n/a'}`,
      );
    }
  }

  const { workbook, worksheet } = await readExcelSheetCompat(inputPath);
  const found = findHeaderRowAndColumns(worksheet, WANTED_HEADERS, {});
  if (!found) {
    throw new Error(`Could not find header row containing one of: ${WANTED_HEADERS.join(', ')}`);
  }

  const { headerRowIndex, colByHeader } = found;
  const funclocAfterCol = colByHeader.get(normalizeHeader('FUNCTIONAL LOCATION AFTER'));
  if (funclocAfterCol === undefined) {
    throw new Error('FUNCTIONAL LOCATION AFTER column not found');
  }

  const funclocBeforeCol = colByHeader.get(normalizeHeader('FUNCTIONAL LOCATION BEFORE'));
  const descCols = collectDescColumnIndexes(colByHeader);
  const descCol = descCols[0];
  const costCenterCol = findCostCenterCol(colByHeader);
  const equipmentGroupCol = colByHeader.get(normalizeHeader('EQUIPMENT GROUP AFTER'));
  const funclocDescCol =
    colByHeader.get(normalizeHeader('FUNCTLOC DESC. AFTER LEVEL 1,2,3')) ??
    colByHeader.get(normalizeHeader('FUNCTLOC DESC AFTRER LEVEL 1,2,3'));
  const regionalCol = colByHeader.get(normalizeHeader('REGIONAL'));
  const pomCol = colByHeader.get(normalizeHeader('POM'));

  const originalLastCol = Math.max(...colByHeader.values(), 0);
  const range = decodeRange(worksheet['!ref'] || 'A1:A1');

  logInfo(
    `${logFullLabel()}: header row ${headerRowIndex + 1}; FUNCLOC AFTER col ${funclocAfterCol + 1}` +
      (funclocBeforeCol !== undefined ? `; BEFORE col ${funclocBeforeCol + 1}` : '') +
      (descCols.length > 0 ? `; DESC col(s) ${descCols.map((c) => c + 1).join(', ')}` : ' (no DESC column)') +
      (costCenterCol !== undefined ? `; COST CENTER col ${costCenterCol + 1}` : '') +
      (equipmentGroupCol !== undefined ? `; EQUIPMENT GROUP AFTER col ${equipmentGroupCol + 1}` : '') +
      (funclocDescCol !== undefined ? `; FUNCTLOC DESC AFTER L1–3 col ${funclocDescCol + 1}` : '') +
      (regionalCol !== undefined ? `; REGIONAL col ${regionalCol + 1}` : '') +
      (pomCol !== undefined ? `; POM col ${pomCol + 1}` : ''),
  );

  const stats = legacyStepReplay
    ? applyChanges({
        worksheet,
        headerRowIndex,
        range,
        originalLastCol,
        funclocAfterCol,
        funclocBeforeCol,
        descCol,
        descCols,
        orderedChanges: ordered,
        alsoBefore,
        precompactMatch,
      })
    : applyChangesDendrogramParity({
        worksheet,
        headerRowIndex,
        range,
        originalLastCol,
        funclocAfterCol,
        funclocBeforeCol,
        descCol,
        descCols,
        normalizedMappings,
        alsoBefore,
      });

  const rangeFinal = decodeRange(worksheet['!ref'] || 'A1:A1');
  const v2Derived = applyV2DerivedColumns({
    worksheet,
    headerRowIndex,
    range: rangeFinal,
    originalLastCol,
    funclocAfterCol,
    costCenterCol,
    equipmentGroupCol,
  });
  logInfo(
    `${logStem()}: v2 derived — station rows shaded (4-seg FUNCLOC): ${v2Derived.stationRowsMarked}` +
      `; COST CENTER cells set: ${v2Derived.costCenterFilled}` +
      `; EQUIPMENT GROUP truncated (>10): ${v2Derived.equipmentTruncated}`,
  );

  const rangeAfterDerived = decodeRange(worksheet['!ref'] || 'A1:A1');
  const driveUnitEg = applyDriveUnitEquipmentGroupRules({
    worksheet,
    headerRowIndex,
    range: rangeAfterDerived,
    originalLastCol,
    funclocAfterCol,
    funclocDescCol,
    equipmentGroupCol,
  });
  if (driveUnitEg.driveUnitRowsSet > 0) {
    logInfo(
      `${logStem()}: DRIVE UNIT-led FUNCTLOC DESC → EQUIPMENT GROUP AFTER ${EQUIPMENT_GROUP_DRVUNIT_VALUE}: ` +
        `${driveUnitEg.driveUnitRowsSet} (inactive red fill inherited from child EG: ${driveUnitEg.driveUnitRowsRedInherit})`,
    );
  }
  if (driveUnitEg.prefixRule2RowsSet > 0) {
    logInfo(
      `${logStem()}: FUNCTLOC DESC prefix rule (GEARMOTOR/GEARBOX/…/PIPE) → EQUIPMENT GROUP AFTER: ${driveUnitEg.prefixRule2RowsSet} row(s)`,
    );
  }
  if (driveUnitEg.keywordRule3RowsSet > 0) {
    logInfo(
      `${logStem()}: FUNCTLOC DESC keyword rule (CONVEYOR/ELEVATOR/PUMP/AIR COMPRESSOR) → EQUIPMENT GROUP AFTER: ${driveUnitEg.keywordRule3RowsSet} row(s)`,
    );
  }

  const rpFill = fillBlankRegionalPomWhenUniform(
    worksheet,
    headerRowIndex,
    rangeAfterDerived,
    originalLastCol,
    regionalCol,
    pomCol,
  );
  if (rpFill.regionalFilled > 0 || rpFill.pomFilled > 0) {
    logInfo(
      `${logStem()}: REGIONAL/POM uniform fill — REGIONAL: ${rpFill.regionalFilled}, POM: ${rpFill.pomFilled}`,
    );
  }

  const egSweep = sweepEmptyEquipmentGroupStripInactiveFill(
    worksheet,
    headerRowIndex,
    rangeAfterDerived,
    equipmentGroupCol,
  );
  if (equipmentGroupCol !== undefined && egSweep > 0) {
    logInfo(`${logStem()}: cleared inactive-style fill on ${egSweep} empty EQUIPMENT GROUP cell(s)`);
  }

  logSuccess(
    `${logStem()}: FUNCLOC AFTER cells updated: ${stats.replacedAfter}` +
      (alsoBefore ? `; BEFORE: ${stats.replacedBefore}` : '') +
      (stats.descUpdates ? `; DESC: ${stats.descUpdates}` : '') +
      (stats.deletesArchived ? `; deletes archived (packed active + FUNCLOC sort): ${stats.deletesArchived}` : '') +
      (stats.createsAppended ? `; rows created: ${stats.createsAppended}` : '') +
      (stats.rowsPacked != null
        ? `; rows laid out: ${stats.rowsPacked} (active ${stats.activeRowsPacked ?? '?'}, archive ${stats.archiveRowsPacked ?? '?'})`
        : '') +
      (stats.duplicatePkCellsMerged
        ? `; duplicate PK merge (cells into primary): ${stats.duplicatePkCellsMerged}`
        : '') +
      (stats.dedupeDuplicateFuncLocRowsCleared
        ? `; duplicate FUNCLOC rows cleared: ${stats.dedupeDuplicateFuncLocRowsCleared}`
        : '') +
      (stats.packDroppedDuplicateArchiveRows
        ? `; duplicate archive rows omitted from output: ${stats.packDroppedDuplicateArchiveRows}`
        : '') +
      (stats.packSuppressedDuplicateActives
        ? `; pack deduped active FUNCLOC: ${stats.packSuppressedDuplicateActives}`
        : ''),
  );

  if (stats.unmatchedSteps.length) {
    logWarn(
      `Steps with zero matching rows: ${stats.unmatchedSteps.slice(0, 20).join(', ')}${
        stats.unmatchedSteps.length > 20 ? ' …' : ''
      }`,
    );
  }
  if (stats.unmatchedDeleteSteps?.length) {
    logWarn(
      `Delete steps with zero matching rows: ${stats.unmatchedDeleteSteps.slice(0, 20).join(', ')}${
        stats.unmatchedDeleteSteps.length > 20 ? ' …' : ''
      }`,
    );
  }

  if (dryRun) {
    logInfo('Dry run: no file written for this workbook.');
    return { stats, outputPath: null, skippedWrite: true };
  }

  ensureDir(path.dirname(outputPath));
  await writeWorkbookWithStyles(workbook, worksheet, outputPath, { preserveSheetLayout: true });
  logSuccess(`Wrote ${outputPath}`);
  return { stats, outputPath, skippedWrite: false };
}

async function runInteractiveDraftApply() {
  log(`\n${bold('=== Apply draft JSON → Excel (Interactive) ===')}`);
  log('↑/↓ move • Enter select • every CLI flag can be set below\n');

  const candidates = collectJsonCandidates(45);
  const manualLabel = 'Type path manually…';
  const jsonChoice = await select({
    message: 'Draft changes JSON (--json)',
    choices: [
      { name: manualLabel, value: '__manual__' },
      ...candidates.map((p) => ({
        name: path.relative(process.cwd(), p) || p,
        value: p,
      })),
    ],
    default: candidates[0] ?? '__manual__',
    ...SELECT_OPTS,
  });

  let jsonPath =
    jsonChoice === '__manual__'
      ? path.resolve(
          String(
            await input({
              message: 'Absolute or relative path to JSON',
              validate: (v) => {
                const p = path.resolve(String(v ?? '').trim());
                return fs.existsSync(p) ? true : 'File not found';
              },
            }),
          ).trim(),
        )
      : jsonChoice;

  jsonPath = path.resolve(jsonPath);
  if (!fs.existsSync(jsonPath)) {
    logError(`JSON not found: ${jsonPath}`);
    process.exit(1);
  }

  const sourceMode = await select({
    message: 'Excel input source',
    choices: [
      { name: 'Local batch — .xlsx under Data/ (--local)', value: 'local' },
      { name: 'Local single file (--input [--output])', value: 'local-single' },
      { name: 'Google Drive batch — folder (--drive)', value: 'drive' },
      { name: 'Google Drive single file — paste link or file ID', value: 'drive-single' },
    ],
    default: 'local',
    ...SELECT_OPTS,
  });

  let driveFilesToProcess = [];
  let downloadedJobs = [];
  let tempDirForSingle = null;

  if (sourceMode === 'local') {
    const sampleChoice = await select({
      message: 'How many local files? (--samples)',
      choices: [
        { name: '1 (single sample)', value: '1' },
        { name: '3', value: '3' },
        { name: '5', value: '5' },
        { name: '10', value: '10' },
        { name: '20', value: '20' },
        { name: '50', value: '50' },
        { name: 'All files', value: 'all' },
      ],
      default: '1',
      ...SELECT_OPTS,
    });
    const samples = sampleChoice === 'all' ? null : parseInt(sampleChoice, 10);

    let excelFiles = collectExcelFiles(DATA_ROOT).sort();
    if (!excelFiles.length) {
      logWarn('No Excel files under Data/.');
      process.exit(1);
    }
    if (samples !== null) excelFiles = excelFiles.slice(0, samples);
    logInfo(`Local: ${excelFiles.length} file(s) selected.`);
    driveFilesToProcess = excelFiles.map((p) => ({
      localPath: p,
      driveFile: null,
      explicitOutputPath: null,
    }));
  } else if (sourceMode === 'local-single') {
    const inputXlsx = path.resolve(
      String(
        await input({
          message: 'Path to input .xlsx (--input)',
          validate: (v) => {
            const p = path.resolve(String(v ?? '').trim());
            if (!p.toLowerCase().endsWith('.xlsx')) return 'Must be a .xlsx file';
            return fs.existsSync(p) ? true : 'File not found';
          },
        }),
      ).trim(),
    );
    const setOut = await confirm({
      message: 'Set explicit output .xlsx (--output)? (No = choose mirror / beside / output-dir next)',
      default: false,
    });
    let explicitOutputPath = null;
    if (setOut) {
      explicitOutputPath = path.resolve(
        String(
          await input({
            message: 'Output .xlsx path (--output)',
            validate: (v) => {
              const p = path.resolve(String(v ?? '').trim());
              if (!p.toLowerCase().endsWith('.xlsx')) return 'Must end with .xlsx';
              const dir = path.dirname(p);
              return fs.existsSync(dir) ? true : 'Parent directory must exist';
            },
          }),
        ).trim(),
      );
    }
    driveFilesToProcess = [{ localPath: inputXlsx, driveFile: null, explicitOutputPath }];
  } else if (sourceMode === 'drive-single') {
    const urlOrId = String(
      await input({
        message: 'Google Drive file URL or file ID (single workbook)',
        validate: (v) => {
          const id = extractFileIdFromAnyUrl(String(v ?? '').trim());
          return id ? true : 'Could not parse a file id from that input';
        },
      }),
    ).trim();
    const fileId = extractFileIdFromAnyUrl(urlOrId);
    tempDirForSingle = createTempDir();
    const localPath = path.join(tempDirForSingle, 'drive-input.xlsx');
    logInfo('Downloading file from Google Drive…');
    await downloadFile(fileId, localPath);
    driveFilesToProcess = [
      {
        localPath,
        driveFile: { id: fileId, path: 'drive-input.xlsx', name: 'drive-input.xlsx' },
        explicitOutputPath: null,
      },
    ];
    return await finishInteractiveAfterSource({
      jsonPath,
      jobs: driveFilesToProcess,
      tempDir: tempDirForSingle,
    });
  } else {
    const driveListMode = await select({
      message: 'Which Drive files? (--drive-samples / picker / all)',
      choices: [
        { name: 'All files', value: 'all' },
        { name: 'First N files in list order (--drive-samples)', value: 'first-n' },
        { name: 'Interactive pick (folders / first N / checkboxes)', value: 'picker' },
      ],
      default: 'picker',
      ...SELECT_OPTS,
    });

    logInfo('Listing Google Drive…');
    const sourceFolderId = extractFolderIdFromUrl(GOOGLE_DRIVE_FOLDER_URL);
    const allFiles = await listFilesRecursively(sourceFolderId, '', null);
    let { driveFiles } = filterDriveSourceExcelFiles(allFiles, {
      allowedOwnerEmails: SOURCE_DRIVE_OWNER_EMAIL ? [SOURCE_DRIVE_OWNER_EMAIL] : [],
      logInfo,
      formatOwnerExcludedCount: (n) =>
        `Filtered out ${n} file(s) not owned by ${SOURCE_DRIVE_OWNER_EMAIL}`,
    });
    if (!driveFiles.length) {
      logWarn('No Excel files in Drive folder.');
      process.exit(1);
    }

    let filesToProcess = driveFiles;
    if (driveListMode === 'first-n') {
      const limitChoice = await select({
        message: 'N for --drive-samples',
        choices: [
          { name: '1', value: '1' },
          { name: '3', value: '3' },
          { name: '5', value: '5' },
          { name: '10', value: '10' },
          { name: '20', value: '20' },
          { name: '50', value: '50' },
        ],
        default: '5',
        ...SELECT_OPTS,
      });
      const limitN = Number(limitChoice);
      filesToProcess = driveFiles.slice(0, Math.min(limitN, driveFiles.length));
      logInfo(`Drive: first ${filesToProcess.length} file(s).`);
    } else if (driveListMode === 'picker') {
      const picked = await pickDriveFilesInteractive(driveFiles);
      filesToProcess = picked.files;
      logInfo(`Drive selection: ${picked.summary}`);
    } else {
      logInfo(`Drive: all ${filesToProcess.length} file(s).`);
    }

    const tempDir = createTempDir();
    try {
      logInfo(`Downloading ${filesToProcess.length} file(s)…`);
      downloadedJobs = await runWithConcurrency(filesToProcess, DEFAULT_CONCURRENCY, async (df) => {
        const localPath = path.join(tempDir, df.path.replace(/\//g, path.sep));
        ensureDir(path.dirname(localPath));
        await downloadFile(df.id, localPath);
        return { localPath, driveFile: df, explicitOutputPath: null };
      });
    } catch (e) {
      cleanupTempDir(tempDir);
      throw e;
    }
    driveFilesToProcess = downloadedJobs;
    return await finishInteractiveAfterSource({
      jsonPath,
      jobs: driveFilesToProcess,
      tempDir,
    });
  }

  return await finishInteractiveAfterSource({ jsonPath, jobs: driveFilesToProcess, tempDir: null });
}

async function finishInteractiveAfterSource({ jsonPath, jobs, tempDir }) {
  const data = loadDraftJson(jsonPath);
  const parityN = normalizeImportedDraftChanges(data.changes).length;
  logInfo(
    `JSON: totalChanged=${data.totalChanged ?? '?'}, linear=${data.linear ?? '?'}, mappings=${parityN} (dendrogram-parity), changedAt=${data.changedAt ?? 'n/a'}`,
  );

  const outLayout = await select({
    message: 'Where to write output .xlsx files? (--output-dir / default mirror / --beside-input)',
    choices: [
      {
        name: `Default mirror under ${path.relative(process.cwd(), OUTPUT_DRAFT_ROOT) || OUTPUT_DRAFT_ROOT}`,
        value: 'mirror',
      },
      { name: 'Beside each source file, same filename (--beside-input; overwrites)', value: 'beside' },
      {
        name: 'Custom base directory (--output-dir; paths mirror Data/ or Drive relative layout)',
        value: 'custom-dir',
      },
    ],
    default: 'mirror',
    ...SELECT_OPTS,
  });
  let besideInput = false;
  let outputDir = null;
  if (outLayout === 'beside') {
    besideInput = true;
  } else if (outLayout === 'custom-dir') {
    const rawDir = String(
      await input({
        message: 'Base output directory (--output-dir)',
        validate: (v) => {
          const p = path.resolve(String(v ?? '').trim());
          return p.length > 0 ? true : 'Required';
        },
      }),
    ).trim();
    outputDir = path.resolve(rawDir);
    ensureDir(outputDir);
  }

  const flagChoices = await checkbox({
    message: 'Options (CLI flags)',
    choices: [
      {
        name: 'Also update FUNCTIONAL LOCATION BEFORE when it equals beforeFuncLoc (--also-before)',
        value: 'alsoBefore',
      },
      {
        name: 'Dry run: no output files written; inputs unchanged on disk (--dry-run)',
        value: 'dryRun',
      },
      {
        name: 'Disable pre-compact row alias matching, legacy replay only (--no-precompact-match)',
        value: 'noPrecompactMatch',
      },
      {
        name: 'Legacy: replay JSON by step order like old CLI (--legacy-step-replay)',
        value: 'legacyStepReplay',
      },
      {
        name: 'Skip Bah Jambi level-5 preprocess before apply (--skip-bah-jambi-preprocess)',
        value: 'skipLevel5',
      },
    ],
  });
  const flagSet = new Set(flagChoices ?? []);
  const alsoBefore = flagSet.has('alsoBefore');
  const dryRun = flagSet.has('dryRun');
  const precompactMatch = !flagSet.has('noPrecompactMatch');
  const legacyStepReplay = flagSet.has('legacyStepReplay');
  const skipBahJambiPreprocess = flagSet.has('skipLevel5');

  const runNow = await confirm({
    message: 'Apply draft JSON with these settings?',
    default: true,
  });
  if (!runNow) {
    logInfo('Cancelled.');
    if (tempDir) cleanupTempDir(tempDir);
    process.exit(0);
  }

  let failures = 0;
  try {
    for (let i = 0; i < jobs.length; i += 1) {
      const job = jobs[i];
      const inputPath = job.localPath ?? job;
      const driveFile = job.driveFile ?? null;
      const namingPath = job.localPath ?? job;
      const outputPath = resolveOutputPath({
        explicitOutputPath: job.explicitOutputPath ?? null,
        inputPath,
        outputBasenameSourcePath: namingPath,
        besideInput,
        outputDir,
        dataRoot: DATA_ROOT,
        draftOutRoot: OUTPUT_DRAFT_ROOT,
        driveFile,
      });
      try {
        await runApplyDraftWithLevel5Preprocess({
          jsonPath,
          inputPath,
          outputPath,
          outputBasenameSourcePath: namingPath,
          alsoBefore,
          precompactMatch,
          dryRun,
          quietJsonLog: true,
          legacyStepReplay,
          skipBahJambiPreprocess,
          bahJambiFixMap: null,
          sharedWorkTempDir: tempDir,
          level5JobKey: `${i}_${job.driveFile?.path ?? namingPath}`,
        });
      } catch (err) {
        failures += 1;
        logError(`${inputPath}: ${err?.message || err}`);
      }
    }
  } finally {
    if (tempDir) cleanupTempDir(tempDir);
  }

  if (failures) logWarn(`Finished with ${failures} error(s).`);
  else logSuccess('Done.');
}

async function main() {
  const argv = process.argv.slice(2);
  if (argv.includes('--help') || argv.includes('-h')) {
    log(`
${bold('apply-draft-changes-json-to-xlsx')}

${bold('Interactive (no arguments, TTY):')}
  Prompts mirror all CLI flags: --json, --local / --input / --output, --drive / Drive link,
  --samples / --drive-samples, --output-dir / --beside-input, --also-before,
  --no-precompact-match, --legacy-step-replay, --dry-run, --skip-bah-jambi-preprocess.
  Default apply = dendrogram JSON import parity (normalize + compact base + applyDraftMappingsToRows).
  v2 always archives removed rows then packs/sorts active rows by FUNCLOC AFTER.

${bold('Flags:')}
  --json <file>       Draft JSON (required for non-interactive). Relative paths resolve from project root.
  --drive-single-file Download one workbook from Drive/Sheets; requires --source and --json (mutually exclusive with --input/--local/--drive)
  --source <url|id>   Google Drive file URL, Sheets URL (…/spreadsheets/d/<ID>/…), or bare file ID (with --drive-single-file)
  --input <xlsx>      Single local workbook
  --output <xlsx>     Single output (with --input or --drive-single-file)
  --local, -l         Batch: all .xlsx under Data/ (use --samples N to limit)
  --drive             Batch: download from same Drive folder as full-validation
  --samples N         Local batch: first N files
  --drive-samples N   Drive batch: first N files (no interactive picker)
  --output-dir <dir>  Batch: write under this folder (mirrors relative paths)
  --beside-input      Batch: write next to each source file (same filename as source; overwrites)
  --also-before       Update FUNCTIONAL LOCATION BEFORE when it matches
  --no-precompact-match Disable compact-gap alias matching (legacy step replay only)
  --legacy-step-replay Old behavior: apply each JSON step in \`step\` order with row matching
  --dry-run           No output file; local disk inputs unchanged
  --skip-bah-jambi-preprocess  Skip level-5 Bah Jambi workbook fix (default: run preprocess to temp, then apply)
  --bah-jambi-fix-map <path>   Override fix-map JSON for that preprocess (default: scripts/drive/mappings/bah-jambi-fix-mapping.json)
`);
    process.exit(0);
  }

  const hasFlags = argv.length > 0;
  const useInteractive = !hasFlags && Boolean(process.stdout.isTTY);

  if (!hasFlags && !process.stdout.isTTY) {
    logError(
      'No TTY: pass flags, e.g. --json Reference/draft.json --input file.xlsx or --drive-single-file --source <url> --json …',
    );
    process.exit(1);
  }

  if (useInteractive) {
    await runInteractiveDraftApply();
    return;
  }

  const args = parseArgs(argv);

  if (args.local && args.drive) {
    logError('Use only one of --local or --drive');
    process.exit(1);
  }

  if (args.driveSingleFile && (args.local || args.drive || args.inputPath)) {
    logError('--drive-single-file cannot be combined with --local, --drive, or --input');
    process.exit(1);
  }

  if (args.driveSingleFile && !args.source) {
    logError('--drive-single-file requires --source <Google Drive / Sheets URL or file id>');
    process.exit(1);
  }

  if (!args.jsonPath || !fs.existsSync(args.jsonPath)) {
    logError('Provide --json <existing file> (relative paths resolve from project root)');
    process.exit(1);
  }

  const draftData = loadDraftJson(args.jsonPath);
  const paritySteps = normalizeImportedDraftChanges(draftData.changes).length;

  if (args.driveSingleFile) {
    const fileId = extractFileIdFromAnyUrl(args.source);
    if (!fileId) {
      logError('Could not parse a Google Drive / Sheets file id from --source');
      process.exit(1);
    }
    const tempDir = createTempDir();
    const localPath = path.join(tempDir, 'drive-single-input.xlsx');
    try {
      logInfo('Downloading workbook from Google Drive (native .xlsx or Sheets export)…');
      await downloadFile(fileId, localPath);
      let besideOut = args.besideInput;
      if (besideOut && !args.outputPath) {
        logWarn(
          '--beside-input would write next to a temp download (deleted after run); using Output/Draft-applied mirror. Use --output <path> for a specific file.',
        );
        besideOut = false;
      }
      const outputPath = resolveOutputPath({
        explicitOutputPath: args.outputPath ?? null,
        inputPath: localPath,
        outputBasenameSourcePath: null,
        besideInput: besideOut,
        outputDir: args.outputDir,
        dataRoot: DATA_ROOT,
        draftOutRoot: OUTPUT_DRAFT_ROOT,
        driveFile: { id: fileId, path: 'drive-single-input.xlsx', name: 'drive-single-input.xlsx' },
      });
      await runApplyDraftWithLevel5Preprocess({
        jsonPath: args.jsonPath,
        inputPath: localPath,
        outputPath,
        outputBasenameSourcePath: localPath,
        alsoBefore: args.alsoBefore,
        precompactMatch: args.precompactMatch,
        dryRun: args.dryRun,
        quietJsonLog: true,
        legacyStepReplay: args.legacyStepReplay,
        skipBahJambiPreprocess: args.skipBahJambiPreprocess,
        bahJambiFixMap: args.bahJambiFixMap,
        sharedWorkTempDir: tempDir,
        level5JobKey: 'drive-single',
      });
    } finally {
      cleanupTempDir(tempDir);
    }
    return;
  }

  if (args.inputPath) {
    if (args.local || args.drive) {
      logWarn('--input is set; ignoring --local / --drive.');
    }
    if (!fs.existsSync(args.inputPath)) {
      logError(`Input not found: ${args.inputPath}`);
      process.exit(1);
    }
    const outputPath = resolveOutputPath({
      explicitOutputPath: args.outputPath,
      inputPath: args.inputPath,
      outputBasenameSourcePath: args.inputPath,
      besideInput: args.besideInput,
      outputDir: args.outputDir,
      dataRoot: DATA_ROOT,
      draftOutRoot: OUTPUT_DRAFT_ROOT,
      driveFile: null,
    });
    await runApplyDraftWithLevel5Preprocess({
      jsonPath: args.jsonPath,
      inputPath: args.inputPath,
      outputPath,
      outputBasenameSourcePath: args.inputPath,
      alsoBefore: args.alsoBefore,
      precompactMatch: args.precompactMatch,
      dryRun: args.dryRun,
      quietJsonLog: true,
      legacyStepReplay: args.legacyStepReplay,
      skipBahJambiPreprocess: args.skipBahJambiPreprocess,
      bahJambiFixMap: args.bahJambiFixMap,
      sharedWorkTempDir: null,
      level5JobKey: args.inputPath,
    });
    return;
  }

  if (args.local) {
    let excelFiles = collectExcelFiles(DATA_ROOT).sort();
    if (!excelFiles.length) {
      logWarn('No Excel files under Data/.');
      process.exit(1);
    }
    if (args.samples != null) excelFiles = excelFiles.slice(0, args.samples);
    logInfo(`Local batch: ${excelFiles.length} file(s), ${paritySteps} dendrogram-parity mapping(s).`);
    for (const inputPath of excelFiles) {
      const outputPath = resolveOutputPath({
        explicitOutputPath: null,
        inputPath,
        outputBasenameSourcePath: inputPath,
        besideInput: args.besideInput,
        outputDir: args.outputDir,
        dataRoot: DATA_ROOT,
        draftOutRoot: OUTPUT_DRAFT_ROOT,
        driveFile: null,
      });
      try {
        await runApplyDraftWithLevel5Preprocess({
          jsonPath: args.jsonPath,
          inputPath,
          outputPath,
          outputBasenameSourcePath: inputPath,
          alsoBefore: args.alsoBefore,
          precompactMatch: args.precompactMatch,
          dryRun: args.dryRun,
          quietJsonLog: true,
          legacyStepReplay: args.legacyStepReplay,
          skipBahJambiPreprocess: args.skipBahJambiPreprocess,
          bahJambiFixMap: args.bahJambiFixMap,
          sharedWorkTempDir: null,
          level5JobKey: inputPath,
        });
      } catch (err) {
        logError(`${inputPath}: ${err?.message || err}`);
      }
    }
    return;
  }

  if (args.drive) {
    const sourceFolderId = extractFolderIdFromUrl(GOOGLE_DRIVE_FOLDER_URL);
    const allFiles = await listFilesRecursively(sourceFolderId, '', null);
    let { driveFiles } = filterDriveSourceExcelFiles(allFiles, {
      allowedOwnerEmails: SOURCE_DRIVE_OWNER_EMAIL ? [SOURCE_DRIVE_OWNER_EMAIL] : [],
      logInfo,
      formatOwnerExcludedCount: (n) =>
        `Filtered out ${n} file(s) not owned by ${SOURCE_DRIVE_OWNER_EMAIL}`,
    });
    if (!driveFiles.length) {
      logWarn('No Excel files in Drive folder.');
      process.exit(1);
    }
    if (args.driveSamples != null) {
      driveFiles = driveFiles.slice(0, args.driveSamples);
    }
    logInfo(`Drive batch: ${driveFiles.length} file(s), ${paritySteps} dendrogram-parity mapping(s).`);
    const tempDir = createTempDir();
    try {
      const downloaded = await runWithConcurrency(driveFiles, DEFAULT_CONCURRENCY, async (df) => {
        const localPath = path.join(tempDir, df.path.replace(/\//g, path.sep));
        ensureDir(path.dirname(localPath));
        await downloadFile(df.id, localPath);
        return { localPath, driveFile: df };
      });
      for (let i = 0; i < downloaded.length; i += 1) {
        const { localPath, driveFile } = downloaded[i];
        const outputPath = resolveOutputPath({
          explicitOutputPath: null,
          inputPath: localPath,
          outputBasenameSourcePath: localPath,
          besideInput: args.besideInput,
          outputDir: args.outputDir,
          dataRoot: DATA_ROOT,
          draftOutRoot: OUTPUT_DRAFT_ROOT,
          driveFile,
        });
        try {
          await runApplyDraftWithLevel5Preprocess({
            jsonPath: args.jsonPath,
            inputPath: localPath,
            outputPath,
            outputBasenameSourcePath: localPath,
            alsoBefore: args.alsoBefore,
            precompactMatch: args.precompactMatch,
            dryRun: args.dryRun,
            quietJsonLog: true,
            legacyStepReplay: args.legacyStepReplay,
            skipBahJambiPreprocess: args.skipBahJambiPreprocess,
            bahJambiFixMap: args.bahJambiFixMap,
            sharedWorkTempDir: tempDir,
            level5JobKey: `${i}_${driveFile.path}`,
          });
        } catch (err) {
          logError(`${driveFile.path}: ${err?.message || err}`);
        }
      }
    } finally {
      cleanupTempDir(tempDir);
    }
    return;
  }

  logError('Use --input <file>, or --local, or --drive (with --json).');
  process.exit(1);
}

export {
  orderChanges,
  applyChanges,
  loadOrderedDraft,
  normFuncloc,
  normDesc,
  normDescLoose,
  deriveCostCenterFromFunclocAfter,
  mergeAfterFunclocKeepExcelFirstTwoSegments,
  normalizeImportedDraftChanges,
  compactGapFuncLocDraft,
  applyDraftMappingsToRowsDendrogramParity,
};

if (isExecutedAsCli()) {
  main().catch((err) => {
    logError(err?.message || String(err));
    process.exit(1);
  });
}
