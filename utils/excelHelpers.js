import path from 'node:path';
import * as xlsxNs from 'xlsx';
import { ensureDir } from './fileSystem.js';

const xlsx = xlsxNs.default ?? xlsxNs;

/**
 * Normalize header for matching: trim, collapse whitespace, optional punctuation.
 * Makes detection robust to leading/trailing spaces, extra spaces, and "DESC." vs "DESC".
 */
export const normalizeHeader = (value) =>
  String(value ?? '')
    .trim()
    .replace(/[`'"\u2018\u2019\u201c\u201d]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/\./g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();

export const toCleanString = (value) => {
  if (value === null || value === undefined) return '';

  // ExcelJS/xlsx can return plain values or objects (richText, formula+result, etc.).
  // Extract a meaningful string before cleaning.
  const extract = (v) => {
    if (v === null || v === undefined) return '';
    const t = typeof v;
    if (t === 'string' || t === 'number' || t === 'boolean') return String(v);
    if (t === 'object') {
      // Rich text: { richText: [{ text: '...' }, ...] }
      if (Array.isArray(v.richText)) {
        return v.richText.map((rt) => (rt && typeof rt.text === 'string' ? rt.text : '')).join('');
      }
      // Simple text wrapper: { text: '...' }
      if (typeof v.text === 'string') {
        return v.text;
      }
      // Formula with cached result: { formula: '...', result: '...' }
      if ('result' in v && v.result !== null && v.result !== undefined) {
        return String(v.result);
      }
      // Fallback: JSON (for debug) rather than bare [object Object]
      try {
        return JSON.stringify(v);
      } catch {
        return String(v);
      }
    }
    return String(v);
  };

  const rawStr = extract(value);
  // Clean the string by removing unwanted characters like backticks, apostrophes, quotes, and extra spaces
  const cleaned = rawStr.replace(/[`'"\u2018\u2019\u201c\u201d]/g, '').replace(/\s+/g, '');
  return cleaned.trim();
};

/**
 * Display string helper: keep internal spaces but normalize quotes.
 */
export const toDisplayString = (value) => {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/[`'"\u2018\u2019\u201c\u201d]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
};

/**
 * ExcelJS cell value → readable text (richText, hyperlink `{ text }`, nested formula `result`).
 * Normalizes whitespace. Use instead of `String(cell.value)` so you never get "[object Object]".
 */
export const excelJsValueToPlainText = (value) => {
  const extract = (v) => {
    if (v === null || v === undefined) return '';
    const t = typeof v;
    if (t === 'string' || t === 'number' || t === 'boolean') return String(v);
    if (t === 'object') {
      if (v instanceof Date) return String(v);
      if (Array.isArray(v.richText)) {
        return v.richText.map((rt) => (rt && typeof rt.text === 'string' ? rt.text : '')).join('');
      }
      if (typeof v.text === 'string') return v.text;
      if ('result' in v && v.result !== null && v.result !== undefined) {
        return extract(v.result);
      }
      try {
        return JSON.stringify(v);
      } catch {
        return '';
      }
    }
    return String(v);
  };
  const raw = extract(value);
  return raw
    .replace(/[`'"\u2018\u2019\u201c\u201d]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
};

/**
 * Levenshtein distance between two strings
 */
export const levenshtein = (a, b) => {
  const m = a.length;
  const n = b.length;
  const d = Array(m + 1)
    .fill(null)
    .map(() => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i += 1) d[i][0] = i;
  for (let j = 0; j <= n; j += 1) d[0][j] = j;
  for (let i = 1; i <= m; i += 1) {
    for (let j = 1; j <= n; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
    }
  }
  return d[m][n];
};

/**
 * Similarity score 0–1 (1 = identical). Uses normalized strings.
 */
export const similarity = (a, b) => {
  const an = normalizeHeader(a);
  const bn = normalizeHeader(b);
  if (!an || !bn) return 0;
  const d = levenshtein(an, bn);
  const maxLen = Math.max(an.length, bn.length);
  return maxLen === 0 ? 1 : 1 - d / maxLen;
};

/**
 * Find best match for wanted header among sheet headers.
 * Returns { best, raw, score, col } or null.
 */
export const findBestMatch = (wanted, headers) => {
  let best = null;
  let bestRaw = '';
  let bestScore = -1;
  let bestCol = -1;
  const wantedNorm = normalizeHeader(wanted);
  for (const { raw, norm, col } of headers) {
    const score = similarity(wantedNorm, norm);
    if (score > bestScore) {
      bestScore = score;
      best = norm;
      bestRaw = raw;
      bestCol = col;
    }
  }
  if (best == null) return null;
  return { best, raw: bestRaw, score: Math.round(bestScore * 1000) / 1000, col: bestCol };
};

/**
 * Collect all headers in a row as { raw, norm, col } for xlsx worksheets.
 */
export const getAllHeadersInRow = (worksheet, rowIndex) => {
  const ref = worksheet['!ref'] || 'A1:A1';
  const range = xlsx.utils.decode_range(ref);
  const out = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = xlsx.utils.encode_cell({ r: rowIndex, c });
    const cell = worksheet[addr];
    const raw = String(cell?.v ?? '').trim();
    const norm = normalizeHeader(raw);
    if (!norm) continue;
    out.push({ raw, norm, col: c });
  }
  return out;
};

/** Min similarity to consider "COST CENTER BEFORE" (typos, COST CENTRE BEFORE, etc.) */
const DEFAULT_COST_CENTER_BEFORE_SIMILARITY = 0.7;

/**
 * Find a likely column for a target header using fuzzy matching (xlsx worksheet).
 * Excludes given columns (e.g. COST CENTER AFTER) so we don't pick the wrong one.
 */
export const findLikelyColumn = (
  worksheet,
  headerRowIndex,
  targetHeader,
  { excludeCols = [], minSimilarity = DEFAULT_COST_CENTER_BEFORE_SIMILARITY } = {},
) => {
  const headers = getAllHeadersInRow(worksheet, headerRowIndex);
  const exclude = new Set(excludeCols);
  const filtered = headers.filter((h) => !exclude.has(h.col));
  const match = findBestMatch(targetHeader, filtered);
  if (!match || match.score < minSimilarity) return undefined;
  return { col: match.col, raw: match.raw, score: match.score };
};

// Scan only the first few rows for headers; most of your files have headers in row 1–10.
// Find header row and columns by scanning only the first few rows of the worksheet.
// If similarityThreshold is provided (0-1), uses fuzzy matching for headers.
// IMPORTANT:
// - Supports "double header" sheets (e.g. summary header row above detailed header row)
//   by scanning candidate rows and picking the row with the MOST matched wanted headers.
//   This mirrors findHeaderRowAndColumnsExcelJs in apply-template-to-others.js.
export const findHeaderRowAndColumns = (
  worksheet,
  wantedHeaders,
  { scanRows = 10, similarityThreshold = null } = {},
) => {
  const wanted = wantedHeaders.map(normalizeHeader);
  const ref = worksheet['!ref'] || 'A1:A1';
  const range = xlsx.utils.decode_range(ref);

  const maxRow = Math.min(range.e.r, range.s.r + scanRows - 1);

  let best = null;
  let bestMatchCount = 0;

  for (let r = range.s.r; r <= maxRow; r += 1) {
    const colByHeader = new Map();
    const headerMatches = new Map(); // Store match info: { originalHeader, matchedHeader, raw, score }

    if (similarityThreshold !== null && similarityThreshold > 0) {
      // Fuzzy matching mode
      const rowHeaders = getAllHeadersInRow(worksheet, r);

      for (const wantedHeader of wanted) {
        // First try exact match
        const exactMatch = rowHeaders.find((h) => h.norm === wantedHeader);
        if (exactMatch) {
          colByHeader.set(wantedHeader, exactMatch.col);
          headerMatches.set(wantedHeader, {
            originalHeader: wantedHeader,
            matchedHeader: exactMatch.norm,
            raw: exactMatch.raw,
            score: 1.0,
          });
          continue;
        }

        // Try substring match (fallback for backwards compatibility)
        const substringMatch = rowHeaders.find(
          (h) => h.norm.includes(wantedHeader) || wantedHeader.includes(h.norm),
        );
        if (substringMatch) {
          colByHeader.set(wantedHeader, substringMatch.col);
          headerMatches.set(wantedHeader, {
            originalHeader: wantedHeader,
            matchedHeader: substringMatch.norm,
            raw: substringMatch.raw,
            score: 1.0,
          });
          continue;
        }

        // Try fuzzy matching
        const bestMatch = findBestMatch(wantedHeader, rowHeaders);
        if (bestMatch && bestMatch.score >= similarityThreshold) {
          colByHeader.set(wantedHeader, bestMatch.col);
          headerMatches.set(wantedHeader, {
            originalHeader: wantedHeader,
            matchedHeader: bestMatch.best,
            raw: bestMatch.raw,
            score: bestMatch.score,
          });
        }
      }
    } else {
      // Exact matching mode (original behavior, but pick best header row)
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const addr = xlsx.utils.encode_cell({ r, c });
        const cell = worksheet[addr];
        if (!cell) continue;
        const norm = normalizeHeader(cell.v);
        if (!norm) continue;

        if (wanted.includes(norm)) {
          colByHeader.set(norm, c);
          headerMatches.set(norm, {
            originalHeader: norm,
            matchedHeader: norm,
            raw: String(cell.v),
            score: 1.0,
          });
        }

        wanted.forEach((header) => {
          if (!colByHeader.has(header) && norm.includes(header)) {
            colByHeader.set(header, c);
            headerMatches.set(header, {
              originalHeader: header,
              matchedHeader: norm,
              raw: String(cell.v),
              score: 1.0,
            });
          }
        });
      }
    }

    const matchCount = wanted.filter((header) => colByHeader.has(header)).length;
    if (matchCount >= bestMatchCount && matchCount > 0) {
      bestMatchCount = matchCount;
      best = {
        headerRowIndex: r,
        colByHeader,
        headerMatches: similarityThreshold !== null ? headerMatches : undefined,
      };
    }
  }

  return best;
};

/**
 * ExcelJS helper: get cell string from a row/column (using cached formula result when present).
 */
export const getCellStringFromRow = (row, col) => {
  if (!col) return '';
  const cell = row.getCell(col);
  const v = cell?.value;
  if (
    v != null &&
    typeof v === 'object' &&
    ('formula' in v || 'sharedFormula' in v) &&
    v.result != null
  ) {
    return toCleanString(v.result);
  }
  return toCleanString(v);
};

/**
 * ExcelJS helper: get display string from a row/column (preserving spaces).
 */
export const getCellDisplayStringFromRow = (row, col) => {
  if (!col) return '';
  const cell = row.getCell(col);
  const v = cell?.value;
  if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
    return excelJsValueToPlainText(v.result ?? '');
  }
  return excelJsValueToPlainText(v);
};

/**
 * ExcelJS helper: find headers in the first few rows using fuzzy matching.
 * Mirrors the logic used in final-process.js but shared here.
 */
export const findHeadersExcelJs = (
  worksheet,
  wantedHeaders,
  { maxScanRows = 10, minSimilarity = 0.6 } = {},
) => {
  const wanted = wantedHeaders.map(normalizeHeader);
  const maxRow = Math.min(worksheet.rowCount, maxScanRows);

  for (let r = 1; r <= maxRow; r += 1) {
    const row = worksheet.getRow(r);
    const colMap = new Map();
    const scoreByWanted = new Map();
    const bestWantedByCol = new Map();
    const headers = [];

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const norm = normalizeHeader(String(cell.value ?? ''));
      if (!norm) return;
      headers.push({ norm, colNumber });

      // Exact match
      if (wanted.includes(norm)) {
        colMap.set(norm, colNumber);
        scoreByWanted.set(norm, 1);
        bestWantedByCol.set(colNumber, { wanted: norm, score: 1 });
      }

      // Substring match (for headers that contain our wanted text)
      for (const w of wanted) {
        if (!colMap.has(w) && (norm.includes(w) || w.includes(norm))) {
          colMap.set(w, colNumber);
          const score = 0.95;
          scoreByWanted.set(w, score);
          const existing = bestWantedByCol.get(colNumber);
          if (!existing || score > existing.score) {
            bestWantedByCol.set(colNumber, { wanted: w, score });
          }
        }
      }
    });

    // Fuzzy match pass for any still-missing wanted headers.
    for (const w of wanted) {
      if (colMap.has(w)) continue;
      let best = null;
      let bestScore = -1;
      for (const h of headers) {
        const sc = similarity(w, h.norm);
        if (sc > bestScore) {
          bestScore = sc;
          best = h;
        }
      }
      if (best && bestScore >= minSimilarity) {
        const existing = bestWantedByCol.get(best.colNumber);
        if (!existing || bestScore > existing.score) {
          // If another wanted already claimed this column with lower score, remove it.
          if (existing) {
            colMap.delete(existing.wanted);
            scoreByWanted.delete(existing.wanted);
          }
          colMap.set(w, best.colNumber);
          scoreByWanted.set(w, bestScore);
          bestWantedByCol.set(best.colNumber, { wanted: w, score: bestScore });
        }
      }
    }

    if (wanted.some((w) => colMap.has(w))) {
      return { headerRowNumber: r, colMap };
    }
  }

  return null;
};

export const readExcelSheet = (excelPath) => {
  // Read with styles so we can preserve formatting when writing back.
  const workbook = xlsx.readFile(excelPath, { cellDates: true, cellStyles: true });
  const sheetName = workbook.SheetNames.at(0);
  if (!sheetName) throw new Error(`No sheets found in ${excelPath}`);

  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`Worksheet "${sheetName}" not found in ${excelPath}`);
  }
  return { workbook, sheetName, worksheet };
};

export const appendWorkCenterColumnToSheet = (
  worksheet,
  headerRowIndex,
  costCenterCol,
  headerValue,
) => {
  const ref = worksheet['!ref'] || 'A1:A1';
  const range = xlsx.utils.decode_range(ref);
  const wantedHeader = normalizeHeader(headerValue);

  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = xlsx.utils.encode_cell({ r: headerRowIndex, c });
    const existingHeader = normalizeHeader(worksheet[addr]?.v ?? '');
    if (existingHeader === wantedHeader) {
      return c;
    }
  }

  // Trim trailing columns that have no values at all (only formatting), so we
  // don't append our new column after a huge block of empty formatted cells.
  let lastDataCol = range.e.c;
  outer: for (let c = range.e.c; c >= range.s.c; c -= 1) {
    for (let r = range.s.r; r <= range.e.r; r += 1) {
      const addr = xlsx.utils.encode_cell({ r, c });
      const cell = worksheet[addr];
      if (cell && cell.v !== undefined && cell.v !== null && `${cell.v}` !== '') {
        lastDataCol = c;
        break outer;
      }
    }
  }

  range.e.c = Math.max(lastDataCol, costCenterCol);
  const appendColIndex = Math.max(range.e.c + 1, costCenterCol + 1);
  // Expand range to include the new column so it shows up in Excel viewers.
  range.e.c = Math.max(range.e.c, appendColIndex);
  worksheet['!ref'] = xlsx.utils.encode_range(range);

  const headerAddr = xlsx.utils.encode_cell({ r: headerRowIndex, c: appendColIndex });
  const cell = worksheet[headerAddr] || {};
  cell.v = headerValue;
  cell.t = 's';

  // Copy style from source header if available.
  const srcHeaderAddr = xlsx.utils.encode_cell({ r: headerRowIndex, c: costCenterCol });
  let srcCell = worksheet[srcHeaderAddr];

  // Some sheets use row- or table-level formatting so the cost center header
  // cell may not carry the style. In that case, fall back to the first styled
  // cell we find on the same header row.
  if (!srcCell || !srcCell.s) {
    const ref = worksheet['!ref'] || 'A1:A1';
    const range = xlsx.utils.decode_range(ref);
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = xlsx.utils.encode_cell({ r: headerRowIndex, c });
      const candidate = worksheet[addr];
      if (candidate && candidate.s) {
        srcCell = candidate;
        break;
      }
    }
  }

  if (srcCell && srcCell.s) {
    cell.s = { ...srcCell.s };
  }

  worksheet[headerAddr] = cell;

  return appendColIndex;
};

export const writeWorkCentersToSheet = (
  worksheet,
  headerRowIndex,
  costCenterCol,
  appendColIndex,
  getCwcForCostCenter,
) => {
  const ref = worksheet['!ref'] || 'A1:A1';
  const range = xlsx.utils.decode_range(ref);
  const results = [];
  let maxLen =
    (worksheet[xlsx.utils.encode_cell({ r: headerRowIndex, c: appendColIndex })]?.v || '').length ||
    0;

  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    // Read cost center from its column.
    const costAddr = xlsx.utils.encode_cell({ r, c: costCenterCol });
    const costCell = worksheet[costAddr];
    const costCenter = toCleanString(costCell?.v);

    if (!costCenter) {
      // Check if the entire row is empty; if so, skip it silently.
      let anyValue = false;
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const addr = xlsx.utils.encode_cell({ r, c });
        const cell = worksheet[addr];
        if (toCleanString(cell?.v)) {
          anyValue = true;
          break;
        }
      }
      if (!anyValue) continue;
    }

    if (!costCenter) continue;

    const cwc = getCwcForCostCenter(costCenter);

    const addr = xlsx.utils.encode_cell({ r, c: appendColIndex });
    const cell = worksheet[addr] || {};
    cell.v = cwc ?? '';
    cell.t = 's';

    // Copy style from the cost center column for this row if possible.
    const srcAddr = xlsx.utils.encode_cell({ r, c: costCenterCol });
    let srcCell = worksheet[srcAddr];

    // If that cell has no explicit style (e.g. banded rows defined elsewhere),
    // fall back to the first styled cell in this row so background banding
    // still looks correct.
    if (!srcCell || !srcCell.s) {
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const addr = xlsx.utils.encode_cell({ r, c });
        const candidate = worksheet[addr];
        if (candidate && candidate.s) {
          srcCell = candidate;
          break;
        }
      }
    }

    if (srcCell && srcCell.s) {
      cell.s = { ...srcCell.s };
    }

    worksheet[addr] = cell;

    const len = String(cell.v || '').length;
    if (len > maxLen) maxLen = len;

    results.push({ row: r + 1, costCenter, cwc });
  }

  // Auto-fit width for the new column based on content length so text is not hidden.
  if (typeof appendColIndex === 'number' && maxLen > 0) {
    const desiredWidth = maxLen + 2; // small padding

    const cols = worksheet['!cols'] || [];
    const existing = cols[appendColIndex];
    const currentWidth = existing?.wch || existing?.wpx || 0;

    cols[appendColIndex] = {
      ...(existing || {}),
      wch: Math.max(currentWidth || 0, desiredWidth),
    };
    worksheet['!cols'] = cols;
  }

  return results;
};

export const writeWorkbookToFile = (workbook, worksheet, outputPath) => {
  // Ensure directory exists and write the modified workbook.
  const ref = worksheet['!ref'] || 'A1:A1';
  worksheet['!ref'] = ref;
  ensureDir(path.dirname(outputPath));
  // Write with cellStyles option to preserve styles (xlsx free version has limited support)
  xlsx.writeFile(workbook, outputPath, { cellStyles: true });
};

export const findFirstNonEmptyInColumn = (worksheet, colIndex, startRowIndex, range) => {
  for (let r = startRowIndex; r <= range.e.r; r += 1) {
    const addr = xlsx.utils.encode_cell({ r, c: colIndex });
    const cell = worksheet[addr];
    const value = toCleanString(cell?.v);
    if (value) return value;
  }
  return null;
};

// Note: applyRowFill is no longer used directly
// Row colors are now applied via writeWorkbookWithRowColors in excelWriteWithStyles.js
// This function is kept for reference but row coloring happens during write
export const applyRowFill = (worksheet, rowIndex, range, rgb) => {
  // This is a no-op now - colors are applied during write
  // Kept for compatibility but actual coloring happens in writeWorkbookWithRowColors
};

export const validateAndColorRows = ({
  worksheet,
  headerRowIndex,
  maintenancePlantCol,
  funclocCol,
  costCenterCol,
}) => {
  const ref = worksheet['!ref'] || 'A1:A1';
  const range = xlsx.utils.decode_range(ref);

  // Find maintenance plant from first non-empty value in that column
  const maintenancePlantRaw = findFirstNonEmptyInColumn(
    worksheet,
    maintenancePlantCol,
    headerRowIndex + 1,
    range,
  );
  const maintenancePlant = maintenancePlantRaw
    ? toCleanString(maintenancePlantRaw).toUpperCase()
    : null;

  if (!maintenancePlant) {
    return {
      totalRows: 0,
      green: 0,
      yellow: 0,
      red: 0,
      maintenancePlant: null,
      error: 'MAINTENANCE PLANT not found',
    };
  }

  let totalRows = 0;
  let green = 0;
  let yellow = 0;
  let red = 0;

  // Color definitions (ARGB format for exceljs - 8 hex digits)
  const COLOR_GREEN = 'FFC6EFCE'; // Light green
  const COLOR_YELLOW = 'FFFFEB9C'; // Light yellow
  const COLOR_RED = 'FFFFC7CE'; // Light red

  // Map to store row colors (rowIndex -> color ARGB)
  const rowColors = new Map();

  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    // Check if row is empty (same logic as existing writer)
    let anyValue = false;
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = xlsx.utils.encode_cell({ r, c });
      const cell = worksheet[addr];
      if (toCleanString(cell?.v)) {
        anyValue = true;
        break;
      }
    }
    if (!anyValue) continue;

    totalRows += 1;

    // Get functional location
    const funclocAddr = xlsx.utils.encode_cell({ r, c: funclocCol });
    const funclocCell = worksheet[funclocAddr];
    const funclocRaw = toCleanString(funclocCell?.v);

    // Get cost center
    const costAddr = xlsx.utils.encode_cell({ r, c: costCenterCol });
    const costCell = worksheet[costAddr];
    const costCenterRaw = toCleanString(costCell?.v);

    // Check 1: Functional Location plant code match
    let funcOk = false;
    if (funclocRaw) {
      const segments = funclocRaw.toUpperCase().split('-');
      if (segments.length > 1) {
        const plantFromFuncLoc = segments[1];
        funcOk = plantFromFuncLoc === maintenancePlant;
      }
    }

    // Check 2: Cost Center last 2 digits match functional location segment[3] last 2 digits
    let costOk = false;
    if (funclocRaw && costCenterRaw) {
      const segments = funclocRaw.toUpperCase().split('-');
      if (segments.length > 3 && segments[3]) {
        const seg3 = segments[3];
        // Extract last 2 digits from segment[3]
        const seg3Match = seg3.match(/\d{2}$/);
        const seg3Last2 = seg3Match ? seg3Match[0] : seg3.length >= 2 ? seg3.slice(-2) : null;

        // Extract last 2 digits from cost center
        const ccMatch = costCenterRaw.match(/\d{2}$/);
        const ccLast2 = ccMatch
          ? ccMatch[0]
          : costCenterRaw.length >= 2
            ? costCenterRaw.slice(-2)
            : null;

        if (seg3Last2 && ccLast2) {
          costOk = seg3Last2 === ccLast2;
        }
      }
    }

    // Determine color based on checks
    let color;
    if (funcOk && costOk) {
      color = COLOR_GREEN;
      green += 1;
    } else if (funcOk || costOk) {
      color = COLOR_YELLOW;
      yellow += 1;
    } else {
      color = COLOR_RED;
      red += 1;
    }

    // Store row color for later application during write
    rowColors.set(r, color);
  }

  return {
    totalRows,
    green,
    yellow,
    red,
    maintenancePlant,
    rowColors, // Return the map of row colors
  };
};

/**
 * ExcelJS style helpers
 */
export const copyStyle = (srcCell, dstCell) => {
  if (!srcCell || !dstCell) return;
  if (srcCell.font) dstCell.font = { ...srcCell.font };
  if (srcCell.alignment) dstCell.alignment = { ...srcCell.alignment };
  if (srcCell.border) dstCell.border = { ...srcCell.border };
  if (srcCell.fill && srcCell.fill.type) dstCell.fill = { ...srcCell.fill };
  if (srcCell.numFmt) dstCell.numFmt = srcCell.numFmt;
};

export const extractStyle = (cell) => {
  if (!cell) return null;
  return {
    font: cell.font ? { ...cell.font } : null,
    alignment: cell.alignment ? { ...cell.alignment } : null,
    border: cell.border ? { ...cell.border } : null,
    fill: cell.fill && cell.fill.type ? { ...cell.fill } : null,
    numFmt: cell.numFmt ?? null,
  };
};

export const applyExtractedStyle = (style, dstCell) => {
  if (!style || !dstCell) return;
  if (style.font) dstCell.font = { ...style.font };
  if (style.alignment) dstCell.alignment = { ...style.alignment };
  if (style.border) dstCell.border = { ...style.border };
  if (style.fill && style.fill.type) dstCell.fill = { ...style.fill };
  if (style.numFmt) dstCell.numFmt = style.numFmt;
};

/**
 * Find a styled reference cell in an ExcelJS row to copy formatting from.
 * Tries the specified column first, then falls back to the first cell with a font.
 */
export const getRefCell = (row, preferCol, lastOrigCol) => {
  if (preferCol) {
    const cell = row.getCell(preferCol);
    if (cell && cell.font && Object.keys(cell.font).length > 0) return cell;
  }
  const maxCol = lastOrigCol || row.cellCount;
  for (let c = 1; c <= maxCol; c += 1) {
    const cell = row.getCell(c);
    if (cell && cell.font && Object.keys(cell.font).length > 0) return cell;
  }
  return null;
};

/**
 * Replace all formulas in an ExcelJS worksheet with their cached values.
 * This ensures we never depend on formulas when reading or writing,
 * and avoids Excel formula errors when deleting rows.
 */
export const stripFormulasFromWorksheet = (worksheet) => {
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    let mutated = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
        const result = v.result;
        if (result !== undefined && result !== null) {
          cell.value = result;
        } else {
          cell.value = '';
        }
        mutated = true;
      }
    });
    if (mutated) {
      row.commit();
    }
  });
};

/**
 * Helper to find cost center column in ExcelJS-based workflows:
 * prefer COST CENTER AFTER, fall back to COST CENTER.
 */
export const findCostCenterCol = (colMap) => {
  const afterKey = normalizeHeader('COST CENTER AFTER');
  if (colMap.has(afterKey)) return colMap.get(afterKey);
  const ccKey = normalizeHeader('COST CENTER');
  if (colMap.has(ccKey)) return colMap.get(ccKey);
  return undefined;
};
