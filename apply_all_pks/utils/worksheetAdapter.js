import ExcelJS from 'exceljs';

import { toCleanString } from './excelHelpers.js';
import { fillIsReddishPreserveFromXlsx } from './excelWriteWithStyles.js';

export const colToLetters = (colIndex0) => {
  let n = colIndex0 + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

export const encodeCell = ({ r, c }) => `${colToLetters(c)}${r + 1}`;

export const lettersToCol0 = (letters) => {
  let n = 0;
  String(letters)
    .toUpperCase()
    .split('')
    .forEach((ch) => {
      n = n * 26 + (ch.charCodeAt(0) - 64);
    });
  return n - 1;
};

export const decodeRange = (ref) => {
  const raw = String(ref || 'A1:A1');
  const [a, b] = raw.split(':');
  const parseAddr = (addr) => {
    const m = String(addr).match(/^([A-Z]+)(\d+)$/i);
    if (!m) return { r: 0, c: 0 };
    return { c: lettersToCol0(m[1]), r: Math.max(0, Number(m[2]) - 1) };
  };
  return { s: parseAddr(a), e: parseAddr(b || a) };
};

export const excelJsColorToRgb = (color) => {
  const argb = color?.argb;
  if (!argb || typeof argb !== 'string') return undefined;
  return argb.length === 8 ? argb.slice(2).toUpperCase() : argb.toUpperCase();
};

export const toXlsxLikeStyle = (style = {}) => {
  const out = {};
  if (style.font) {
    out.font = { ...style.font };
    if (style.font.color) {
      const rgb = excelJsColorToRgb(style.font.color);
      out.font.color = rgb ? { rgb } : style.font.color;
    }
  }
  if (style.alignment) out.alignment = { ...style.alignment };
  if (style.border) out.border = { ...style.border };
  if (style.fill?.fgColor) {
    const rgb = excelJsColorToRgb(style.fill.fgColor);
    if (rgb) {
      out.fill = {
        fgColor: { rgb },
        patternType: style.fill.pattern || 'solid',
      };
    }
  }
  return Object.keys(out).length ? out : undefined;
};

export const excelJsCellToRaw = (cell) => {
  const v = cell?.value;
  if (v == null) return '';
  if (typeof v !== 'object') return String(v);
  if ('formula' in v || 'sharedFormula' in v) {
    if (v.result == null) return '';
    return String(v.result);
  }
  if (Array.isArray(v.richText)) return v.richText.map((rt) => rt?.text ?? '').join('');
  if (v.text != null) return String(v.text);
  if (v.error != null) return String(v.error);
  if (v instanceof Date) return v.toISOString();
  return '';
};

export const readExcelSheetCompat = async (excelPath) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(excelPath);
  const sheet = wb.worksheets[0];
  if (!sheet) throw new Error(`No sheets found in ${excelPath}`);

  const worksheet = {};
  const maxRow = Math.max(sheet.actualRowCount || 0, sheet.rowCount || 0, sheet.lastRow?.number || 0, 1);
  const maxCol = Math.max(sheet.columnCount || 0, 1);

  for (let r = 1; r <= maxRow; r += 1) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= maxCol; c += 1) {
      const cell = row.getCell(c);
      const value = excelJsCellToRaw(cell);
      const style = toXlsxLikeStyle(cell.style);
      if (value === '' && !style) continue;
      const addr = encodeCell({ r: r - 1, c: c - 1 });
      worksheet[addr] = {
        ...(value !== '' ? { v: value, t: 's' } : {}),
        ...(style ? { s: style } : {}),
      };
    }
  }

  worksheet['!ref'] =
    `A1:${encodeCell({ r: Math.max(0, maxRow - 1), c: Math.max(0, maxCol - 1) })}`;

  const cols = [];
  for (let c = 1; c <= maxCol; c += 1) {
    const w = sheet.getColumn(c).width;
    if (w != null) cols[c - 1] = { wch: w };
  }
  if (cols.length) worksheet['!cols'] = cols;

  const rows = [];
  for (let r = 1; r <= maxRow; r += 1) {
    const h = sheet.getRow(r).height;
    if (h != null) rows[r - 1] = { hpt: h };
  }
  if (rows.length) worksheet['!rows'] = rows;

  return { workbook: wb, sheet, worksheet };
};

export const getCellValue = (worksheet, rowIndex, colIndex) => {
  if (colIndex === undefined) return undefined;
  return worksheet[encodeCell({ r: rowIndex, c: colIndex })]?.v;
};

export const getCleanCellValue = (worksheet, rowIndex, colIndex) =>
  toCleanString(getCellValue(worksheet, rowIndex, colIndex));

/**
 * Drop light-red / pink “inactive equipment” fill from a SheetJS-style cell style object.
 * Other style keys (font, border, …) are kept.
 */
export const stripInactiveRedFillFromStyle = (s) => {
  if (!s?.fill || !fillIsReddishPreserveFromXlsx(s.fill)) return s;
  const { fill: _drop, ...rest } = s;
  return Object.keys(rest).length ? rest : undefined;
};

/**
 * After row moves / clears: empty text must not keep inactive red fill (misleading).
 */
export const sanitizeInactiveRedOnEmptyTextCell = (cell) => {
  if (!cell) return null;
  const out = { ...cell };
  if (cell.s) out.s = { ...cell.s };
  const v = out.v;
  const empty =
    v === undefined || v === null || (typeof v === 'string' && String(v).trim() === '');
  if (empty && out.s && fillIsReddishPreserveFromXlsx(out.s.fill)) {
    const s2 = stripInactiveRedFillFromStyle(out.s);
    if (s2 && Object.keys(s2).length) out.s = s2;
    else delete out.s;
  }
  return out;
};

/**
 * @param {{ stripInactiveFillWhenEmpty?: boolean }} [options] When true (default), clearing a cell removes
 *   reddish inactive-style fill so empty cells are not mistaken for inactive rows.
 */
export const setCellString = (worksheet, rowIndex, colIndex, value, options = {}) => {
  if (colIndex === undefined) return;
  const addr = encodeCell({ r: rowIndex, c: colIndex });
  const prev = worksheet[addr];
  const cell = prev ? { ...prev } : {};
  const preservedStyle = prev?.s;
  const strVal = value == null ? '' : String(value);
  const isEmpty = strVal.trim() === '';
  cell.v = isEmpty ? '' : strVal;
  cell.t = 's';
  const strip = options.stripInactiveFillWhenEmpty !== false;
  let nextStyle = preservedStyle;
  if (preservedStyle && isEmpty && strip) {
    nextStyle = stripInactiveRedFillFromStyle(preservedStyle);
  }
  if (nextStyle && Object.keys(nextStyle).length > 0) cell.s = nextStyle;
  else delete cell.s;
  worksheet[addr] = cell;
};

export const clearRowColumns = (worksheet, rowIndex, columnIndexes) => {
  columnIndexes.forEach((colIndex) => setCellString(worksheet, rowIndex, colIndex, ''));
};

export const rowHasOriginalData = (worksheet, rowIndex, range, originalLastCol) => {
  for (let c = range.s.c; c <= originalLastCol; c += 1) {
    if (getCleanCellValue(worksheet, rowIndex, c)) return true;
  }
  return false;
};

export const readRowFields = (worksheet, rowIndex, fieldColumns, cleaners = {}) => {
  const out = {};
  Object.entries(fieldColumns).forEach(([field, colIndex]) => {
    const raw = getCellValue(worksheet, rowIndex, colIndex);
    const clean = cleaners[field];
    out[field] = clean ? clean(raw) : raw ?? '';
  });
  return out;
};

export const cloneWorksheetRowBelow = (
  worksheet,
  range,
  sourceRowIndex,
  rowDataOverride = new Map(),
) => {
  const insertRowIndex = sourceRowIndex + 1;
  for (let rowIndex = range.e.r; rowIndex >= insertRowIndex; rowIndex -= 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const fromAddr = encodeCell({ r: rowIndex, c });
      const toAddr = encodeCell({ r: rowIndex + 1, c });
      const sourceCell = worksheet[fromAddr];
      if (sourceCell) worksheet[toAddr] = { ...sourceCell };
      else delete worksheet[toAddr];
    }
    if (worksheet['!rows']?.[rowIndex]) {
      worksheet['!rows'][rowIndex + 1] = { ...worksheet['!rows'][rowIndex] };
    } else if (worksheet['!rows']) {
      delete worksheet['!rows'][rowIndex + 1];
    }
  }
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const srcAddr = encodeCell({ r: sourceRowIndex, c });
    const destAddr = encodeCell({ r: insertRowIndex, c });
    const srcCell = worksheet[srcAddr];
    if (srcCell) worksheet[destAddr] = { ...srcCell };
    else delete worksheet[destAddr];
  }
  rowDataOverride.forEach((value, colIndex) => {
    setCellString(worksheet, insertRowIndex, colIndex, value);
  });
  range.e.r += 1;
  worksheet['!ref'] =
    `${colToLetters(range.s.c)}${range.s.r + 1}:${colToLetters(range.e.c)}${range.e.r + 1}`;
  return insertRowIndex;
};

/**
 * Remove one row (0-based index) from a sparse SheetJS-style worksheet: shift rows below up,
 * clear the former last row, shrink range and !ref. Used after inserts/clones left blank rows.
 */
export const deleteWorksheetRowAt = (worksheet, range, rowIndex) => {
  if (rowIndex < range.s.r || rowIndex > range.e.r) return;
  for (let r = rowIndex; r < range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const fromAddr = encodeCell({ r: r + 1, c });
      const toAddr = encodeCell({ r, c });
      const sourceCell = worksheet[fromAddr];
      if (sourceCell) worksheet[toAddr] = { ...sourceCell };
      else delete worksheet[toAddr];
    }
    if (worksheet['!rows']) {
      if (worksheet['!rows'][r + 1]) {
        worksheet['!rows'][r] = { ...worksheet['!rows'][r + 1] };
      } else {
        delete worksheet['!rows'][r];
      }
    }
  }
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    delete worksheet[encodeCell({ r: range.e.r, c })];
  }
  if (worksheet['!rows']) {
    delete worksheet['!rows'][range.e.r];
  }
  range.e.r -= 1;
  worksheet['!ref'] =
    `${colToLetters(range.s.c)}${range.s.r + 1}:${colToLetters(range.e.c)}${range.e.r + 1}`;
};
