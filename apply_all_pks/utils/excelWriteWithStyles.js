import ExcelJS from 'exceljs';
import path from 'node:path';
import * as xlsxNs from 'xlsx';
import { COLOR_ORANGE, COLOR_YELLOW } from './validationHelpers.js';
import { ensureDir } from './fileSystem.js';

const xlsx = xlsxNs.default ?? xlsxNs;

const getWorkbookFirstSheetName = (workbook) => {
  if (Array.isArray(workbook?.SheetNames) && workbook.SheetNames.length > 0) {
    return workbook.SheetNames[0];
  }
  if (Array.isArray(workbook?.worksheets) && workbook.worksheets.length > 0) {
    return workbook.worksheets[0]?.name ?? 'Sheet1';
  }
  return 'Sheet1';
};

const cellHasMeaningfulValue = (cell) => {
  if (!cell) return false;
  const value = cell.v;
  if (value === undefined || value === null) return false;
  if (typeof value === 'string') return value.trim() !== '';
  return true;
};

const collectNonEmptyRowIndexes = (worksheet, range) => {
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    let hasValue = false;
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = xlsx.utils.encode_cell({ r, c });
      if (cellHasMeaningfulValue(worksheet[addr])) {
        hasValue = true;
        break;
      }
    }
    if (hasValue) rows.push(r);
  }
  return rows;
};

/** Normalize SheetJS fgColor to 6-char RGB (no alpha). */
const fgColorToRgb6 = (fgColor) => {
  if (!fgColor) return null;
  const raw = fgColor.rgb ?? fgColor;
  if (!raw || typeof raw !== 'string') return null;
  let hex = raw.replace(/^#/, '').toUpperCase();
  if (hex.length === 8) hex = hex.slice(2);
  if (hex.length !== 6) return null;
  return hex;
};

/**
 * True if the source cell fill is a light red / pink error highlight (not row yellow).
 * Used to keep EQUIPMENT GROUP AFTER (etc.) source reds when the row is yellow.
 */
export const fillIsReddishPreserveFromXlsx = (fill) => {
  if (!fill?.fgColor) return false;
  const hex = fgColorToRgb6(fill.fgColor);
  if (!hex) return false;
  if (hex === 'FFC7CE') return true;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return false;
  // Yellow row highlight is ~FFEB9C — high G and relatively close R/G.
  // We only preserve when it's clearly red/pink-ish: R dominates G and B.
  const looksRed =
    r >= 180 &&
    r > g + 20 &&
    r > b + 20 &&
    // allow both darker and light pink variants
    g <= 230 &&
    b <= 240;

  // Extra tolerance for Excel-like light red/pink highlight.
  const looksPinkish = r >= 200 && g <= 210 && b <= 230 && r >= g - 10;

  return looksRed || looksPinkish;
};

/**
 * Write workbook with proper style support using exceljs
 * This function reads the xlsx workbook, converts it to exceljs format,
 * applies row colors, and writes it back with full style support
 *
 * @param {object} [options]
 * @param {number} [options.preserveSourceRedFillColumn] 0-based column index: when row is yellow,
 *   keep original reddish fill on that column (e.g. EQUIPMENT GROUP AFTER) instead of painting yellow.
 * @param {Set<string>} [options.skipValidationFillAddresses] Encoded SheetJS addresses (A1-style)
 *   that keep their worksheet fill when a validation row color would otherwise overwrite them.
 */
export const writeWorkbookWithRowColors = async (
  xlsxWorkbook,
  worksheet,
  outputPath,
  rowColors, // Map of rowIndex -> color (ARGB format)
  options = {},
) => {
  const preserveRedCol = options.preserveSourceRedFillColumn;
  const skipFillAddrs = options.skipValidationFillAddresses;
  const applyValidationRowColors = options.applyValidationRowColors ?? true;
  ensureDir(path.dirname(outputPath));

  // Create new ExcelJS workbook
  const exceljsWorkbook = new ExcelJS.Workbook();

  // Get the sheet name
  const sheetName = getWorkbookFirstSheetName(xlsxWorkbook);
  const exceljsSheet = exceljsWorkbook.addWorksheet(sheetName);

  // Get the range
  const ref = worksheet['!ref'] || 'A1:A1';
  const range = xlsx.utils.decode_range(ref);
  const sourceRows = collectNonEmptyRowIndexes(worksheet, range);

  // Copy column widths
  if (worksheet['!cols']) {
    worksheet['!cols'].forEach((colDef, index) => {
      if (colDef) {
        const col = exceljsSheet.getColumn(index + 1); // ExcelJS uses 1-based indexing
        if (colDef.wch) {
          col.width = colDef.wch;
        } else if (colDef.wpx) {
          // Convert pixels to character width (approximate: 7 pixels per character)
          col.width = colDef.wpx / 7;
        }
        if (colDef.hidden !== undefined) {
          col.hidden = Boolean(colDef.hidden);
        }
      }
    });
  }

  // Copy row heights
  if (worksheet['!rows']) {
    sourceRows.forEach((sourceRowIndex, targetRowIndex) => {
      const rowDef = worksheet['!rows']?.[sourceRowIndex];
      if (rowDef && rowDef.hpt !== undefined) {
        const row = exceljsSheet.getRow(targetRowIndex + 1); // ExcelJS uses 1-based indexing
        row.height = rowDef.hpt; // hpt is already in points
      }
    });
  }

  // Convert xlsx data to exceljs format
  sourceRows.forEach((sourceRowIndex, targetRowIndex) => {
    const row = exceljsSheet.getRow(targetRowIndex + 1); // ExcelJS uses 1-based indexing

    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = xlsx.utils.encode_cell({ r: sourceRowIndex, c });
      const xlsxCell = worksheet[addr];

      if (!xlsxCell) continue;

      const exceljsCell = row.getCell(c + 1); // ExcelJS uses 1-based indexing

      // Copy value
      if (xlsxCell.v !== undefined && xlsxCell.v !== null) {
        exceljsCell.value = xlsxCell.v;
      }

      // Copy existing styles if present
      if (xlsxCell.s) {
        // Copy font styles
        if (xlsxCell.s.font) {
          exceljsCell.font = {
            ...exceljsCell.font,
            ...(xlsxCell.s.font.color && {
              color: { argb: xlsxCell.s.font.color.rgb || xlsxCell.s.font.color },
            }),
            ...(xlsxCell.s.font.bold !== undefined && { bold: xlsxCell.s.font.bold }),
            ...(xlsxCell.s.font.size !== undefined && { size: xlsxCell.s.font.size }),
            ...(xlsxCell.s.font.italic !== undefined && { italic: xlsxCell.s.font.italic }),
            ...(xlsxCell.s.font.underline !== undefined && {
              underline: xlsxCell.s.font.underline,
            }),
          };
        }

        // Copy alignment
        if (xlsxCell.s.alignment) {
          exceljsCell.alignment = { ...exceljsCell.alignment, ...xlsxCell.s.alignment };
        }

        // Copy fill (background color) - but only if no row color override,
        // except: yellow rows may still show source red on preserveRedCol (e.g. EQUIPMENT GROUP AFTER).
        const rawRowColor = rowColors.get(sourceRowIndex);
        const rowColor =
          applyValidationRowColors || rawRowColor === COLOR_ORANGE ? rawRowColor : undefined;
        const preserveRedHere =
          preserveRedCol !== undefined &&
          c === preserveRedCol &&
          fillIsReddishPreserveFromXlsx(xlsxCell.s?.fill);
        const skipFillOverride =
          skipFillAddrs instanceof Set && skipFillAddrs.has(addr);
        if ((!rowColor || preserveRedHere || skipFillOverride) && xlsxCell.s?.fill) {
          if (xlsxCell.s.fill.fgColor) {
            const rgbValue = xlsxCell.s.fill.fgColor.rgb || xlsxCell.s.fill.fgColor;
            if (rgbValue) {
              // Ensure ARGB format (add FF prefix if only 6 digits)
              const argbValue = rgbValue.length === 6 ? `FF${rgbValue}` : rgbValue;
              exceljsCell.fill = {
                type: 'pattern',
                pattern: xlsxCell.s.fill.patternType || 'solid',
                fgColor: { argb: argbValue },
              };
            }
          }
        }

        // Copy borders
        if (xlsxCell.s.border) {
          exceljsCell.border = { ...exceljsCell.border, ...xlsxCell.s.border };
        }
      }
    }

    // Apply row fill color if specified (overrides existing fill)
    const rawRowColor = rowColors.get(sourceRowIndex);
    const rowColor =
      applyValidationRowColors || rawRowColor === COLOR_ORANGE ? rawRowColor : undefined;
    if (rowColor) {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const c = colNumber - 1;
        const addr = xlsx.utils.encode_cell({ r: sourceRowIndex, c });
        const src = worksheet[addr];
        if (skipFillAddrs instanceof Set && skipFillAddrs.has(addr)) {
          return;
        }
        if (
          preserveRedCol !== undefined &&
          c === preserveRedCol &&
          fillIsReddishPreserveFromXlsx(src?.s?.fill)
        ) {
          return;
        }
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: rowColor },
        };
      });
    }
  });

  // Write the file
  await exceljsWorkbook.xlsx.writeFile(outputPath);
};

/**
 * Write workbook with proper style support using exceljs (without row colors)
 * This preserves all existing styles when writing
 *
 * @param {object} [options]
 * @param {boolean} [options.preserveSheetLayout] When true, copy every row from `!ref` in order
 *   (including blank rows) so header position and row gaps match the source worksheet.
 */
export const writeWorkbookWithStyles = async (xlsxWorkbook, worksheet, outputPath, options = {}) => {
  const preserveSheetLayout = options.preserveSheetLayout ?? false;
  ensureDir(path.dirname(outputPath));

  // Create new ExcelJS workbook
  const exceljsWorkbook = new ExcelJS.Workbook();

  // Get the sheet name
  const sheetName = getWorkbookFirstSheetName(xlsxWorkbook);
  const exceljsSheet = exceljsWorkbook.addWorksheet(sheetName);

  // Get the range
  const ref = worksheet['!ref'] || 'A1:A1';
  const range = xlsx.utils.decode_range(ref);
  const sourceRows = preserveSheetLayout
    ? Array.from({ length: range.e.r - range.s.r + 1 }, (_, i) => range.s.r + i)
    : collectNonEmptyRowIndexes(worksheet, range);

  // Copy column widths
  if (worksheet['!cols']) {
    worksheet['!cols'].forEach((colDef, index) => {
      if (colDef) {
        const col = exceljsSheet.getColumn(index + 1); // ExcelJS uses 1-based indexing
        if (colDef.wch) {
          col.width = colDef.wch;
        } else if (colDef.wpx) {
          // Convert pixels to character width (approximate: 7 pixels per character)
          col.width = colDef.wpx / 7;
        }
        if (colDef.hidden !== undefined) {
          col.hidden = Boolean(colDef.hidden);
        }
      }
    });
  }

  // Copy row heights
  if (worksheet['!rows']) {
    sourceRows.forEach((sourceRowIndex, targetRowIndex) => {
      const rowDef = worksheet['!rows']?.[sourceRowIndex];
      if (rowDef && rowDef.hpt !== undefined) {
        const row = exceljsSheet.getRow(targetRowIndex + 1); // ExcelJS uses 1-based indexing
        row.height = rowDef.hpt; // hpt is already in points
      }
    });
  }

  // Convert xlsx data to exceljs format
  sourceRows.forEach((sourceRowIndex, targetRowIndex) => {
    const row = exceljsSheet.getRow(targetRowIndex + 1); // ExcelJS uses 1-based indexing

    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = xlsx.utils.encode_cell({ r: sourceRowIndex, c });
      const xlsxCell = worksheet[addr];

      if (!xlsxCell) continue;

      const exceljsCell = row.getCell(c + 1); // ExcelJS uses 1-based indexing

      // Copy value
      if (xlsxCell.v !== undefined && xlsxCell.v !== null) {
        exceljsCell.value = xlsxCell.v;
      }

      // Copy existing styles if present
      if (xlsxCell.s) {
        // Copy font styles
        if (xlsxCell.s.font) {
          exceljsCell.font = {
            ...exceljsCell.font,
            ...(xlsxCell.s.font.color && {
              color: { argb: xlsxCell.s.font.color.rgb || xlsxCell.s.font.color },
            }),
            ...(xlsxCell.s.font.bold !== undefined && { bold: xlsxCell.s.font.bold }),
            ...(xlsxCell.s.font.size !== undefined && { size: xlsxCell.s.font.size }),
            ...(xlsxCell.s.font.italic !== undefined && { italic: xlsxCell.s.font.italic }),
            ...(xlsxCell.s.font.underline !== undefined && {
              underline: xlsxCell.s.font.underline,
            }),
          };
        }

        // Copy alignment
        if (xlsxCell.s.alignment) {
          exceljsCell.alignment = { ...exceljsCell.alignment, ...xlsxCell.s.alignment };
        }

        // Copy fill (background color) - including header cell colors
        if (xlsxCell.s.fill) {
          if (xlsxCell.s.fill.fgColor) {
            const rgbValue = xlsxCell.s.fill.fgColor.rgb || xlsxCell.s.fill.fgColor;
            if (rgbValue) {
              // Ensure ARGB format (add FF prefix if only 6 digits)
              const argbValue = rgbValue.length === 6 ? `FF${rgbValue}` : rgbValue;
              exceljsCell.fill = {
                type: 'pattern',
                pattern: xlsxCell.s.fill.patternType || 'solid',
                fgColor: { argb: argbValue },
              };
            }
          }
        }

        // Copy borders
        if (xlsxCell.s.border) {
          exceljsCell.border = { ...exceljsCell.border, ...xlsxCell.s.border };
        }
      }
    }
  });

  // Write the file
  await exceljsWorkbook.xlsx.writeFile(outputPath);
};
