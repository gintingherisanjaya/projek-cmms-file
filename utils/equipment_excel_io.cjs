/**
 * Baca/tulis sel Excel konsisten untuk alur equipment-gathering (ExcelJS saja).
 */
const ExcelJS = require("exceljs");

/**
 * ExcelJS cell value → nilai plain (richText, formula result, Date).
 * @param {unknown} value
 * @returns {unknown}
 */
function excelCellToPlain(value) {
    const extract = v => {
        if (v === null || v === undefined) return "";
        const t = typeof v;
        if (t === "string" || t === "number" || t === "boolean") return v;
        if (v instanceof Date) return v;
        if (t === "object") {
            if (Array.isArray(v.richText)) {
                return v.richText
                    .map(rt => (rt && typeof rt.text === "string" ? rt.text : ""))
                    .join("");
            }
            if (typeof v.text === "string") return v.text;
            if (v.result !== undefined && v.result !== null) {
                return extract(v.result);
            }
        }
        return v;
    };

    return extract(value);
}

/**
 * @param {import('exceljs').Worksheet} ws
 * @param {number} maxScanRows
 */
function detectMaxColumn(ws, maxScanRows = 30) {
    let maxCol = Math.max(ws.actualColumnCount ?? 0, ws.columnCount ?? 0);

    if (ws.dimensions?.right) {
        maxCol = Math.max(maxCol, ws.dimensions.right);
    }

    const rowLimit = Math.min(ws.rowCount || 0, maxScanRows);
    for (let r = 1; r <= rowLimit; r++) {
        const row = ws.getRow(r);
        if (!row) continue;
        maxCol = Math.max(maxCol, row.cellCount ?? 0, row.actualCellCount ?? 0);
    }

    return Math.max(maxCol, 1);
}

/**
 * Jumlah baris efektif worksheet (rowCount / dimensions).
 * @param {import('exceljs').Worksheet} ws
 */
function getWorksheetDataRowCount(ws) {
    const rowCount = ws.rowCount || 0;
    const bottom = ws.dimensions?.bottom ?? 0;
    return Math.max(rowCount, bottom);
}

/**
 * Baca worksheet ke array 2D dengan lebar kolom tetap (hindari row.cellCount per baris).
 * @param {import('exceljs').Worksheet} ws
 * @param {{ maxRows?: number, maxCol?: number, headerScanRows?: number }} [options]
 * @returns {Array<Array<unknown>>}
 */
function worksheetToDenseRows(ws, options = {}) {
    const sheetRows = getWorksheetDataRowCount(ws);
    const maxRows = options.maxRows ?? sheetRows;
    const maxCol =
        options.maxCol ?? detectMaxColumn(ws, options.headerScanRows ?? 30);
    const rows = [];
    const rowCount = Math.min(sheetRows, maxRows);

    for (let r = 1; r <= rowCount; r++) {
        const row = ws.getRow(r);
        const arr = [];
        for (let c = 1; c <= maxCol; c++) {
            arr[c - 1] = excelCellToPlain(row.getCell(c).value);
        }
        rows.push(arr);
    }

    return rows;
}

/**
 * @param {string} localPath
 * @param {{ maxRows?: number, headerScanRows?: number }} [options]
 * @returns {Promise<
 *   | { ok: true, rawRows: Array<Array<unknown>> }
 *   | { ok: false, reason: string }
 * >}
 */
async function readSpreadsheetRows(localPath, options = {}) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(localPath);
    const ws = wb.worksheets[0];
    if (!ws) {
        return { ok: false, reason: "Sheet kosong" };
    }

    const rawRows = worksheetToDenseRows(ws, options);
    return { ok: true, rawRows };
}

module.exports = {
    excelCellToPlain,
    detectMaxColumn,
    getWorksheetDataRowCount,
    worksheetToDenseRows,
    readSpreadsheetRows
};
