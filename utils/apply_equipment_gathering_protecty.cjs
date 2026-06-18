/**
 * Terapkan data gathering ke workbook Protect Y: clear semua before, lalu fill by primary key.
 */
const ExcelJS = require("exceljs");
const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");
const { findDataLayout, normalizeHeader } = require("./lsmw_cell_fill.cjs");
const {
    PROTECT_Y_APPLY_HEADERS,
    findFunclocDescAfterLevel123ColumnIndex,
    resolveProtectYApplyColumnIndex
} = require("./equipment_gathering_columns.cjs");
const { normalizeFunclocDesc } = require("./string_similarity.cjs");
const {
    worksheetToDenseRows,
    getWorksheetDataRowCount,
    excelCellToPlain
} = require("./equipment_excel_io.cjs");

function setCellValue(cell, value) {
    if (value === null || value === undefined || value === "") {
        cell.value = null;
        return;
    }
    cell.value = value;
}

/**
 * @param {import('exceljs').Row} excelRow
 * @param {Map<string, number>} applyColIndices
 */
function clearBeforeCells(excelRow, applyColIndices) {
    for (const [, colIdx] of applyColIndices) {
        setCellValue(excelRow.getCell(colIdx + 1), "");
    }
}

/**
 * @param {import('exceljs').Worksheet} ws
 * @param {Map<string, Record<string, unknown>>} gatheringIndex
 * @param {{ regional?: string, pom?: string }} [fileMeta]
 * @returns {{ matched: number, unmatched: number, emptyKey: number, clearedBeforeRows: number, skippedDuplicateProtectYRows: number, regionalPomFilledRows: number, missingApplyCols: string[], missingRegionalPom: boolean, appliedKeys: Set<string> }}
 */
function cellPlainAt(ws, rowIndex0, colIndex0) {
    const row = ws.getRow(rowIndex0 + 1);
    return excelCellToPlain(row.getCell(colIndex0 + 1).value);
}

function applyGatheringToWorksheet(ws, gatheringIndex, fileMeta = {}) {
    const sheetRowCount = getWorksheetDataRowCount(ws);
    const rows = worksheetToDenseRows(ws, { maxRows: sheetRowCount });
    const layout = findDataLayout(rows);
    const headerRow = rows[layout.headerRowIndex];
    if (!headerRow) {
        throw new Error("Header Protect Y tidak ditemukan");
    }

    const colByHeader = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFunclocAfter = findFunclocDescAfterLevel123ColumnIndex(headerRow);

    if (idxFunclocAfter === undefined) {
        throw new Error("Kolom FUNCTLOC DESC. AFTER tidak ditemukan di Protect Y");
    }

    const missingApplyCols = [];
    const applyColIndices = new Map();

    for (const h of PROTECT_Y_APPLY_HEADERS) {
        const idx = resolveProtectYApplyColumnIndex(headerRow, colByHeader, h);
        if (idx === undefined) {
            missingApplyCols.push(h);
        } else {
            applyColIndices.set(h, idx);
        }
    }

    if (missingApplyCols.length > 0) {
        return {
            matched: 0,
            unmatched: 0,
            emptyKey: 0,
            clearedBeforeRows: 0,
            skippedDuplicateProtectYRows: 0,
            regionalPomFilledRows: 0,
            missingApplyCols,
            missingRegionalPom: false,
            appliedKeys: new Set()
        };
    }

    const fileRegional = String(fileMeta.regional ?? "").trim();
    const filePom = String(fileMeta.pom ?? "").trim();
    if (!fileRegional || !filePom) {
        return {
            matched: 0,
            unmatched: 0,
            emptyKey: 0,
            clearedBeforeRows: 0,
            skippedDuplicateProtectYRows: 0,
            regionalPomFilledRows: 0,
            missingApplyCols: [],
            missingRegionalPom: true,
            appliedKeys: new Set()
        };
    }

    const idxRegionalApply = applyColIndices.get("REGIONAL");
    const idxPomApply = applyColIndices.get("POM");

    const dataStart = layout.headerRowIndex + 1;
    const dataEnd = Math.max(sheetRowCount, rows.length);
    let clearedBeforeRows = 0;

    for (let r = dataStart; r < dataEnd; r++) {
        const excelRow = ws.getRow(r + 1);
        clearBeforeCells(excelRow, applyColIndices);
        clearedBeforeRows += 1;
    }

    let matched = 0;
    let unmatched = 0;
    let emptyKey = 0;
    let skippedDuplicateProtectYRows = 0;
    /** @type {Set<string>} */
    const appliedKeys = new Set();
    /** @type {Set<string>} */
    const keysAlreadyApplied = new Set();

    for (let r = dataStart; r < dataEnd; r++) {
        const excelRow = ws.getRow(r + 1);
        const sourceRow = rows[r];
        const afterRaw =
            sourceRow !== undefined
                ? sourceRow[idxFunclocAfter]
                : cellPlainAt(ws, r, idxFunclocAfter);

        const key = normalizeFunclocDesc(afterRaw);
        if (!key) {
            emptyKey += 1;
            continue;
        }

        if (keysAlreadyApplied.has(key)) {
            skippedDuplicateProtectYRows += 1;
            continue;
        }

        const gatherRow = gatheringIndex.get(key);
        if (!gatherRow) {
            unmatched += 1;
            continue;
        }

        for (const [h, colIdx] of applyColIndices) {
            setCellValue(excelRow.getCell(colIdx + 1), gatherRow[h]);
        }
        matched += 1;
        appliedKeys.add(key);
        keysAlreadyApplied.add(key);
    }

    let regionalPomFilledRows = 0;
    for (let r = dataStart; r < dataEnd; r++) {
        const excelRow = ws.getRow(r + 1);
        if (idxRegionalApply !== undefined) {
            setCellValue(excelRow.getCell(idxRegionalApply + 1), fileRegional);
        }
        if (idxPomApply !== undefined) {
            setCellValue(excelRow.getCell(idxPomApply + 1), filePom);
        }
        regionalPomFilledRows += 1;
    }

    return {
        matched,
        unmatched,
        emptyKey,
        clearedBeforeRows,
        skippedDuplicateProtectYRows,
        regionalPomFilledRows,
        missingApplyCols,
        missingRegionalPom: false,
        appliedKeys
    };
}

/**
 * @param {string} localPath
 * @param {Map<string, Record<string, unknown>>} gatheringIndex
 * @param {{ regional?: string, pom?: string }} [fileMeta]
 */
async function applyGatheringToProtectYFile(localPath, gatheringIndex, fileMeta = {}) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(localPath);
    const ws = wb.worksheets[0];
    if (!ws) {
        return { ok: false, reason: "Sheet kosong" };
    }

    const stats = applyGatheringToWorksheet(ws, gatheringIndex, fileMeta);
    if (stats.missingRegionalPom) {
        return {
            ok: false,
            reason: "REGIONAL atau POM kosong — wajib diisi per PKS",
            appliedKeys: stats.appliedKeys ?? new Set(),
            stats
        };
    }
    if (stats.missingApplyCols?.length > 0) {
        return {
            ok: false,
            reason: `Kolom tidak ada di Protect Y: ${stats.missingApplyCols.join(", ")}`,
            appliedKeys: stats.appliedKeys ?? new Set(),
            stats
        };
    }

    await wb.xlsx.writeFile(localPath);
    return { ok: true, stats, appliedKeys: stats.appliedKeys };
}

module.exports = {
    applyGatheringToWorksheet,
    applyGatheringToProtectYFile,
    clearBeforeCells
};
