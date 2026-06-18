/**
 * Tulis Excel laporan baris gathering yang gagal di-apply ke Protect Y.
 */
const ExcelJS = require("exceljs");
const { OUTPUT_HEADERS } = require("./equipment_gathering_columns.cjs");

const FAILED_APPLY_HEADERS = ["REGIONAL", "PKS", ...OUTPUT_HEADERS];

/**
 * @param {string} filePath
 * @param {Array<{ regional: string, pks: string, values: unknown[] }>} failedRows
 */
async function writeFailedApplyExcel(filePath, failedRows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Failed", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(FAILED_APPLY_HEADERS);
    ws.getRow(1).font = { bold: true };

    for (const { regional, pks, values } of failedRows) {
        ws.addRow([regional, pks, ...values]);
    }

    ws.columns = FAILED_APPLY_HEADERS.map(h => {
        const wide =
            h.includes("FUNCTLOC DESC") ||
            h.includes("EQKTU") ||
            h === "COST CENTER BEFORE" ||
            h === "COST CENTER AFTER";
        return { width: wide ? 48 : h === "REGIONAL" ? 14 : h === "PKS" ? 28 : 22 };
    });

    if (failedRows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + failedRows.length, column: FAILED_APPLY_HEADERS.length }
        };
    }

    await wb.xlsx.writeFile(filePath);
}

module.exports = { writeFailedApplyExcel, FAILED_APPLY_HEADERS };
