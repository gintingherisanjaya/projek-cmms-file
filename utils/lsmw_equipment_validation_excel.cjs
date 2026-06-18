/**
 * validation.xlsx audit kolom Protect Y untuk lsmw-equipment.
 */
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const {
    PROTECT_Y_CANONICAL_COLUMNS,
    VALIDATION_FILE_HEADER,
    getValidationHeaders
} = require("./lsmw_equipment_column_audit.cjs");

/**
 * @param {{ fileName: string, resolutions: Map<string, string> }} row
 * @returns {string[]}
 */
function rowToExcelValues(row) {
    const values = [row.fileName ?? ""];
    for (const canonical of PROTECT_Y_CANONICAL_COLUMNS) {
        values.push(row.resolutions?.get(canonical) ?? "");
    }
    return values;
}

/**
 * @param {string} filePath
 * @param {Array<{ fileName: string, resolutions: Map<string, string> }>} rows
 */
async function writeEquipmentColumnValidationExcel(filePath, rows) {
    const headers = getValidationHeaders();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Validasi", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(headers);
    ws.getRow(1).font = { bold: true };

    for (const row of rows) {
        ws.addRow(rowToExcelValues(row));
    }

    ws.columns = headers.map((h, i) => ({
        width: i === 0 ? 36 : Math.min(48, Math.max(14, h.length + 2))
    }));

    if (rows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + rows.length, column: headers.length }
        };
    }

    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    await wb.xlsx.writeFile(filePath);
}

module.exports = {
    VALIDATION_FILE_HEADER,
    getValidationHeaders,
    writeEquipmentColumnValidationExcel
};
