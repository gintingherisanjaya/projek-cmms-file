/**
 * Baca layout template maintenance item & tulis output (header baris 1–3 tetap).
 */

const ExcelJS = require("exceljs");

const OUTPUT_START_ROW = 4;
const COLUMN_NAME_ROW = 1;

function cellHeaderText(value) {
    if (value === null || value === undefined) return "";
    return String(value)
        .replace(/\r\n/g, " ")
        .replace(/\n/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

/**
 * @param {string} templatePath
 * @returns {Promise<string[]>}
 */
async function loadMaintenanceItemColumnOrder(templatePath) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const sheet = wb.worksheets[0];
    if (!sheet) throw new Error("Template tidak memiliki sheet");

    const row = sheet.getRow(COLUMN_NAME_ROW);
    const names = [];
    for (let c = 1; c <= 30; c += 1) {
        const text = cellHeaderText(row.getCell(c).value);
        if (!text) {
            if (names.length > 0 && c > names.length + 3) break;
            continue;
        }
        names.push(text);
    }

    if (names.length === 0) {
        throw new Error("Template maintenance item: baris 1 tidak berisi nama kolom");
    }

    return names;
}

/**
 * @param {string} outPath
 * @param {string} templatePath
 * @param {Array<Record<string, unknown>>} outputRows
 * @param {string[]} columnOrder
 * @param {number} [outputStartRow]
 */
async function writeMaintenanceItemExcel(
    outPath,
    templatePath,
    outputRows,
    columnOrder,
    outputStartRow = OUTPUT_START_ROW
) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const sheet = wb.worksheets[0];
    if (!sheet) throw new Error("Template tidak memiliki sheet");

    sheet.conditionalFormattings = [];

    const colIndexByName = {};
    for (let c = 0; c < columnOrder.length; c += 1) {
        colIndexByName[columnOrder[c]] = c + 1;
    }

    const existingLast = sheet.rowCount || outputStartRow;
    for (let r = outputStartRow; r <= existingLast; r += 1) {
        const row = sheet.getRow(r);
        for (let c = 1; c <= columnOrder.length; c += 1) {
            row.getCell(c).value = null;
        }
        row.commit();
    }

    let startRow = outputStartRow;
    for (const values of outputRows) {
        const row = sheet.getRow(startRow++);
        for (const colName of columnOrder) {
            const col = colIndexByName[colName];
            if (!col) continue;
            if (Object.prototype.hasOwnProperty.call(values, colName)) {
                row.getCell(col).value = values[colName];
            }
        }
        row.commit();
    }

    await wb.xlsx.writeFile(outPath);
}

module.exports = {
    OUTPUT_START_ROW,
    loadMaintenanceItemColumnOrder,
    writeMaintenanceItemExcel
};
