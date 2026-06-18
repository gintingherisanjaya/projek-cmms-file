/**
 * Baca layout & tulis output LSMW measuring document dari template.
 */

const ExcelJS = require("exceljs");

// Header template hanya baris 1; body mulai baris 2.
const OUTPUT_START_ROW = 2;
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
async function loadMeasuringDocumentColumnOrder(templatePath) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const sheet = wb.worksheets[0];
    if (!sheet) throw new Error("Template tidak memiliki sheet");

    const row = sheet.getRow(COLUMN_NAME_ROW);
    const names = [];
    for (let c = 1; c <= 20; c += 1) {
        const text = cellHeaderText(row.getCell(c).value);
        if (!text) {
            if (names.length > 0 && c > names.length + 3) break;
            continue;
        }
        names.push(text);
    }

    if (names.length === 0) {
        throw new Error(
            "Template measuring document: baris 1 tidak berisi nama kolom"
        );
    }

    return names;
}

/**
 * @param {string} outPath
 * @param {string} templatePath
 * @param {Array<Record<string, unknown>>} outputRows
 * @param {string[]} columnOrder
 */
async function writeMeasuringDocumentExcel(
    outPath,
    templatePath,
    outputRows,
    columnOrder
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

    const existingLast = sheet.rowCount || OUTPUT_START_ROW;
    for (let r = OUTPUT_START_ROW; r <= existingLast; r += 1) {
        const row = sheet.getRow(r);
        for (let c = 1; c <= columnOrder.length; c += 1) {
            row.getCell(c).value = null;
        }
        row.commit();
    }

    let startRow = OUTPUT_START_ROW;
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
    cellHeaderText,
    loadMeasuringDocumentColumnOrder,
    writeMeasuringDocumentExcel
};
