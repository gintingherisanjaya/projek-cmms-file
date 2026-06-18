/**
 * Tulis output LSMW measuring point dari template (header baris 1–3 tetap).
 */

const ExcelJS = require("exceljs");

const OUTPUT_START_ROW = 4;

/**
 * @param {string} outPath
 * @param {string} templatePath
 * @param {Array<Record<string, unknown>>} outputRows
 * @param {string[]} columnOrder
 */
async function writeMeasuringPointExcel(
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
            row.getCell(col).value = values[colName];
        }
        row.commit();
    }

    await wb.xlsx.writeFile(outPath);
}

module.exports = {
    OUTPUT_START_ROW,
    writeMeasuringPointExcel
};
