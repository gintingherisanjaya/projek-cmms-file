/**
 * Tulis Excel hasil equipment-gathering.
 */
const ExcelJS = require("exceljs");
const { OUTPUT_HEADERS } = require("./equipment_gathering_columns.cjs");

/**
 * @param {string} filePath
 * @param {Array<Array<unknown>>} dataRows — nilai per baris (urutan OUTPUT_HEADERS)
 */
async function writeGatheringExcel(filePath, dataRows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Data", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(OUTPUT_HEADERS);
    ws.getRow(1).font = { bold: true };

    for (const values of dataRows) {
        ws.addRow(values);
    }

    ws.columns = OUTPUT_HEADERS.map(h => {
        const wide =
            h.includes("FUNCTLOC DESC") ||
            h.includes("EQKTU") ||
            h === "COST CENTER BEFORE" ||
            h === "COST CENTER AFTER";
        return { width: wide ? 48 : h === "REGIONAL" ? 14 : 22 };
    });

    if (dataRows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + dataRows.length, column: OUTPUT_HEADERS.length }
        };
    }

    await wb.xlsx.writeFile(filePath);
}

module.exports = { writeGatheringExcel };
