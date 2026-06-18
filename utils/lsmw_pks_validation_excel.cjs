/**
 * Excel ringkasan validasi EQUNR / func loc untuk LSMW per PKS.
 */
const ExcelJS = require("exceljs");

const EXCEL_HEADERS = [
    "Nama Sub Folder",
    "Nama File Excel",
    "Jumlah EQUNR Ditemukan",
    "Jumlah EQUNR Tidak Ditemukan",
    "Plant",
    "Valid Func Loc",
    "Invalid Func Loc",
    "Invalid Func Loc List"
];

function rowToExcelValues(row) {
    return [
        row.subFolder ?? "",
        row.fileName ?? "",
        row.equnrFound ?? 0,
        row.equnrNotFound ?? 0,
        row.plant ?? "",
        row.validFuncLoc ?? 0,
        row.invalidFuncLoc ?? 0,
        row.invalidFuncLocList ?? ""
    ];
}

/**
 * @param {string} filePath
 * @param {Array<object>} rows
 */
async function writePksValidationExcel(filePath, rows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Validasi", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(EXCEL_HEADERS);
    ws.getRow(1).font = { bold: true };

    for (const row of rows) {
        ws.addRow(rowToExcelValues(row));
    }

    ws.columns = [
        { width: 22 },
        { width: 48 },
        { width: 28 },
        { width: 32 },
        { width: 12 },
        { width: 16 },
        { width: 18 },
        { width: 56 }
    ];

    if (rows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + rows.length, column: 8 }
        };
    }

    await wb.xlsx.writeFile(filePath);
}

module.exports = {
    EXCEL_HEADERS,
    writePksValidationExcel
};
