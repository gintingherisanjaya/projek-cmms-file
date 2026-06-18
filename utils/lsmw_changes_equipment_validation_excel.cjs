/**
 * Excel validasi lsmw-changes-equipment: baris tanpa match Truth atau DESC AFTER kosong.
 */
const ExcelJS = require("exceljs");

const EXCEL_HEADERS = [
    "Nama Sub Folder",
    "Nama File Excel",
    "EQUNR",
    "TPLNR",
    "OLD_NAME",
    "NEW_NAME",
    "Keterangan"
];

function rowToExcelValues(row) {
    return [
        row.subFolder ?? "",
        row.fileName ?? "",
        row.equnr ?? "",
        row.tplnr ?? "",
        row.oldName ?? "",
        row.newName ?? "",
        row.note ?? ""
    ];
}

/**
 * @param {string} filePath
 * @param {Array<object>} rows
 */
async function writeChangesEquipmentValidationExcel(filePath, rows) {
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
        { width: 22 },
        { width: 48 },
        { width: 40 },
        { width: 40 },
        { width: 36 }
    ];

    if (rows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + rows.length, column: 7 }
        };
    }

    await wb.xlsx.writeFile(filePath);
}

module.exports = {
    EXCEL_HEADERS,
    writeChangesEquipmentValidationExcel
};
