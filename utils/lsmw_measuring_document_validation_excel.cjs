/**
 * Excel detail validasi kolom wajib measuring document.
 */
const ExcelJS = require("exceljs");

const EXCEL_HEADERS = [
    "Regional",
    "Nama PKS",
    "Plant",
    "Counter Reading",
    "Equipment Number",
    "Measuring point",
    "Alasan tidak valid"
];

function rowToExcelValues(row) {
    return [
        row.regional ?? "",
        row.fileName ?? "",
        row.plant ?? "",
        row.counterReading ?? "",
        row.equipmentNumber ?? "",
        row.measuringPoint ?? "",
        row.message ?? ""
    ];
}

/**
 * @param {string} filePath
 * @param {Array<object>} rows
 */
async function writeMeasuringDocumentValidationExcel(filePath, rows) {
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
        { width: 14 },
        { width: 42 },
        { width: 10 },
        { width: 18 },
        { width: 18 },
        { width: 18 },
        { width: 56 }
    ];

    if (rows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + rows.length, column: EXCEL_HEADERS.length }
        };
    }

    await wb.xlsx.writeFile(filePath);
}

module.exports = {
    EXCEL_HEADERS,
    writeMeasuringDocumentValidationExcel
};
