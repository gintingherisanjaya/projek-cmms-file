/**
 * Excel detail validasi kolom wajib maintenance plan.
 */
const ExcelJS = require("exceljs");

const EXCEL_HEADERS = [
    "Regional",
    "Nama File PKS",
    "Maintenance Plan",
    "Baris Sumber",
    "FUNCLOC",
    "EQUNR",
    "MAINTENANCE_ITEM",
    "POINT",
    "SZAEH",
    "Kolom Bermasalah",
    "Masalah"
];

function rowToExcelValues(row) {
    return [
        row.regional ?? "",
        row.fileName ?? "",
        row.wptxt ?? "",
        row.sourceExcelRow ?? "",
        row.tplnr ?? "",
        row.equnr ?? "",
        row.maintItem ?? "",
        row.point ?? "",
        row.szaeh ?? "",
        row.column ?? "",
        row.message ?? ""
    ];
}

/**
 * @param {string} filePath
 * @param {Array<object>} rows
 */
async function writeMaintenancePlanValidationExcel(filePath, rows) {
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
        { width: 40 },
        { width: 14 },
        { width: 36 },
        { width: 14 },
        { width: 18 },
        { width: 14 },
        { width: 14 },
        { width: 20 },
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
    writeMaintenancePlanValidationExcel
};
