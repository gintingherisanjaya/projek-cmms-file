/**
 * Excel hasil check-equipment-loss (per PKS & ringkasan semua PKS).
 */
const ExcelJS = require("exceljs");

const EXCEL_HEADERS = [
    "EQUIPMENT NUMBER",
    "EQKTU BEFORE",
    "COST CENTER BEFORE",
    "NAMA PKS",
    "STATUS"
];

const STATUS_ORDER = { EXIST: 0, MISSING: 1, ANOMALY: 2, UNUSED: 3, SKIPPED: 4 };

function sortRows(rows) {
    return [...rows].sort((a, b) => {
        const byStatus =
            (STATUS_ORDER[a.status] ?? 99) - (STATUS_ORDER[b.status] ?? 99);
        if (byStatus !== 0) return byStatus;

        return String(a.equipmentNumber ?? "").localeCompare(
            String(b.equipmentNumber ?? ""),
            undefined,
            { sensitivity: "base" }
        );
    });
}

function rowToExcelValues(row) {
    return [
        row.equipmentNumber ?? "",
        row.eqktuBefore ?? "",
        row.costCenterBefore ?? "",
        row.namaPks ?? "",
        row.status ?? ""
    ];
}

/**
 * @param {string} filePath
 * @param {object[]} rows
 * @param {object} [meta] reason jika skipped / kosong
 */
async function writeEquipmentLossExcel(filePath, rows, meta = {}) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Equipment Loss", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(EXCEL_HEADERS);
    ws.getRow(1).font = { bold: true };

    const sorted = sortRows(rows);
    for (const row of sorted) {
        ws.addRow(rowToExcelValues(row));
    }

    if (meta.reason && sorted.length === 0) {
        ws.getCell("G1").value = meta.reason;
    }

    ws.columns = [
        { width: 22 },
        { width: 48 },
        { width: 18 },
        { width: 36 },
        { width: 14 }
    ];

    if (sorted.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + sorted.length, column: 5 }
        };
    }

    await wb.xlsx.writeFile(filePath);
}

module.exports = {
    EXCEL_HEADERS,
    sortRows,
    writeEquipmentLossExcel
};
