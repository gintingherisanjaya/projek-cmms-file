/**
 * Excel hasil per PKS: semua baris perbandingan (sama + berbeda), sort & merge vertikal.
 */
const ExcelJS = require("exceljs");

const EXCEL_HEADERS = [
    "FUNC LOC",
    "FUNC LOC DESC",
    "EQUIPMENT GROUP",
    "EQUIPMENT GROUP BARU"
];

/** Kolom 3–4: merge jika nilai sama dengan baris di bawahnya. */
const MERGE_COLUMN_NUMBERS = [3, 4];

function rowToExcelValues(item) {
    return [
        item.funcLoc ?? "",
        item.funcDesc ?? "",
        item.equipmentGroupSource ?? "",
        item.equipmentGroupLsmw ?? ""
    ];
}

function sortComparisonRows(items) {
    return [...items].sort((a, b) => {
        const bySource = String(a.equipmentGroupSource ?? "").localeCompare(
            String(b.equipmentGroupSource ?? ""),
            undefined,
            { sensitivity: "base" }
        );
        if (bySource !== 0) return bySource;

        const byLsmw = String(a.equipmentGroupLsmw ?? "").localeCompare(
            String(b.equipmentGroupLsmw ?? ""),
            undefined,
            { sensitivity: "base" }
        );
        if (byLsmw !== 0) return byLsmw;

        return String(a.funcLoc ?? "").localeCompare(
            String(b.funcLoc ?? ""),
            undefined,
            { sensitivity: "base" }
        );
    });
}

function applyVerticalMerges(worksheet, dataStartRow, dataEndRow, colNumbers) {
    if (dataEndRow <= dataStartRow) return;

    for (const col of colNumbers) {
        let blockStart = dataStartRow;

        for (let row = dataStartRow; row <= dataEndRow; row += 1) {
            const cur = String(
                worksheet.getRow(row).getCell(col).value ?? ""
            );
            const below =
                row < dataEndRow
                    ? String(
                          worksheet.getRow(row + 1).getCell(col).value ?? ""
                      )
                    : null;

            if (row === dataEndRow || cur !== below) {
                if (row > blockStart) {
                    worksheet.mergeCells(blockStart, col, row, col);
                    const master = worksheet.getRow(blockStart).getCell(col);
                    master.alignment = {
                        ...(master.alignment ?? {}),
                        vertical: "middle",
                        wrapText: true
                    };
                }
                blockStart = row + 1;
            }
        }
    }
}

/**
 * @param {string} filePath
 * @param {object[]} items — baris perbandingan (field internal)
 * @param {object} [meta] status, reason, pks
 */
async function writeGroupCheckExcel(filePath, items, meta = {}) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Perbandingan", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(EXCEL_HEADERS);
    ws.getRow(1).font = { bold: true };

    const sorted = sortComparisonRows(items);
    const dataStartRow = 2;

    for (const item of sorted) {
        ws.addRow(rowToExcelValues(item));
    }
    const dataEndRow = ws.rowCount;

    if (sorted.length > 0) {
        applyVerticalMerges(ws, dataStartRow, dataEndRow, MERGE_COLUMN_NUMBERS);
    }

    if (meta.status === "skipped" && meta.reason) {
        ws.getCell("F1").value = `skipped: ${meta.reason}`;
    }

    ws.columns = [
        { width: 42 },
        { width: 48 },
        { width: 22 },
        { width: 22 }
    ];

    await wb.xlsx.writeFile(filePath);
}

module.exports = {
    EXCEL_HEADERS,
    sortComparisonRows,
    writeGroupCheckExcel
};
