/**
 * Baris dengan kolom EQUIPMENT GROUP AFTER ber-background merah → di-skip.
 */

const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const {
    isCellRedFill,
    isRedGoogleBackground,
    isRedXlsxRgbString,
    columnIndexToLetter
} = require("./lsmw_cell_fill.cjs");

const GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet";

function redRowsFromXlsxSheetjs(
    sourcePath,
    colIndex0,
    dataRowCount,
    dataStartExcelRow
) {
    const red = new Set();
    if (colIndex0 === undefined || dataRowCount <= 0) return red;

    let wb;
    try {
        wb = XLSX.readFile(sourcePath, { cellStyles: true });
    } catch {
        return red;
    }

    const sheet = wb.Sheets[wb.SheetNames[0]];
    const styles = wb.Styles;
    if (!sheet || !styles?.CellXf || !styles?.Fills) return red;

    for (let i = 0; i < dataRowCount; i++) {
        const r0 = dataStartExcelRow - 1 + i;
        const addr = XLSX.utils.encode_cell({ r: r0, c: colIndex0 });
        const cell = sheet[addr];
        if (!cell || cell.s === undefined || cell.s === null) continue;

        const xf = styles.CellXf[cell.s];
        if (!xf || xf.fillId === undefined) continue;

        const fill = styles.Fills[xf.fillId];
        if (!fill) continue;

        const rgb = fill.fgColor?.rgb || fill.bgColor?.rgb;
        if (isRedXlsxRgbString(rgb)) red.add(i);
    }

    return red;
}

async function redRowsFromXlsxExceljs(
    sourcePath,
    colIndex0,
    dataRowCount,
    dataStartExcelRow
) {
    const red = new Set();
    if (colIndex0 === undefined || dataRowCount <= 0) return red;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(sourcePath);
    const ws = wb.worksheets[0];

    for (let i = 0; i < dataRowCount; i++) {
        const cell = ws
            .getRow(dataStartExcelRow + i)
            .getCell(colIndex0 + 1);
        if (isCellRedFill(cell)) red.add(i);
    }
    return red;
}

async function redRowsFromGoogleSheet(
    sheetsApi,
    spreadsheetId,
    colIndex0,
    dataRowCount,
    dataStartExcelRow
) {
    const red = new Set();
    if (!sheetsApi || colIndex0 === undefined || dataRowCount <= 0) {
        return red;
    }

    const meta = await sheetsApi.spreadsheets.get({
        spreadsheetId,
        fields: "sheets.properties.title"
    });
    const sheetTitle =
        meta.data.sheets?.[0]?.properties?.title ?? "Sheet1";
    const colLetter = columnIndexToLetter(colIndex0);
    const endRow = dataStartExcelRow + dataRowCount - 1;
    const safeTitle = sheetTitle.replace(/'/g, "''");
    const range = `'${safeTitle}'!${colLetter}1:${colLetter}${endRow}`;

    const res = await sheetsApi.spreadsheets.get({
        spreadsheetId,
        ranges: [range],
        fields:
            "sheets(data(rowData(values(userEnteredFormat.backgroundColor,effectiveFormat.backgroundColor))))"
    });

    const rowData = res.data.sheets?.[0]?.data?.[0]?.rowData ?? [];

    for (let i = 0; i < dataRowCount; i++) {
        const gridIdx = dataStartExcelRow - 1 + i;
        const values = rowData[gridIdx]?.values ?? [];
        const cell = values[0];
        if (!cell) continue;
        const bg =
            cell.effectiveFormat?.backgroundColor ??
            cell.userEnteredFormat?.backgroundColor;
        if (isRedGoogleBackground(bg)) red.add(i);
    }

    return red;
}

/**
 * @returns {Promise<Set<number>>} index baris dalam dataRows (0-based)
 */
async function collectRedEquipGroupRowIndices({
    sourcePath,
    colIndex0,
    dataRowCount,
    dataStartExcelRow,
    driveFile,
    sheetsApi
}) {
    const red = new Set();
    if (colIndex0 === undefined || dataRowCount <= 0) return red;

    if (
        driveFile?.mimeType === GOOGLE_SHEET_MIME &&
        driveFile.id &&
        sheetsApi
    ) {
        try {
            const fromSheets = await redRowsFromGoogleSheet(
                sheetsApi,
                driveFile.id,
                colIndex0,
                dataRowCount,
                dataStartExcelRow
            );
            for (const i of fromSheets) red.add(i);
        } catch (err) {
            console.warn(
                "  [red-skip] Sheets API:",
                err.message
            );
        }
    }

    for (const i of redRowsFromXlsxSheetjs(
        sourcePath,
        colIndex0,
        dataRowCount,
        dataStartExcelRow
    )) {
        red.add(i);
    }

    const fromExceljs = await redRowsFromXlsxExceljs(
        sourcePath,
        colIndex0,
        dataRowCount,
        dataStartExcelRow
    );
    for (const i of fromExceljs) red.add(i);

    return red;
}

module.exports = {
    GOOGLE_SHEET_MIME,
    collectRedEquipGroupRowIndices
};
