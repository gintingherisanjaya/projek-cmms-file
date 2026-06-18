/**
 * Baca struktur template tasklist & tulis output (sheet HEADER + OPERATION + SUB_OPERATION).
 */

const ExcelJS = require("exceljs");

const HEADER_SHEET_NAME = "HEADER";
const OPERATION_SHEET_NAME = "OPERATION";
const SUB_OPERATION_SHEET_NAME = "SUB_OPERATION";
const HEADER_OUTPUT_START_ROW = 4;
const OPERATION_OUTPUT_START_ROW = 3;
const SUB_OPERATION_OUTPUT_START_ROW = 5;
const HEADER_NAME_ROW = 2;
const OPERATION_NAME_ROW = 2;
const SUB_OPERATION_NAME_ROW = 1;

function cellHeaderText(value) {
    if (value === null || value === undefined) return "";
    return String(value)
        .replace(/\r\n/g, " ")
        .replace(/\n/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function readRowColumnNames(sheet, rowNumber, maxCol = 80) {
    const row = sheet.getRow(rowNumber);
    const names = [];
    for (let c = 1; c <= maxCol; c += 1) {
        const text = cellHeaderText(row.getCell(c).value);
        if (!text) {
            if (names.length > 0 && c > names.length + 5) break;
            continue;
        }
        names.push(text);
    }
    return names;
}

function findHourMeterColumns(operationColumnOrder) {
    return operationColumnOrder.filter(name => /^\d+\s*HM$/i.test(name));
}

/**
 * @param {string} templatePath
 */
async function loadTasklistTemplateLayout(templatePath) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);

    const headerSheet = wb.getWorksheet(HEADER_SHEET_NAME);
    const operationSheet = wb.getWorksheet(OPERATION_SHEET_NAME);
    const subOperationSheet = wb.getWorksheet(SUB_OPERATION_SHEET_NAME);

    if (!headerSheet) {
        throw new Error(`Template tidak memiliki sheet ${HEADER_SHEET_NAME}`);
    }
    if (!operationSheet) {
        throw new Error(`Template tidak memiliki sheet ${OPERATION_SHEET_NAME}`);
    }
    if (!subOperationSheet) {
        throw new Error(
            `Template tidak memiliki sheet ${SUB_OPERATION_SHEET_NAME}`
        );
    }

    const headerColumnOrder = readRowColumnNames(
        headerSheet,
        HEADER_NAME_ROW,
        20
    );
    const operationColumnOrder = readRowColumnNames(
        operationSheet,
        OPERATION_NAME_ROW,
        80
    );
    const subOperationColumnOrder = readRowColumnNames(
        subOperationSheet,
        SUB_OPERATION_NAME_ROW,
        30
    );
    const hourMeterColumns = findHourMeterColumns(operationColumnOrder);
    const subOperationHourMeterColumns = findHourMeterColumns(
        subOperationColumnOrder
    );

    return {
        headerColumnOrder,
        operationColumnOrder,
        hourMeterColumns,
        subOperationColumnOrder,
        subOperationHourMeterColumns
    };
}

function clearSheetBody(sheet, startRow, columnCount, lastRow) {
    const end = Math.max(lastRow || startRow, startRow);
    for (let r = startRow; r <= end; r += 1) {
        const row = sheet.getRow(r);
        for (let c = 1; c <= columnCount; c += 1) {
            const cell = row.getCell(c);
            cell.value = null;
            if (cell.font?.bold) {
                cell.font = { ...cell.font, bold: false };
            }
        }
        row.commit();
    }
}

function writeRowValues(sheet, rowIndex, valuesByColName, columnOrder, colIndexByName) {
    const row = sheet.getRow(rowIndex);
    for (const colName of columnOrder) {
        const col = colIndexByName[colName];
        if (!col) continue;
        if (Object.prototype.hasOwnProperty.call(valuesByColName, colName)) {
            row.getCell(col).value = valuesByColName[colName];
        }
    }
    row.commit();
}

function applyRowBold(sheet, rowIndex, columnCount, bold) {
    if (!bold) return;
    const row = sheet.getRow(rowIndex);
    for (let c = 1; c <= columnCount; c += 1) {
        const cell = row.getCell(c);
        const prev = cell.font ?? {};
        cell.font = { ...prev, bold: true };
    }
    row.commit();
}

/**
 * @param {string} outPath
 * @param {string} templatePath
 * @param {{
 *   headerRows: Array<Record<string, unknown>>,
 *   operationRows: Array<Record<string, unknown>>,
 *   operationRowBoldFlags?: boolean[],
 *   subOperationRows?: Array<Record<string, unknown>>,
 *   headerColumnOrder: string[],
 *   operationColumnOrder: string[],
 *   subOperationColumnOrder: string[]
 * }} data
 */
async function writeTasklistExcel(outPath, templatePath, data) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);

    const headerSheet = wb.getWorksheet(HEADER_SHEET_NAME);
    const operationSheet = wb.getWorksheet(OPERATION_SHEET_NAME);
    const subOperationSheet = wb.getWorksheet(SUB_OPERATION_SHEET_NAME);
    if (!headerSheet || !operationSheet || !subOperationSheet) {
        throw new Error(
            "Template tasklist tidak lengkap (HEADER/OPERATION/SUB_OPERATION)"
        );
    }

    headerSheet.conditionalFormattings = [];
    operationSheet.conditionalFormattings = [];
    subOperationSheet.conditionalFormattings = [];

    const headerColIndex = {};
    data.headerColumnOrder.forEach((name, i) => {
        headerColIndex[name] = i + 1;
    });

    const operationColIndex = {};
    data.operationColumnOrder.forEach((name, i) => {
        operationColIndex[name] = i + 1;
    });

    const subOperationColIndex = {};
    data.subOperationColumnOrder.forEach((name, i) => {
        subOperationColIndex[name] = i + 1;
    });

    clearSheetBody(
        headerSheet,
        HEADER_OUTPUT_START_ROW,
        data.headerColumnOrder.length,
        headerSheet.rowCount
    );
    clearSheetBody(
        operationSheet,
        OPERATION_OUTPUT_START_ROW,
        data.operationColumnOrder.length,
        operationSheet.rowCount
    );
    clearSheetBody(
        subOperationSheet,
        SUB_OPERATION_OUTPUT_START_ROW,
        data.subOperationColumnOrder.length,
        subOperationSheet.rowCount
    );

    let headerRowNum = HEADER_OUTPUT_START_ROW;
    for (let i = 0; i < data.headerRows.length; i += 1) {
        const values = { ...data.headerRows[i] };
        const noCol = data.headerColumnOrder[0];
        if (noCol) values[noCol] = i + 1;

        writeRowValues(
            headerSheet,
            headerRowNum,
            values,
            data.headerColumnOrder,
            headerColIndex
        );
        headerRowNum += 1;
    }

    let opRowNum = OPERATION_OUTPUT_START_ROW;
    const boldFlags = data.operationRowBoldFlags ?? [];
    for (let i = 0; i < data.operationRows.length; i += 1) {
        writeRowValues(
            operationSheet,
            opRowNum,
            data.operationRows[i],
            data.operationColumnOrder,
            operationColIndex
        );
        applyRowBold(
            operationSheet,
            opRowNum,
            data.operationColumnOrder.length,
            Boolean(boldFlags[i])
        );
        opRowNum += 1;
    }

    const subOpRows = data.subOperationRows ?? [];
    let subOpRowNum = SUB_OPERATION_OUTPUT_START_ROW;
    for (let i = 0; i < subOpRows.length; i += 1) {
        writeRowValues(
            subOperationSheet,
            subOpRowNum,
            subOpRows[i],
            data.subOperationColumnOrder,
            subOperationColIndex
        );
        subOpRowNum += 1;
    }

    await wb.xlsx.writeFile(outPath);
}

module.exports = {
    HEADER_OUTPUT_START_ROW,
    OPERATION_OUTPUT_START_ROW,
    SUB_OPERATION_OUTPUT_START_ROW,
    loadTasklistTemplateLayout,
    writeTasklistExcel
};
