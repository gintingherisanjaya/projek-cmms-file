/**
 * Batas panjang karakter sheet SUB_OPERATION di template_lsmw_tasklist.xlsx
 * (baris 1 = nama kolom, baris 4 = max length, body baris 5+).
 */

const path = require("path");
const ExcelJS = require("exceljs");
const { cellPlainText } = require("./template_lsmw_pure_check.cjs");
const { validateOutputRow } = require("./lsmw_measuring_point_limits.cjs");

const DEFAULT_TEMPLATE_PATH = path.join(
    __dirname,
    "..",
    "template_lsmw_tasklist.xlsx"
);

const SUB_OPERATION_SHEET = "SUB_OPERATION";
const HEADER_ROW = 1;
const MAX_LEN_ROW = 4;
const OUTPUT_START_ROW = 5;

function parseMaxLength(raw) {
    if (raw === null || raw === undefined || raw === "") return null;
    const n = parseInt(String(raw).trim(), 10);
    return Number.isFinite(n) && n > 0 ? n : null;
}

async function loadSubOperationTemplateLimits(
    templatePath = DEFAULT_TEMPLATE_PATH
) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const sheet = wb.getWorksheet(SUB_OPERATION_SHEET);
    if (!sheet) {
        throw new Error(`Template tidak memiliki sheet ${SUB_OPERATION_SHEET}`);
    }

    const nameRow = sheet.getRow(HEADER_ROW);
    const maxRow = sheet.getRow(MAX_LEN_ROW);
    const colCount = Math.max(nameRow.cellCount, maxRow.cellCount, 30);
    const columns = [];
    const maxLengths = [];

    for (let c = 1; c <= colCount; c += 1) {
        const name = cellPlainText(nameRow.getCell(c).value).trim();
        if (!name) break;
        columns.push(name);
        maxLengths.push(parseMaxLength(maxRow.getCell(c).value));
    }

    if (columns.length === 0) {
        throw new Error(
            `Kolom SUB_OPERATION tidak ditemukan (baris ${HEADER_ROW})`
        );
    }

    const limitByTemplate = {};
    for (let i = 0; i < columns.length; i += 1) {
        limitByTemplate[columns[i]] = maxLengths[i];
    }

    return { columns, maxLengths, limitByTemplate };
}

function formatCharViolation(v) {
    return (
        `${v.fileName} baris SUB_OPERATION ${v.sourceExcelRow}: kolom ${v.column} ` +
        `(${v.actualLength}/${v.maxLength}) "${v.value}"`
    );
}

module.exports = {
    DEFAULT_TEMPLATE_PATH,
    SUB_OPERATION_SHEET,
    HEADER_ROW,
    MAX_LEN_ROW,
    OUTPUT_START_ROW,
    loadSubOperationTemplateLimits,
    validateOutputRow,
    formatCharViolation
};
