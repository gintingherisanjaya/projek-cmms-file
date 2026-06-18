/**
 * Agregat output func-loc dari file Excel lokal (body baris 5+, kolom A = STRNO).
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const {
    parsePksFilePrefix,
    parseRegionalNumber
} = require("./pks_sort_keys.cjs");

const FUNC_LOC_BODY_START_ROW = 5;
const FUNC_LOC_BODY_KEY_COLUMN = 1;

function cellValueToText(value) {
    if (value === null || value === undefined || value === "") return "";
    if (value instanceof Date) return value.toISOString();
    if (typeof value === "object") {
        if (value.text !== undefined) return String(value.text).trim();
        if (value.result !== undefined) return String(value.result).trim();
        if (value.richText) {
            return value.richText.map(part => part.text ?? "").join("").trim();
        }
    }
    return String(value).trim();
}

function isBodyRowEmpty(row, keyColumn) {
    return cellValueToText(row.getCell(keyColumn).value) === "";
}

function copyRowValues(srcRow, destRow) {
    srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        destRow.getCell(colNumber).value = cell.value;
    });
    destRow.commit();
}

function sortPksPaths(paths) {
    return [...paths].sort((a, b) => {
        const byPrefix =
            parsePksFilePrefix(path.basename(a)) -
            parsePksFilePrefix(path.basename(b));
        if (byPrefix !== 0) return byPrefix;
        return path.basename(a).localeCompare(path.basename(b), "id");
    });
}

function sortRegionalAggregatePaths(paths) {
    return [...paths].sort((a, b) => {
        const byRegional =
            parseRegionalNumber(path.basename(a)) -
            parseRegionalNumber(path.basename(b));
        if (byRegional !== 0) return byRegional;
        return path.basename(a).localeCompare(path.basename(b), "id");
    });
}

/**
 * @param {{
 *   templatePath: string,
 *   outputPath: string,
 *   sourcePaths: string[],
 *   bodyStartRow?: number,
 *   keyColumn?: number
 * }} params
 * @returns {Promise<{ bodyRowCount: number }>}
 */
async function mergeTemplateBodiesFromFiles(params) {
    const {
        templatePath,
        outputPath,
        sourcePaths,
        bodyStartRow = FUNC_LOC_BODY_START_ROW,
        keyColumn = FUNC_LOC_BODY_KEY_COLUMN
    } = params;

    const destWorkbook = new ExcelJS.Workbook();
    await destWorkbook.xlsx.readFile(templatePath);
    const destSheet = destWorkbook.worksheets[0];
    let destRow = bodyStartRow;
    let bodyRowCount = 0;

    for (const sourcePath of sourcePaths) {
        if (!sourcePath || !fs.existsSync(sourcePath)) continue;

        const srcWorkbook = new ExcelJS.Workbook();
        await srcWorkbook.xlsx.readFile(sourcePath);
        const srcSheet = srcWorkbook.worksheets[0];
        if (!srcSheet) continue;

        for (let r = bodyStartRow; r <= srcSheet.rowCount; r += 1) {
            const srcRow = srcSheet.getRow(r);
            if (isBodyRowEmpty(srcRow, keyColumn)) break;

            copyRowValues(srcRow, destSheet.getRow(destRow));
            destRow += 1;
            bodyRowCount += 1;
        }
    }

    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    await destWorkbook.xlsx.writeFile(outputPath);

    return { bodyRowCount };
}

module.exports = {
    FUNC_LOC_BODY_START_ROW,
    FUNC_LOC_BODY_KEY_COLUMN,
    mergeTemplateBodiesFromFiles,
    sortPksPaths,
    sortRegionalAggregatePaths
};
