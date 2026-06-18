/**
 * Batas panjang karakter dari template_measuring_point.xlsx (baris 1 = nama, baris 3 = max).
 */

const path = require("path");
const { cellPlainText, loadTemplateLimits } = require("./template_lsmw_pure_check.cjs");

const DEFAULT_TEMPLATE_PATH = path.join(
    __dirname,
    "..",
    "template_measuring_point.xlsx"
);

const TEMPLATE_HEADER_ROW = 1;
const TEMPLATE_MAX_LEN_ROW = 3;

/**
 * @param {string} [templatePath]
 */
async function loadMeasuringPointTemplateLimits(templatePath = DEFAULT_TEMPLATE_PATH) {
    return loadTemplateLimits(
        templatePath,
        TEMPLATE_HEADER_ROW,
        TEMPLATE_MAX_LEN_ROW
    );
}

/**
 * @param {Record<string, unknown>} valuesByColumn
 * @param {Record<string, number|null>} limitByTemplate
 * @param {{ fileName: string, sourceExcelRow: number }} meta
 * @returns {Array<{ column: string, maxLength: number, actualLength: number, value: string, fileName: string, sourceExcelRow: number }>}
 */
function validateOutputRow(valuesByColumn, limitByTemplate, meta) {
    const violations = [];

    for (const [column, max] of Object.entries(limitByTemplate)) {
        if (!max) continue;
        const raw = valuesByColumn[column];
        const text = cellPlainText(raw);
        const len = text.length;
        if (len > max) {
            violations.push({
                column,
                maxLength: max,
                actualLength: len,
                value: text.length > 60 ? `${text.slice(0, 60)}…` : text,
                fileName: meta.fileName,
                sourceExcelRow: meta.sourceExcelRow
            });
        }
    }

    return violations;
}

module.exports = {
    DEFAULT_TEMPLATE_PATH,
    TEMPLATE_HEADER_ROW,
    TEMPLATE_MAX_LEN_ROW,
    loadMeasuringPointTemplateLimits,
    validateOutputRow
};
