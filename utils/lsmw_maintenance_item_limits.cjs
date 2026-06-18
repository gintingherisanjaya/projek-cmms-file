/**
 * Batas panjang karakter dari template_lsmw_maintenance_item.xlsx (baris 1 = nama, baris 3 = max).
 */

const path = require("path");
const { loadTemplateLimits } = require("./template_lsmw_pure_check.cjs");
const { validateOutputRow } = require("./lsmw_measuring_point_limits.cjs");

const DEFAULT_TEMPLATE_PATH = path.join(
    __dirname,
    "..",
    "template_lsmw_maintenance_item.xlsx"
);

const TEMPLATE_HEADER_ROW = 1;
const TEMPLATE_MAX_LEN_ROW = 3;

/**
 * @param {string} [templatePath]
 */
async function loadMaintenanceItemTemplateLimits(
    templatePath = DEFAULT_TEMPLATE_PATH
) {
    return loadTemplateLimits(
        templatePath,
        TEMPLATE_HEADER_ROW,
        TEMPLATE_MAX_LEN_ROW
    );
}

module.exports = {
    DEFAULT_TEMPLATE_PATH,
    TEMPLATE_HEADER_ROW,
    TEMPLATE_MAX_LEN_ROW,
    loadMaintenanceItemTemplateLimits,
    validateOutputRow
};
