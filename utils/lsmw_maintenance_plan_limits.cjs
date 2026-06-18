/**
 * Batas panjang karakter & validasi kolom wajib maintenance plan.
 */

const path = require("path");
const { cellPlainText, loadTemplateLimits } = require("./template_lsmw_pure_check.cjs");
const { formatHeriCounterOutput } = require("./heri_sheet_loader.cjs");
const { validateOutputRow } = require("./lsmw_measuring_point_limits.cjs");
const {
    normalizeCounterDisplayForCompare
} = require("./measuring_document_counter_loader.cjs");

const DEFAULT_TEMPLATE_PATH = path.join(
    __dirname,
    "..",
    "template_lsmw_maintenance_plan.xlsx"
);

const TEMPLATE_HEADER_ROW = 1;
const TEMPLATE_MAX_LEN_ROW = 4;
const MAINTENANCE_PLAN_OUTPUT_START_ROW = 5;

const REQUIRED_FIELD_COLUMNS = [
    "EQUNR",
    "MAINTENANCE_ITEM",
    "SZAEH",
    "WPTXT",
    "POINT",
    "Counter Reading"
];

const DEFAULT_STRICT_FIELD_COLUMNS = [...REQUIRED_FIELD_COLUMNS];

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function formatSzaehDisplay(values) {
    const raw = values?.SZAEH;
    if (raw === null || raw === undefined || raw === "") return "0";
    if (typeof raw === "number") {
        return formatHeriCounterOutput(raw) ?? "0";
    }
    const text = String(raw).trim();
    return text || "0";
}

/**
 * @param {{
 *   equnr: string,
 *   equnrFromMapping?: boolean,
 *   maintItemFromIp18?: boolean,
 *   pointFromIk07?: boolean,
 *   counterByEqunr?: Map<string, string>,
 *   values: Record<string, unknown>
 * }} params
 * @param {{ fileName: string, sourceExcelRow: number, tplnr: string, regional?: string }} meta
 */
function validateMaintenancePlanFields(params, meta) {
    const {
        equnr,
        equnrFromMapping = false,
        maintItemFromIp18 = false,
        pointFromIk07 = false,
        counterByEqunr,
        values
    } = params;
    const issues = [];

    const snapshot = {
        regional: meta.regional ?? "",
        fileName: meta.fileName,
        wptxt: cellText(values.WPTXT),
        sourceExcelRow: meta.sourceExcelRow,
        tplnr: meta.tplnr,
        equnr: cellText(equnr),
        maintItem: cellText(values.MAINTENANCE_ITEM),
        point: cellText(values.POINT),
        szaeh: formatSzaehDisplay(values)
    };

    const base = {
        ...snapshot,
        fileName: meta.fileName,
        sourceExcelRow: meta.sourceExcelRow,
        tplnr: meta.tplnr
    };

    if (!equnrFromMapping) {
        issues.push({
            ...base,
            column: "EQUNR",
            message: "EQUNR tidak ditemukan di new-equipment-number.xlsx"
        });
    }

    if (!maintItemFromIp18) {
        issues.push({
            ...base,
            column: "MAINTENANCE_ITEM",
            message: cellText(equnr)
                ? "MAINTENANCE_ITEM tidak ditemukan di maintenance-item-ip18.xlsx"
                : "MAINTENANCE_ITEM kosong (EQUNR tidak tersedia)"
        });
    }

    if (!cellText(values.SZAEH) && values.SZAEH !== 0) {
        issues.push({
            ...base,
            column: "SZAEH",
            message: "SZAEH kosong"
        });
    }

    if (counterByEqunr && cellText(equnr)) {
        const eqKey = cellText(equnr);
        if (counterByEqunr.has(eqKey)) {
            const expected = counterByEqunr.get(eqKey);
            const actual = formatSzaehDisplay(values);
            if (
                normalizeCounterDisplayForCompare(actual) !==
                normalizeCounterDisplayForCompare(expected)
            ) {
                issues.push({
                    ...base,
                    column: "Counter Reading",
                    message: `SZAEH "${actual}" tidak sama dengan Counter Reading measuring-document "${expected}" untuk EQUNR ${eqKey}`
                });
            }
        }
    }

    if (!cellText(values.WPTXT)) {
        issues.push({
            ...base,
            column: "WPTXT",
            message: "WPTXT kosong"
        });
    }

    if (!pointFromIk07) {
        issues.push({
            ...base,
            column: "POINT",
            message: cellText(equnr)
                ? "POINT tidak ditemukan di measuring-item-ik07.xlsx"
                : "POINT kosong (EQUNR tidak tersedia)"
        });
    }

    return issues;
}

/**
 * @param {string} [templatePath]
 */
async function loadMaintenancePlanTemplateLimits(
    templatePath = DEFAULT_TEMPLATE_PATH
) {
    return loadTemplateLimits(
        templatePath,
        TEMPLATE_HEADER_ROW,
        TEMPLATE_MAX_LEN_ROW
    );
}

/**
 * @param {{ column: string, maxLength: number, actualLength: number, value: string, fileName: string, sourceExcelRow: number, tplnr?: string }} v
 */
function formatCharViolation(v) {
    const loc = v.tplnr ? ` FUNCLOC=${v.tplnr}` : "";
    return (
        `${v.fileName} baris sumber ${v.sourceExcelRow}:${loc} kolom ${v.column} ` +
        `(${v.actualLength}/${v.maxLength}) "${v.value}"`
    );
}

/**
 * @param {{ column: string, fileName: string, sourceExcelRow: number, tplnr: string, message?: string }} issue
 */
function formatFieldViolation(issue) {
    const msg = issue.message ? ` — ${issue.message}` : "";
    return (
        `${issue.fileName} baris sumber ${issue.sourceExcelRow}: kolom wajib ` +
        `${issue.column} FUNCLOC=${issue.tplnr}${msg}`
    );
}

/** @deprecated gunakan formatFieldViolation */
function formatRequiredViolation(issue) {
    return formatFieldViolation(issue);
}

module.exports = {
    DEFAULT_TEMPLATE_PATH,
    TEMPLATE_HEADER_ROW,
    TEMPLATE_MAX_LEN_ROW,
    MAINTENANCE_PLAN_OUTPUT_START_ROW,
    REQUIRED_FIELD_COLUMNS,
    DEFAULT_STRICT_FIELD_COLUMNS,
    loadMaintenancePlanTemplateLimits,
    validateOutputRow,
    validateMaintenancePlanFields,
    formatCharViolation,
    formatFieldViolation,
    formatRequiredViolation
};
