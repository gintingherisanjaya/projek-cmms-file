/**
 * Validasi kolom wajib output LSMW measuring document.
 */

const { formatHeriCounterOutput } = require("./heri_sheet_loader.cjs");

const VALIDATION_FIELD_COLUMNS = [
    "Counter Reading",
    "Equipment Number",
    "Measuring point"
];

const DEFAULT_STRICT_FIELD_COLUMNS = [...VALIDATION_FIELD_COLUMNS];

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function formatCounterDisplay(counterValue) {
    if (counterValue === null || counterValue === undefined || counterValue === "") {
        return "";
    }
    return formatHeriCounterOutput(counterValue) ?? "";
}

/**
 * @param {{
 *   counterIssue?: string | null,
 *   equnrFromMapping: boolean,
 *   equnr: string,
 *   point: string,
 *   counterValue?: number | null
 * }} params
 * @param {{
 *   regional?: string,
 *   fileName: string,
 *   sourceExcelRow: number,
 *   equipmentDescription: string,
 *   plant?: string,
 *   funcLoc?: string
 * }} meta
 */
function validateMeasuringDocumentFields(params, meta) {
    const { counterIssue, equnrFromMapping, equnr, point, counterValue } = params;
    const issues = [];

    const snapshot = {
        regional: meta.regional ?? "",
        fileName: meta.fileName,
        plant: cellText(meta.plant),
        sourceExcelRow: meta.sourceExcelRow,
        equipmentDescription: cellText(meta.equipmentDescription),
        funcLoc: cellText(meta.funcLoc),
        counterReading: formatCounterDisplay(counterValue),
        equipmentNumber: cellText(equnr),
        measuringPoint: cellText(point)
    };

    const base = { ...snapshot };

    if (counterIssue) {
        issues.push({
            ...base,
            column: "Counter Reading",
            message: counterIssue
        });
    } else if (counterValue === null || counterValue === undefined) {
        issues.push({
            ...base,
            column: "Counter Reading",
            message: "counter HERI kosong"
        });
    }

    if (!equnrFromMapping) {
        issues.push({
            ...base,
            column: "Equipment Number",
            message: "tidak ditemukan di new-equipment-number.xlsx"
        });
    }

    if (!cellText(point)) {
        issues.push({
            ...base,
            column: "Measuring point",
            message: cellText(equnr)
                ? `tidak ditemukan di measuring-item-ik07.xlsx untuk EQUNR ${cellText(equnr)}`
                : "Measuring point kosong (EQUNR tidak tersedia)"
        });
    }

    return issues;
}

function formatFieldViolation(issue) {
    const msg = issue.message ? ` — ${issue.message}` : "";
    const loc = issue.funcLoc ? ` FUNCLOC=${issue.funcLoc}` : "";
    return (
        `${issue.fileName} baris sumber ${issue.sourceExcelRow}: kolom wajib ` +
        `${issue.column}${loc}${msg}`
    );
}

module.exports = {
    VALIDATION_FIELD_COLUMNS,
    DEFAULT_STRICT_FIELD_COLUMNS,
    validateMeasuringDocumentFields,
    formatFieldViolation
};
