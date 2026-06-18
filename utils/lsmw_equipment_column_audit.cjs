/**
 * Audit kolom header Protect Y untuk lsmw-equipment (column.json + validation.xlsx).
 */

const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");
const { normalizeHeader } = require("./lsmw_cell_fill.cjs");
const {
    findRegionalColumnIndex,
    findMaintenancePlantColumnIndex,
    findCompanyCodeColumnIndex,
    findMeasurmntRangeUnitColumnIndex,
    findFunctionalLocationBeforeColumnIndex,
    findFunclocDescBeforeColumnIndex,
    findFunclocDescAfterLevel123ColumnIndex,
    isFunctionalLocationAfterHeader
} = require("./equipment_gathering_columns.cjs");

const PROTECT_Y_CANONICAL_COLUMNS = [
    "REGIONAL",
    "POM",
    "EQUIPMENT NUMBER",
    "MAINTENANCE PLANT",
    "COMPANY CODE",
    "COST CENTER BEFORE",
    "EQKTU BEFORE",
    "OBJECT NUMBER",
    "EQUIPMENT CATEGORY",
    "EQUIPMENT GROUP BEFORE",
    "CONSTRUCTION YEAR",
    "GROES",
    "MANUFACTURE",
    "MEASURING POINT",
    "MeasurmntRangeUnit",
    "MAINTENANCE PLANNER GROUP",
    "WORK CENTER",
    "FUNCTIONAL LOCATION BEFORE",
    "FUNCTLOC DESC. BEFORE",
    "COST CENTER AFTER",
    "FUNCTIONAL LOCATION AFTER",
    "FUNCTLOC DESC. AFTER LEVEL 1,2,3",
    "EQKTU AFTER LEVEL 2 dan 3",
    "EQUIPMENT GROUP AFTER",
    "POSITION",
    "MERK",
    "TYPE",
    "CAPACITY",
    "STANDAR UMUR TEKNIS (TAHUN)",
    "UMUR PEMAKAIAN ALAT-% Kondisi Mesin/Peralatan (AKTUAL)",
    "NILAI EKONOMIS AKTUAL Nilai Aktiva Terakhir (Rp)",
    "WORK CENTER AFTER",
    "PLANNER GROUP AFTER",
    "DESCRIPTION",
    "MAINTENANCE PLAN AFTER",
    "ABC INDICATORS"
];

const VALIDATION_FILE_HEADER = "nama file";

function headerText(headerRow, index) {
    if (index === undefined || index < 0) return "";
    const v = headerRow[index];
    if (v === null || v === undefined) return "";
    return String(v);
}

function findCostCenterBeforeColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = normalizeHeader(headerRow[i]);
        if (h.includes("COST CENTER") && h.includes("BEFORE")) return i;
    }
    return undefined;
}

function findCostCenterAfterColumnIndexExact(headerRow) {
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = normalizeHeader(headerRow[i]);
        if (h.includes("COST CENTER") && h.includes("AFTER")) return i;
    }
    return undefined;
}

function findEquipmentGroupAfterNativeColumnIndex(headerRow, colIndex) {
    if (colIndex["EQUIPMENT GROUP AFTER"] !== undefined) {
        return colIndex["EQUIPMENT GROUP AFTER"];
    }
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = normalizeHeader(headerRow[i]);
        if (
            h.includes("EQUIPMENT GROUP") &&
            h.includes("AFTER") &&
            !h.includes("BEFORE")
        ) {
            return i;
        }
    }
    return undefined;
}

function findEquipmentGroupPlainAfterCostCenter(headerRow) {
    const idxCostCenterAfter = findCostCenterAfterColumnIndexExact(headerRow);
    if (idxCostCenterAfter === undefined) return undefined;

    for (let i = idxCostCenterAfter + 1; i < headerRow.length; i += 1) {
        if (normalizeHeader(headerRow[i]) === "EQUIPMENT GROUP") return i;
    }
    return undefined;
}

/**
 * @param {Array<unknown>} headerRow
 * @param {Record<string, number>} colIndex
 * @returns {{
 *   columnIndex: number|undefined,
 *   resolvedHeader: string,
 *   source: "equipment_group_after"|"equipment_group_plain"|undefined
 * }}
 */
function resolveEquipmentGroupAfterColumn(headerRow, colIndex) {
    const nativeIndex = findEquipmentGroupAfterNativeColumnIndex(
        headerRow,
        colIndex
    );
    if (nativeIndex !== undefined) {
        return {
            columnIndex: nativeIndex,
            resolvedHeader: headerText(headerRow, nativeIndex),
            source: "equipment_group_after"
        };
    }

    const plainIndex = findEquipmentGroupPlainAfterCostCenter(headerRow);
    if (plainIndex !== undefined) {
        return {
            columnIndex: plainIndex,
            resolvedHeader: headerText(headerRow, plainIndex),
            source: "equipment_group_plain"
        };
    }

    return {
        columnIndex: undefined,
        resolvedHeader: "",
        source: undefined
    };
}

function findEqktuAfterLevelColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = normalizeHeader(headerRow[i]);
        if (h.includes("EQKTU") && h.includes("AFTER") && h.includes("LEVEL")) {
            return i;
        }
    }
    return undefined;
}

function findStandarUmurTeknisColumnIndex(headerRow, colIndex) {
    const keys = [
        "STANDAR UMUR TEKNIS (TAHUN)",
        "STANDART UMUR TEKNIS (TAHUN)"
    ];
    for (const key of keys) {
        if (colIndex[key] !== undefined) return colIndex[key];
    }
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = normalizeHeader(headerRow[i]);
        if (
            (h.includes("STANDAR") || h.includes("STANDART")) &&
            h.includes("UMUR TEKNIS")
        ) {
            return i;
        }
    }
    return undefined;
}

function findUmurPemakaianColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = normalizeHeader(headerRow[i]);
        if (h.includes("UMUR PEMAKAIAN")) return i;
    }
    return undefined;
}

function findNilaiEkonomisColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = normalizeHeader(headerRow[i]);
        if (h.includes("NILAI EKONOMIS")) return i;
    }
    return undefined;
}

function findFunctionalLocationAfterColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i += 1) {
        if (isFunctionalLocationAfterHeader(headerRow[i])) return i;
    }
    return undefined;
}

/**
 * @param {string} canonical
 * @param {Array<unknown>} headerRow
 * @param {Record<string, number>} colIndex
 * @returns {number|undefined}
 */
function resolveColumnIndex(canonical, headerRow, colIndex) {
    switch (canonical) {
        case "REGIONAL":
            return findRegionalColumnIndex(headerRow);
        case "MAINTENANCE PLANT":
            return findMaintenancePlantColumnIndex(headerRow);
        case "COMPANY CODE":
            return findCompanyCodeColumnIndex(headerRow);
        case "MeasurmntRangeUnit":
            return findMeasurmntRangeUnitColumnIndex(headerRow);
        case "FUNCTIONAL LOCATION BEFORE":
            return findFunctionalLocationBeforeColumnIndex(headerRow);
        case "FUNCTLOC DESC. BEFORE":
            return findFunclocDescBeforeColumnIndex(headerRow);
        case "FUNCTLOC DESC. AFTER LEVEL 1,2,3":
            return findFunclocDescAfterLevel123ColumnIndex(headerRow);
        case "FUNCTIONAL LOCATION AFTER":
            return findFunctionalLocationAfterColumnIndex(headerRow);
        case "COST CENTER BEFORE":
            return findCostCenterBeforeColumnIndex(headerRow);
        case "COST CENTER AFTER":
            return findCostCenterAfterColumnIndexExact(headerRow);
        case "EQUIPMENT GROUP AFTER":
            return resolveEquipmentGroupAfterColumn(headerRow, colIndex)
                .columnIndex;
        case "EQKTU AFTER LEVEL 2 dan 3":
            return findEqktuAfterLevelColumnIndex(headerRow);
        case "STANDAR UMUR TEKNIS (TAHUN)":
            return findStandarUmurTeknisColumnIndex(headerRow, colIndex);
        case "UMUR PEMAKAIAN ALAT-% Kondisi Mesin/Peralatan (AKTUAL)":
            return findUmurPemakaianColumnIndex(headerRow);
        case "NILAI EKONOMIS AKTUAL Nilai Aktiva Terakhir (Rp)":
            return findNilaiEkonomisColumnIndex(headerRow);
        default: {
            const key = normalizeHeader(canonical);
            return colIndex[key];
        }
    }
}

function getValidationHeaders() {
    return [VALIDATION_FILE_HEADER, ...PROTECT_Y_CANONICAL_COLUMNS];
}

/**
 * @param {Array<unknown>} headerRow
 * @returns {{
 *   resolutions: Map<string, { resolvedHeader: string, columnIndex: number }>,
 *   missingColumn: string[],
 *   colIndex: Record<string, number>
 * }}
 */
function auditProtectYColumns(headerRow) {
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    /** @type {Map<string, { resolvedHeader: string, columnIndex: number }>} */
    const resolutions = new Map();
    /** @type {string[]} */
    const missingColumn = [];

    for (const canonical of PROTECT_Y_CANONICAL_COLUMNS) {
        if (canonical === "EQUIPMENT GROUP AFTER") {
            const equipGroup = resolveEquipmentGroupAfterColumn(
                headerRow,
                colIndex
            );
            if (equipGroup.columnIndex === undefined) {
                missingColumn.push(canonical);
                continue;
            }
            resolutions.set(canonical, {
                resolvedHeader: equipGroup.resolvedHeader,
                columnIndex: equipGroup.columnIndex,
                source: equipGroup.source
            });
            continue;
        }

        const columnIndex = resolveColumnIndex(canonical, headerRow, colIndex);
        if (columnIndex === undefined) {
            missingColumn.push(canonical);
            continue;
        }
        resolutions.set(canonical, {
            resolvedHeader: headerText(headerRow, columnIndex),
            columnIndex
        });
    }

    return { resolutions, missingColumn, colIndex };
}

/**
 * @param {Map<string, { resolvedHeader: string, columnIndex: number }>} resolutions
 * @returns {Map<string, string>}
 */
function resolutionsToValidationMap(resolutions) {
    /** @type {Map<string, string>} */
    const out = new Map();
    for (const canonical of PROTECT_Y_CANONICAL_COLUMNS) {
        out.set(canonical, resolutions.get(canonical)?.resolvedHeader ?? "");
    }
    return out;
}

module.exports = {
    PROTECT_Y_CANONICAL_COLUMNS,
    VALIDATION_FILE_HEADER,
    getValidationHeaders,
    auditProtectYColumns,
    resolutionsToValidationMap,
    resolveColumnIndex,
    resolveEquipmentGroupAfterColumn
};
