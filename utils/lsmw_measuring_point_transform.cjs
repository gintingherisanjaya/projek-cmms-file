/**
 * Transformasi baris Protect Y → baris output LSMW measuring point.
 */

const {
    buildColumnIndexFirstWins,
    shouldSkipLevel5Stas13Row
} = require("./lsmw_lookups.cjs");
const { findDataLayout, normalizeHeader } = require("./lsmw_cell_fill.cjs");
const { findFunclocDescAfterColumnIndex } = require("./equipment_gathering_columns.cjs");
const { validateOutputRow } = require("./lsmw_measuring_point_limits.cjs");
const { makeKey, lookupFunctionalLoc } = require("./equipment_number_mapping.cjs");
const {
    createValidationStats,
    recordValidationItem,
    validationStatsToResult
} = require("./lsmw_pks_validation_stats.cjs");
const {
    fixFunclocDescSourceTypos,
    applyFunclocDescAlias
} = require("./funcloc_desc_alias.cjs");

const OUTPUT_COLUMN_ORDER = [
    "MPOTY",
    "EQUNR",
    "MPTYP",
    "PTTXT",
    "PSORT",
    "ATNAM",
    "BEGRU",
    "CJUMC",
    "PYEAC"
];

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

/** Segmen ke-2 FUNCTIONAL LOCATION AFTER (split "-"). */
function getFuncLocSecondSegment(raw) {
    const parts = String(raw ?? "")
        .split("-")
        .map(p => p.trim())
        .filter(Boolean);
    return parts.length >= 2 ? parts[1] : "";
}

function funcLocLevel(code) {
    if (!code || typeof code !== "string") return 0;
    return code.trim().split("-").filter(s => s.length > 0).length;
}

function isExcludedDescription(desc) {
    const d = String(desc ?? "");
    if (/drive\s*unit/i.test(d)) return true;
    if (/accessories\s+weigh\s*bridge/i.test(d)) return true;
    return false;
}

function isMeasuringPointSourceRow(funcLoc, desc) {
    if (funcLocLevel(funcLoc) < 5) return false;
    if (isExcludedDescription(desc)) return false;
    return true;
}

function buildEqunr(segment, desc) {
    const seg = cellText(segment);
    const d = cellText(desc);
    if (!seg) return d;
    if (!d) return seg;
    return `${seg}-${d}`;
}

function buildPttxt(desc) {
    return "Pencatatan HM";
}

function descForEquipmentLookup(rawDesc) {
    return applyFunclocDescAlias(fixFunclocDescSourceTypos(rawDesc));
}

function lookupEquipmentNumber(segment, desc, equipmentByPlantDesc) {
    const key = makeKey(segment, descForEquipmentLookup(desc));
    return equipmentByPlantDesc?.get(key) ?? "";
}

function buildOutputValues(funcLoc, desc, equipmentByPlantDesc) {
    const segment = getFuncLocSecondSegment(funcLoc);
    const mappedEqunr = lookupEquipmentNumber(segment, desc, equipmentByPlantDesc);
    return {
        values: {
            MPOTY: "IEQ",
            EQUNR: mappedEqunr || buildEqunr(segment, desc),
            MPTYP: "M",
            PTTXT: buildPttxt(desc),
            PSORT: "HMKM_FRONTEND",
            ATNAM: "ZPM_HOURMETER",
            BEGRU: segment,
            CJUMC: "999999",
            PYEAC: "8400"
        },
        equnrFromMapping: Boolean(mappedEqunr)
    };
}

/**
 * @param {Array<Array<unknown>>} rawRows
 * @param {string} fileName
 * @param {Record<string, number|null>} limitByTemplate
 * @param {{
 *   redEquipGroupRowIndices?: Set<number>,
 *   seenFuncLocKeys?: Set<string>,
 *   workCenterByRowIndex?: Map<number, unknown>,
 *   equipmentByPlantDesc?: Map<string, string>,
 *   funcLocByPlantDesc?: Map<string, string>
 * }} [options]
 * @returns {{
 *   ok: boolean,
 *   reason?: string,
 *   rows?: Array<Record<string, unknown>>,
 *   violations?: Array<object>,
 *   equnrFoundCount?: number,
 *   equnrNotFoundCount?: number,
 *   plant?: string,
 *   validFuncLocCount?: number,
 *   invalidFuncLocCount?: number,
 *   invalidFuncLocList?: string
 * }}
 */
function buildMeasuringPointRows(rawRows, fileName, limitByTemplate, options = {}) {
    const layout = findDataLayout(rawRows);
    const headerRow = rawRows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header tidak ditemukan" };
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxDesc = findFunclocDescAfterColumnIndex(headerRow);

    if (idxFuncLoc === undefined) {
        return {
            ok: false,
            reason: "Kolom FUNCTIONAL LOCATION AFTER tidak ditemukan"
        };
    }

    if (idxDesc === undefined) {
        return {
            ok: false,
            reason: "Kolom FUNCTLOC DESC. AFTER tidak ditemukan"
        };
    }

    const dataRows = rawRows.slice(layout.headerRowIndex + 1);
    const outputRows = [];
    const violations = [];
    const validationStats = createValidationStats();

    for (let i = 0; i < dataRows.length; i++) {
        if (options.redEquipGroupRowIndices?.has(i)) continue;

        const r = dataRows[i] || [];
        const funcLoc = cellText(r[idxFuncLoc]);
        if (!funcLoc) continue;

        const desc = cellText(r[idxDesc]);
        if (!isMeasuringPointSourceRow(funcLoc, desc)) continue;
        const funcLocKey = normalizeFuncLocKey(funcLoc);
        if (funcLocKey && options.seenFuncLocKeys?.has(funcLocKey)) continue;

        const effectiveCostCenter =
            options.effectiveCostCenterByRowIndex?.get(i);
        if (shouldSkipLevel5Stas13Row(funcLocKey, effectiveCostCenter)) continue;

        if (funcLocKey) options.seenFuncLocKeys?.add(funcLocKey);

        const built = buildOutputValues(
            funcLoc,
            desc,
            options.equipmentByPlantDesc
        );
        const sourceExcelRow = layout.headerRowIndex + 2 + i;
        const segment = getFuncLocSecondSegment(funcLoc);
        const lookupDesc = descForEquipmentLookup(desc);
        const mappingFuncLoc = lookupFunctionalLoc(
            segment,
            lookupDesc,
            options.funcLocByPlantDesc
        );

        recordValidationItem(validationStats, {
            funcLoc,
            plantSegment: segment,
            equnrFromMapping: built.equnrFromMapping,
            mappingFuncLoc
        });

        const rowViolations = validateOutputRow(built.values, limitByTemplate, {
            fileName,
            sourceExcelRow
        });
        violations.push(...rowViolations);
        outputRows.push(built.values);
    }

    return {
        ok: true,
        rows: outputRows,
        violations,
        ...validationStatsToResult(validationStats)
    };
}

module.exports = {
    OUTPUT_COLUMN_ORDER,
    getFuncLocSecondSegment,
    buildEqunr,
    lookupEquipmentNumber,
    buildPttxt,
    buildOutputValues,
    buildMeasuringPointRows
};
