/**
 * Transformasi baris Protect Y → baris output LSMW maintenance item.
 */

const {
    buildColumnIndexFirstWins,
    shouldSkipLevel5Stas13Row,
    isEquipmentSourceRow
} = require("./lsmw_lookups.cjs");
const { findDataLayout, normalizeHeader } = require("./lsmw_cell_fill.cjs");
const { findFunclocDescAfterColumnIndex } = require("./equipment_gathering_columns.cjs");
const {
    loadPreventiveMaintenanceRules,
    matchFilterKeyword,
    pickJsonEntryForRow,
    buildTaskListGroup,
    getFuncLocSecondSegment
} = require("./lsmw_tasklist_transform.cjs");
const { makeKey, lookupFunctionalLoc } = require("./equipment_number_mapping.cjs");
const { validateOutputRow } = require("./lsmw_maintenance_item_limits.cjs");
const {
    createValidationStats,
    recordValidationItem,
    validationStatsToResult
} = require("./lsmw_pks_validation_stats.cjs");
const {
    fixFunclocDescSourceTypos,
    applyFunclocDescAlias
} = require("./funcloc_desc_alias.cjs");

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

function descForEquipmentLookup(rawDesc) {
    return applyFunclocDescAlias(fixFunclocDescSourceTypos(rawDesc));
}

function buildMaintenanceItemFallbackEqunr(plant, desc) {
    const seg = cellText(plant);
    const d = cellText(desc);
    if (!seg) return d;
    if (!d) return seg;
    return `${seg}-${d}`;
}

function lookupEquipmentNumber(plant, desc, equipmentByPlantDesc) {
    const key = makeKey(plant, descForEquipmentLookup(desc));
    return equipmentByPlantDesc?.get(key) ?? "";
}

/**
 * @param {{
 *   plant: string,
 *   funcLoc: string,
 *   desc: string,
 *   jsonEntry: object,
 *   workCenter: unknown,
 *   plannerGroup: unknown,
 *   equipmentByPlantDesc?: Map<string, string>
 * }} params
 * @returns {Record<string, unknown>}
 */
function buildOutputRow(params) {
    const {
        plant,
        funcLoc,
        desc,
        jsonEntry,
        workCenter,
        plannerGroup,
        equipmentByPlantDesc
    } = params;
    const mappedEqunr = lookupEquipmentNumber(plant, desc, equipmentByPlantDesc);

    return {
        values: {
            MPTYP: "PM",
            WSTRA: "PALMPB",
            PSTXT: `${cellText(plant)} ${cellText(jsonEntry.text)}`.trim(),
            EQUNR: mappedEqunr || buildMaintenanceItemFallbackEqunr(plant, desc),
            IWERK: cellText(plant),
            WPGRP: plannerGroup ?? null,
            AUART: "PM02",
            ILART: "PRE",
            GEWERK: workCenter ?? null,
            WERGW: cellText(plant),
            PLNTY: "A",
            PLNNR: buildTaskListGroup(plant, jsonEntry.task_list_group),
            PLNAL: "1",
            PRIOK: "3"
        },
        equnrFromMapping: Boolean(mappedEqunr)
    };
}

/**
 * @param {Array<Array<unknown>>} rawRows
 * @param {{
 *   redEquipGroupRowIndices?: Set<number>,
 *   seenFuncLocKeys?: Set<string>,
 *   workCenterByRowIndex?: Map<number, unknown>,
 *   plannerGroupByRowIndex?: Map<number, unknown>,
 *   plannerGroupByRowIndex?: Map<number, unknown>,
 *   equipmentByPlantDesc?: Map<string, string>,
 *   funcLocByPlantDesc?: Map<string, string>
 * }} options
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
function buildMaintenanceItemRows(
    rawRows,
    fileName,
    limitByTemplate,
    options = {}
) {
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

    const { byFilter, filterKeywords } = loadPreventiveMaintenanceRules();
    const dataRows = rawRows.slice(layout.headerRowIndex + 1);

    /** @type {Array<Record<string, unknown>>} */
    const rows = [];
    const violations = [];
    const validationStats = createValidationStats();

    for (let i = 0; i < dataRows.length; i++) {
        if (options.redEquipGroupRowIndices?.has(i)) continue;

        const r = dataRows[i] || [];
        const funcLoc = cellText(r[idxFuncLoc]);
        if (!funcLoc) continue;

        const desc = cellText(r[idxDesc]);
        if (!isEquipmentSourceRow(funcLoc, desc)) continue;

        const funcLocKey = normalizeFuncLocKey(funcLoc);
        if (funcLocKey && options.seenFuncLocKeys?.has(funcLocKey)) continue;

        const effectiveCostCenter =
            options.effectiveCostCenterByRowIndex?.get(i);
        if (shouldSkipLevel5Stas13Row(funcLocKey, effectiveCostCenter)) continue;

        const workCenter = options.workCenterByRowIndex?.get(i);

        const filterKey = matchFilterKeyword(desc, filterKeywords);
        if (!filterKey) continue;

        const jsonEntry = pickJsonEntryForRow(
            byFilter.get(filterKey),
            desc
        );
        if (!jsonEntry) continue;

        if (funcLocKey) options.seenFuncLocKeys?.add(funcLocKey);

        const plant = getFuncLocSecondSegment(funcLoc);
        const plannerGroup = options.plannerGroupByRowIndex?.get(i) ?? null;

        const built = buildOutputRow({
            plant,
            funcLoc,
            desc,
            jsonEntry,
            workCenter,
            plannerGroup,
            equipmentByPlantDesc: options.equipmentByPlantDesc
        });

        const lookupDesc = descForEquipmentLookup(desc);
        const mappingFuncLoc = lookupFunctionalLoc(
            plant,
            lookupDesc,
            options.funcLocByPlantDesc
        );

        recordValidationItem(validationStats, {
            funcLoc,
            plantSegment: plant,
            equnrFromMapping: built.equnrFromMapping,
            mappingFuncLoc
        });

        const sourceExcelRow = layout.headerRowIndex + 2 + i;
        violations.push(
            ...validateOutputRow(built.values, limitByTemplate, {
                fileName,
                sourceExcelRow
            })
        );

        rows.push(built.values);
    }

    return {
        ok: true,
        rows,
        violations,
        ...validationStatsToResult(validationStats)
    };
}

module.exports = {
    buildMaintenanceItemRows,
    buildOutputRow,
    lookupEquipmentNumber,
    buildMaintenanceItemFallbackEqunr
};
