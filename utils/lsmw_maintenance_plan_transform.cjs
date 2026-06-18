/**
 * Transformasi baris Protect Y → baris output LSMW maintenance plan.
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
    getFuncLocSecondSegment
} = require("./lsmw_tasklist_transform.cjs");
const {
    lookupEquipmentNumber,
    buildMaintenanceItemFallbackEqunr
} = require("./lsmw_maintenance_item_transform.cjs");
const { lookupFirstMaintItem } = require("./maintenance_item_ip18_mapping.cjs");
const { lookupMapByEquipmentKeys } = require("./equipment_key_lookup.cjs");
const {
    resolveHeriCounterForMaintenancePlan,
    formatHeriCounterOutput
} = require("./heri_sheet_loader.cjs");
const { validateOutputRow } = require("./lsmw_measuring_point_limits.cjs");
const { validateMaintenancePlanFields } = require("./lsmw_maintenance_plan_limits.cjs");

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

/**
 * @param {{
 *   plant: string,
 *   funcLoc: string,
 *   desc: string,
 *   jsonEntry: object,
 *   equipmentByPlantDesc?: Map<string, string>,
 *   maintByEquipment?: Map<string, string[]>,
 *   pointByEquipment?: Map<string, string>,
 *   counterValue?: number
 * }} params
 */
function buildMaintenancePlanOutputRow(params) {
    const {
        plant,
        desc,
        jsonEntry,
        equipmentByPlantDesc,
        maintByEquipment,
        pointByEquipment,
        counterValue = 0
    } = params;

    const mappedEqunr = lookupEquipmentNumber(plant, desc, equipmentByPlantDesc);
    const fallbackEqunrKey = buildMaintenanceItemFallbackEqunr(plant, desc);
    const equnr = mappedEqunr || fallbackEqunrKey;

    const maintLookup = lookupMapByEquipmentKeys(
        mappedEqunr,
        fallbackEqunrKey,
        maintByEquipment,
        { isEmpty: arr => !arr?.length }
    );
    const { maintItem, extraCount } = lookupFirstMaintItem(
        /** @type {string[]} */ (maintLookup.hit ?? [])
    );

    const pointLookup = lookupMapByEquipmentKeys(
        mappedEqunr,
        fallbackEqunrKey,
        pointByEquipment,
        { isEmpty: v => !cellText(v) }
    );
    const point = cellText(pointLookup.hit);

    const values = {
        MPTYP: "PM",
        WSTRA: "PALMPB",
        MAINTENANCE_ITEM: maintItem || null,
        WPTXT: `${cellText(plant)} ${cellText(jsonEntry.text)}`.trim(),
        POINT: point || null,
        HORIZ: "90",
        HORIZ_QUALIFIER: "%",
        ABRHO: "5",
        HUNIT: "YR",
        SZAEH: formatHeriCounterOutput(counterValue)
    };

    return {
        values,
        equnr,
        equnrFromMapping: Boolean(mappedEqunr),
        maintItemFromIp18: maintLookup.fromPrimary && Boolean(maintItem),
        pointFromIk07: pointLookup.fromPrimary && Boolean(point),
        maintItemExtra: extraCount > 0
    };
}

/**
 * @param {Array<Array<unknown>>} rawRows
 * @param {string} fileName
 * @param {Record<string, number|null>} limitByTemplate
 * @param {{
 *   regional?: string,
 *   redEquipGroupRowIndices?: Set<number>,
 *   seenFuncLocKeys?: Set<string>,
 *   workCenterByRowIndex?: Map<number, unknown>,
 *   plannerGroupByRowIndex?: Map<number, unknown>,
 *   effectiveCostCenterByRowIndex?: Map<number, string>,
 *   equipmentByPlantDesc?: Map<string, string>,
 *   maintByEquipment?: Map<string, string[]>,
 *   pointByEquipment?: Map<string, string>,
 *   counterByEqunr?: Map<string, string>,
 *   heriData?: object
 * }} options
 */
function buildMaintenancePlanRows(rawRows, fileName, limitByTemplate, options = {}) {
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
    const fieldIssues = [];
    const maintItemExtras = [];

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

        const filterKey = matchFilterKeyword(desc, filterKeywords);
        if (!filterKey) continue;

        const jsonEntry = pickJsonEntryForRow(
            byFilter.get(filterKey),
            desc
        );
        if (!jsonEntry) continue;

        if (funcLocKey) options.seenFuncLocKeys?.add(funcLocKey);

        const plant = getFuncLocSecondSegment(funcLoc);

        const heriResult = resolveHeriCounterForMaintenancePlan(
            options.heriData,
            plant,
            desc,
            { allowSterilizerAlias: true }
        );
        const counterValue = heriResult.counterValue ?? 0;

        const built = buildMaintenancePlanOutputRow({
            plant,
            funcLoc,
            desc,
            jsonEntry,
            equipmentByPlantDesc: options.equipmentByPlantDesc,
            maintByEquipment: options.maintByEquipment,
            pointByEquipment: options.pointByEquipment,
            counterValue
        });

        if (built.maintItemExtra) {
            maintItemExtras.push({
                fileName,
                tplnr: funcLoc,
                sourceExcelRow: layout.headerRowIndex + 2 + i
            });
        }

        const meta = {
            fileName,
            sourceExcelRow: layout.headerRowIndex + 2 + i,
            tplnr: funcLoc,
            regional: options.regional ?? ""
        };

        const rowViolations = validateOutputRow(
            built.values,
            limitByTemplate,
            meta
        ).map(v => ({ ...v, tplnr: funcLoc }));
        violations.push(...rowViolations);

        fieldIssues.push(
            ...validateMaintenancePlanFields(
                {
                    equnr: built.equnr,
                    equnrFromMapping: built.equnrFromMapping,
                    maintItemFromIp18: built.maintItemFromIp18,
                    pointFromIk07: built.pointFromIk07,
                    counterByEqunr: options.counterByEqunr,
                    values: built.values
                },
                meta
            )
        );

        rows.push(built.values);
    }

    return {
        ok: true,
        rows,
        violations,
        fieldIssues,
        maintItemExtras,
        warnings: []
    };
}

module.exports = {
    buildMaintenancePlanRows,
    buildMaintenancePlanOutputRow
};
