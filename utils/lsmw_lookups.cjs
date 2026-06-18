/**
 * Lookup bersama CWC & Planner Group untuk skrip LSMW.
 * Kolom regional CSV ditentukan dari maintenance plant per file (bukan 4 char pertama cost center).
 */

const path = require("path");

let getCWC;
let getPlannerGroup;
let getRegionalColumnForMaintenancePlant;
/** @type {((costCenter: unknown) => string) | null} */
let extractCostCenterSuffixForCsv = null;

function isEmpty(v) {
    return v === null || v === undefined || String(v).trim() === "";
}

async function ensureCwcPlannerLookups() {
    if (getCWC && extractCostCenterSuffixForCsv) {
        return { getCWC, getPlannerGroup, getRegionalColumnForMaintenancePlant };
    }
    const cwcMod = await import("./extractCWC.js");
    const pgMod = await import("../extractPlannerGroup.js");
    const regionalMod = await import("./regionalPlantMapping.js");
    getCWC = cwcMod.getCWC;
    getPlannerGroup = pgMod.getPlannerGroup;
    getRegionalColumnForMaintenancePlant =
        regionalMod.getRegionalColumnForMaintenancePlant;
    extractCostCenterSuffixForCsv = regionalMod.extractCostCenterSuffixForCsv;
    return { getCWC, getPlannerGroup, getRegionalColumnForMaintenancePlant };
}

function funcLocLevel(code) {
    if (!code || typeof code !== "string") return 0;
    return code.trim().split("-").filter(s => s.length > 0).length;
}

/**
 * Cost center station STAS13 (Effluent/Limbah) — suffix sama dengan baris cwc.csv.
 * Panggil ensureCwcPlannerLookups() terlebih dahulu.
 */
function isStas13CostCenter(costCenter) {
    if (!extractCostCenterSuffixForCsv) {
        throw new Error(
            "extractCostCenterSuffixForCsv belum dimuat — panggil ensureCwcPlannerLookups() dulu"
        );
    }
    return extractCostCenterSuffixForCsv(costCenter) === "STAS13";
}

/**
 * Level 5 + cost center efektif STAS13: bukan equipment (skip di semua skrip LSMW terkait).
 */
function shouldSkipLevel5Stas13Row(funcLocKey, effectiveCostCenter) {
    if (funcLocLevel(funcLocKey) !== 5) return false;
    return isStas13CostCenter(effectiveCostCenter);
}

function isExcludedEquipmentDescription(desc) {
    const d = String(desc ?? "");
    if (/drive\s*unit/i.test(d)) return true;
    if (/accessories\s+weigh\s*bridge/i.test(d)) return true;
    return false;
}

/** Dipakai lsmw-maintenance-item/plan — hanya funcloc level 5 (5 segmen). */
function isEquipmentSourceRow(funcLoc, desc) {
    if (funcLocLevel(funcLoc) !== 5) return false;
    if (isExcludedEquipmentDescription(desc)) return false;
    return true;
}

function getEffectiveCostCenter(r, funcLocKey, funcLocMap, idxCostCenter) {
    if (idxCostCenter === undefined) return null;

    let value = r[idxCostCenter];
    if (!isEmpty(value)) return value;

    let fl = funcLocKey;
    while (fl && fl.includes("-")) {
        fl = fl.substring(0, fl.lastIndexOf("-"));
        const parentRow = funcLocMap[fl];
        if (!parentRow) continue;
        const parentVal = parentRow[idxCostCenter];
        if (!isEmpty(parentVal)) return parentVal;
    }

    return null;
}

/** Log maintenance plant + kolom regional CSV untuk satu file output. */
async function logRegionalMappingForFile(maintenancePlant, sourceLabel) {
    await ensureCwcPlannerLookups();
    const cwcMod = await import("./extractCWC.js");
    const cwcPath = cwcMod.setCwcSourceForMaintenancePlant(maintenancePlant);
    const availableColumns = cwcMod.availableColumns ?? new Set();
    const { plantCode, regionColumn } = getRegionalColumnForMaintenancePlant(
        maintenancePlant,
        availableColumns
    );
    const cwcLabel =
        cwcPath && cwcPath.includes("cwc_special")
            ? `cwc_special/${path.basename(cwcPath)}`
            : "cwc.csv (default)";
    console.log(
        `[regional] ${sourceLabel}: maintenance plant=${plantCode ?? "(n/a)"}, ` +
            `kolom CSV regional="${regionColumn ?? "(n/a)"}", ` +
            `CWC file: ${cwcLabel} ` +
            `(CWC + planner-group; nilai sumber diganti mapping CSV bila ada)`
    );
    return { plantCode, regionColumn };
}

/**
 * Planner group dari planner-group.csv; maintenance plant menentukan kolom regional.
 * Mapping CSV menggantikan nilai sumber bila lookup berhasil.
 */
function resolvePlannerGroup(r, funcLocKey, effectiveCostCenter, valFn, maintenancePlant) {
    if (!getPlannerGroup) {
        const src = valFn(r, "PLANNER GROUP AFTER");
        return isEmpty(src) ? null : src;
    }
    const mapped = getPlannerGroup(funcLocKey, effectiveCostCenter, {
        maintenancePlant
    });
    if (!isEmpty(mapped)) return mapped;
    const src = valFn(r, "PLANNER GROUP AFTER");
    return isEmpty(src) ? null : src;
}

/**
 * Work center dari cwc.csv; maintenance plant menentukan kolom regional.
 * Mapping CSV menggantikan nilai sumber bila lookup berhasil.
 */
function resolveWorkCenter(r, funcLocKey, effectiveCostCenter, valFn, maintenancePlant) {
    if (!getCWC) {
        const src = valFn(r, "WORK CENTER AFTER");
        return isEmpty(src) ? null : src;
    }
    const mapped = getCWC(funcLocKey, effectiveCostCenter, { maintenancePlant });
    if (!isEmpty(mapped)) return mapped;
    const src = valFn(r, "WORK CENTER AFTER");
    return isEmpty(src) ? null : src;
}

/** Indeks kolom header: nama duplikat → kolom pertama (bukan terakhir). */
function buildColumnIndexFirstWins(headerRow, normalizeHeaderFn) {
    const colIndex = {};
    headerRow.forEach((h, i) => {
        const key = normalizeHeaderFn(h);
        if (key && colIndex[key] === undefined) colIndex[key] = i;
    });
    return colIndex;
}

module.exports = {
    isEmpty,
    ensureCwcPlannerLookups,
    getEffectiveCostCenter,
    logRegionalMappingForFile,
    buildColumnIndexFirstWins,
    resolvePlannerGroup,
    resolveWorkCenter,
    isStas13CostCenter,
    shouldSkipLevel5Stas13Row,
    isExcludedEquipmentDescription,
    isEquipmentSourceRow
};
