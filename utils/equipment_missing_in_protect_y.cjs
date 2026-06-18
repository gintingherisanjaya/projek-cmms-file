/**
 * Loader equipment-missing-in-protect-y.xlsx — sumber before untuk equipment
 * yang tidak lengkap di Protect Y (match Plant + Target Func Loc = FUNC LOC AFTER).
 */

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");

const DEFAULT_MISSING_EQUIPMENT_PATH = "./equipment-missing-in-protect-y.xlsx";

const REQUIRED_HEADERS = [
    "PLANT",
    "EQUIPMENT",
    "DESCRIPTION",
    "FUNCTIONAL LOC.",
    "TARGET FUNC LOC"
];

function normalizeHeader(name) {
    if (!name) return "";
    return String(name).trim().toUpperCase();
}

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function normalizePlant(raw) {
    const s = String(raw ?? "")
        .replace(/[`]/g, "")
        .replace(/\s+/g, "")
        .trim()
        .toUpperCase();
    if (!s) return "";
    return s.length >= 4 ? s.slice(0, 4) : s;
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

function normalizeEquipmentNumber(v) {
    if (v === null || v === undefined || v === "") return "";
    return String(v).trim();
}

function normalizePercentage(raw) {
    if (raw === null || raw === undefined || raw === "") return "";
    if (typeof raw === "number" && Number.isFinite(raw)) return raw;
    const s = String(raw).trim();
    if (!s) return "";
    const n = Number(s);
    return Number.isFinite(n) ? n : s;
}

function makeLookupKey(plant, targetFuncLoc) {
    return `${normalizePlant(plant)}\0${normalizeFuncLocKey(targetFuncLoc)}`;
}

function isValidPlantCode(code) {
    const s = String(code ?? "").trim().toUpperCase();
    return /^[A-Z0-9]{4}$/.test(s);
}

/**
 * Plant dari kelompok ke-2 FUNCTIONAL LOCATION AFTER (split "-").
 * Contoh: PALM-4F05-0005-... → 4F05
 * @param {unknown} funcLocAfter
 * @returns {string}
 */
function plantFromFuncLocAfter(funcLocAfter) {
    const raw = normalizeFuncLocKey(funcLocAfter);
    if (!raw) return "";

    const segs = raw.split("-").filter(Boolean);
    if (segs.length < 2) return "";

    const plant = segs[1].trim().toUpperCase();
    if (!isValidPlantCode(plant)) return "";

    return plant;
}

/**
 * @param {string} [filePath]
 * @returns {{ rows: Array<object>, byPlantAndTarget: Map<string, Array<object>>, filePath: string }}
 */
function loadEquipmentMissingInProtectY(
    filePath = DEFAULT_MISSING_EQUIPMENT_PATH
) {
    const resolved = path.resolve(filePath);
    if (!fs.existsSync(resolved)) {
        throw new Error(`File tidak ditemukan: ${resolved}`);
    }

    const workbook = XLSX.readFile(resolved);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    if (!rawRows.length) {
        throw new Error(`Sheet kosong: ${resolved}`);
    }

    const headerRow = rawRows[0];
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);

    for (const h of REQUIRED_HEADERS) {
        if (colIndex[h] === undefined) {
            throw new Error(
                `Kolom "${h}" tidak ditemukan di ${path.basename(resolved)}`
            );
        }
    }

    const idxPlant = colIndex.PLANT;
    const idxEquipment = colIndex.EQUIPMENT;
    const idxDescription = colIndex.DESCRIPTION;
    const idxFunctionalLoc = colIndex["FUNCTIONAL LOC."];
    const idxDescFuncLocOld =
        colIndex["DESCRIPTION FUNC LOC (OLD)"] ??
        colIndex["DESCRIPTION FUNC LOC(OLD)"];
    const idxTargetFuncLoc = colIndex["TARGET FUNC LOC"];
    const idxPercentage = colIndex.PERCENTAGE;

    const rows = [];
    const byPlantAndTarget = new Map();

    for (let i = 1; i < rawRows.length; i += 1) {
        const r = rawRows[i];
        const plant = normalizePlant(r[idxPlant]);
        const equipment = normalizeEquipmentNumber(r[idxEquipment]);
        const targetFuncLoc = normalizeFuncLocKey(r[idxTargetFuncLoc]);
        const functionalLoc = normalizeFuncLocKey(r[idxFunctionalLoc]);

        if (!plant || !targetFuncLoc) continue;
        if (!equipment && !functionalLoc) continue;

        const entry = {
            plant,
            equipment,
            description: cellText(r[idxDescription]),
            functionalLoc,
            descriptionFuncLocOld:
                idxDescFuncLocOld !== undefined
                    ? cellText(r[idxDescFuncLocOld])
                    : "",
            targetFuncLoc,
            percentage:
                idxPercentage !== undefined
                    ? normalizePercentage(r[idxPercentage])
                    : ""
        };

        rows.push(entry);

        const key = makeLookupKey(plant, targetFuncLoc);
        if (!byPlantAndTarget.has(key)) {
            byPlantAndTarget.set(key, []);
        }
        byPlantAndTarget.get(key).push(entry);
    }

    return { rows, byPlantAndTarget, filePath: resolved };
}

/**
 * @param {Map<string, Array<object>>} byPlantAndTarget
 * @param {string} plant
 * @param {string} targetFuncLoc
 * @returns {Array<object>}
 */
function lookupMissingByPlantAndTarget(byPlantAndTarget, plant, targetFuncLoc) {
    if (!byPlantAndTarget || !plant || !targetFuncLoc) return [];
    const key = makeLookupKey(plant, targetFuncLoc);
    return byPlantAndTarget.get(key) ?? [];
}

module.exports = {
    DEFAULT_MISSING_EQUIPMENT_PATH,
    normalizePlant,
    normalizeFuncLocKey,
    normalizeEquipmentNumber,
    isValidPlantCode,
    plantFromFuncLocAfter,
    loadEquipmentMissingInProtectY,
    lookupMissingByPlantAndTarget
};
