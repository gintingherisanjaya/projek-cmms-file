const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DEFAULT_MAPPING_PATH = path.join(
    __dirname,
    "..",
    "new-equipment-number.xlsx"
);

function normalizePlant(value) {
    return String(value ?? "").trim().toUpperCase();
}

function normalizeDescription(value) {
    return String(value ?? "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();
}

function makeKey(plant, description) {
    return `${normalizePlant(plant)}\x1f${normalizeDescription(description)}`;
}

function findHeaderIndex(headerRow, expectedName) {
    const target = String(expectedName ?? "").trim().toUpperCase();
    for (let i = 0; i < headerRow.length; i += 1) {
        if (String(headerRow[i] ?? "").trim().toUpperCase() === target) return i;
    }
    return -1;
}

function findFirstHeaderIndex(headerRow, expectedNames) {
    for (const name of expectedNames) {
        const idx = findHeaderIndex(headerRow, name);
        if (idx >= 0) return idx;
    }
    return -1;
}

function normalizeFuncLocForCompare(value) {
    return String(value ?? "").trim().toUpperCase();
}

function lookupFunctionalLoc(plant, desc, funcLocByPlantDesc) {
    return funcLocByPlantDesc?.get(makeKey(plant, desc)) ?? "";
}

/**
 * Load mapping Planning Plant + Description -> Equipment (first wins).
 * @param {string} [filePath]
 * @returns {{
 *   byPlantDesc: Map<string, string>,
 *   funcLocByPlantDesc: Map<string, string>,
 *   duplicateKeys: number,
 *   totalRows: number,
 *   hasFunctionalLocColumn: boolean
 * }}
 */
function loadEquipmentNumberMapping(filePath = DEFAULT_MAPPING_PATH) {
    if (!fs.existsSync(filePath)) {
        throw new Error(`File mapping tidak ditemukan: ${filePath}`);
    }

    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) {
        throw new Error(`Sheet kosong pada file mapping: ${filePath}`);
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rows.length < 1) {
        throw new Error(`Header mapping tidak ditemukan: ${filePath}`);
    }

    const headerRow = rows[0] || [];
    const idxPlant = findFirstHeaderIndex(headerRow, [
        "Planning Plant",
        "Maintenance Plant"
    ]);
    const idxDesc = findFirstHeaderIndex(headerRow, ["Description", "EQKTU"]);
    const idxEquip = findHeaderIndex(headerRow, "Equipment");
    const idxFuncLoc = findFirstHeaderIndex(headerRow, [
        "Functional Loc.",
        "DESCRIPTION_FUNC_LOCATION"
    ]);

    if (idxPlant < 0 || idxDesc < 0 || idxEquip < 0) {
        throw new Error(
            "Kolom wajib mapping tidak lengkap: Planning Plant / Maintenance Plant, Description / EQKTU, Equipment"
        );
    }

    const byPlantDesc = new Map();
    const funcLocByPlantDesc = new Map();
    let duplicateKeys = 0;
    let totalRows = 0;

    for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i] || [];
        const plant = normalizePlant(row[idxPlant]);
        const description = normalizeDescription(row[idxDesc]);
        const equipment = String(row[idxEquip] ?? "").trim();
        if (!plant || !description || !equipment) continue;

        totalRows += 1;
        const key = makeKey(plant, description);
        if (byPlantDesc.has(key)) {
            duplicateKeys += 1;
            continue;
        }
        byPlantDesc.set(key, equipment);
        if (idxFuncLoc >= 0) {
            funcLocByPlantDesc.set(key, String(row[idxFuncLoc] ?? "").trim());
        }
    }

    return {
        byPlantDesc,
        funcLocByPlantDesc,
        duplicateKeys,
        totalRows,
        hasFunctionalLocColumn: idxFuncLoc >= 0
    };
}

/**
 * Load Equipment → Functional Loc. index (first wins per equipment).
 * @param {string} [filePath]
 * @returns {{
 *   byEquipment: Map<string, string>,
 *   entries: Array<{ equipment: string, funcLoc: string }>,
 *   totalRows: number
 * }}
 */
function loadEquipmentFuncLocIndex(filePath = DEFAULT_MAPPING_PATH) {
    if (!fs.existsSync(filePath)) {
        throw new Error(`File mapping tidak ditemukan: ${filePath}`);
    }

    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) {
        throw new Error(`Sheet kosong pada file mapping: ${filePath}`);
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rows.length < 1) {
        throw new Error(`Header mapping tidak ditemukan: ${filePath}`);
    }

    const headerRow = rows[0] || [];
    const idxEquip = findHeaderIndex(headerRow, "Equipment");
    const idxFuncLoc = findFirstHeaderIndex(headerRow, [
        "Functional Loc.",
        "DESCRIPTION_FUNC_LOCATION"
    ]);

    if (idxEquip < 0 || idxFuncLoc < 0) {
        throw new Error(
            "Kolom wajib mapping tidak lengkap: Equipment, Functional Loc. / DESCRIPTION_FUNC_LOCATION"
        );
    }

    /** @type {Map<string, string>} */
    const byEquipment = new Map();
    /** @type {Array<{ equipment: string, funcLoc: string }>} */
    const entries = [];
    let totalRows = 0;

    for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i] || [];
        const equipment = String(row[idxEquip] ?? "").trim();
        const funcLoc = String(row[idxFuncLoc] ?? "").trim();
        if (!equipment || !funcLoc) continue;

        totalRows += 1;
        entries.push({ equipment, funcLoc });
        if (!byEquipment.has(equipment)) {
            byEquipment.set(equipment, funcLoc);
        }
    }

    return { byEquipment, entries, totalRows };
}

module.exports = {
    DEFAULT_MAPPING_PATH,
    normalizePlant,
    normalizeDescription,
    normalizeFuncLocForCompare,
    makeKey,
    findFirstHeaderIndex,
    lookupFunctionalLoc,
    loadEquipmentNumberMapping,
    loadEquipmentFuncLocIndex
};
