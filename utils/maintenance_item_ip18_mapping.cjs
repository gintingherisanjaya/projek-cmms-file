const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DEFAULT_IP18_PATH = path.join(
    __dirname,
    "..",
    "maintenance-item-ip18.xlsx"
);

function findHeaderIndex(headerRow, expectedName) {
    const target = String(expectedName ?? "").trim().toUpperCase();
    for (let i = 0; i < headerRow.length; i += 1) {
        if (String(headerRow[i] ?? "").trim().toUpperCase() === target) return i;
    }
    return -1;
}

function normalizeEquipment(value) {
    return String(value ?? "").trim();
}

function normalizeMaintItem(value) {
    return String(value ?? "").trim();
}

/**
 * Ordered unique MaintItem values (first occurrence wins).
 * @param {string[]} items
 * @returns {string[]}
 */
function uniqueMaintItemsOrdered(items) {
    const seen = new Set();
    const out = [];
    for (const raw of items) {
        const mi = normalizeMaintItem(raw);
        if (!mi || seen.has(mi)) continue;
        seen.add(mi);
        out.push(mi);
    }
    return out;
}

/**
 * Load maintenance-item-ip18: Equipment → MaintItem[] (file order, unique).
 * @param {string} [filePath]
 * @returns {{ byEquipment: Map<string, string[]>, totalRows: number }}
 */
function loadMaintItemsByEquipment(filePath = DEFAULT_IP18_PATH) {
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
    const idxMaint = findHeaderIndex(headerRow, "MaintItem");

    if (idxEquip < 0 || idxMaint < 0) {
        throw new Error(
            "Kolom wajib maintenance-item-ip18 tidak lengkap: Equipment, MaintItem"
        );
    }

    /** @type {Map<string, string[]>} */
    const rawByEquipment = new Map();
    let totalRows = 0;

    for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i] || [];
        const equipment = normalizeEquipment(row[idxEquip]);
        const maintItem = normalizeMaintItem(row[idxMaint]);
        if (!equipment || !maintItem) continue;

        totalRows += 1;
        if (!rawByEquipment.has(equipment)) {
            rawByEquipment.set(equipment, []);
        }
        rawByEquipment.get(equipment).push(maintItem);
    }

    const byEquipment = new Map();
    for (const [equipment, items] of rawByEquipment) {
        byEquipment.set(equipment, uniqueMaintItemsOrdered(items));
    }

    return { byEquipment, totalRows };
}

/**
 * @param {string[]} maintItems
 * @returns {{ slow: Record<string, string>, truncated: boolean, uniqueCount: number }}
 */
/**
 * Satu MaintItem per maintenance plan (first wins).
 * @param {string[]} maintItems
 * @returns {{ maintItem: string, extraCount: number }}
 */
function lookupFirstMaintItem(maintItems) {
    const unique = uniqueMaintItemsOrdered(maintItems);
    return {
        maintItem: unique[0] ?? "",
        extraCount: Math.max(0, unique.length - 1)
    };
}

function fillSlowColumns(maintItems) {
    const unique = uniqueMaintItemsOrdered(maintItems);
    /** @type {Record<string, string>} */
    const slow = {};

    if (unique.length === 0) {
        return { slow, truncated: false, uniqueCount: 0 };
    }

    if (unique.length === 1) {
        slow.SLOW_I_01 = unique[0];
        slow.SLOW_I_02 = unique[0];
        return { slow, truncated: false, uniqueCount: 1 };
    }

    slow.SLOW_I_01 = unique[0];
    slow.SLOW_I_02 = unique[0];
    for (let i = 1; i < unique.length && i + 2 <= 7; i += 1) {
        const slot = String(i + 2).padStart(2, "0");
        slow[`SLOW_I_${slot}`] = unique[i];
    }

    return {
        slow,
        truncated: unique.length > 6,
        uniqueCount: unique.length
    };
}

module.exports = {
    DEFAULT_IP18_PATH,
    loadMaintItemsByEquipment,
    uniqueMaintItemsOrdered,
    lookupFirstMaintItem,
    fillSlowColumns
};
