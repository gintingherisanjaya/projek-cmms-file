const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DEFAULT_IK07_PATH = path.join(
    __dirname,
    "..",
    "measuring-item-ik07.xlsx"
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

/**
 * Load measuring-item-ik07: Equipment → Measuring point (first wins).
 * @param {string} [filePath]
 * @returns {{ byEquipment: Map<string, string>, totalRows: number, duplicateEquipment: number }}
 */
function loadMeasuringPointByEquipment(filePath = DEFAULT_IK07_PATH) {
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
    const idxPoint = findHeaderIndex(headerRow, "Measuring point");

    if (idxEquip < 0 || idxPoint < 0) {
        throw new Error(
            "Kolom wajib measuring-item-ik07 tidak lengkap: Equipment, Measuring point"
        );
    }

    const byEquipment = new Map();
    let totalRows = 0;
    let duplicateEquipment = 0;

    for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i] || [];
        const equipment = normalizeEquipment(row[idxEquip]);
        const point = String(row[idxPoint] ?? "").trim();
        if (!equipment || !point) continue;

        totalRows += 1;
        if (byEquipment.has(equipment)) {
            duplicateEquipment += 1;
            continue;
        }
        byEquipment.set(equipment, point);
    }

    return { byEquipment, totalRows, duplicateEquipment };
}

module.exports = {
    DEFAULT_IK07_PATH,
    loadMeasuringPointByEquipment
};
