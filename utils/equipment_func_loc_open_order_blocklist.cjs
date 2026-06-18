const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DEFAULT_BLOCKLIST_PATH = path.join(
    __dirname,
    "..",
    "equipment-func-loc-belum-close-order.xlsx"
);

function findHeaderIndex(headerRow, expectedName) {
    const target = String(expectedName ?? "").trim().toUpperCase();
    for (let i = 0; i < headerRow.length; i += 1) {
        if (String(headerRow[i] ?? "").trim().toUpperCase() === target) return i;
    }
    return -1;
}

function normalizeValue(value) {
    return String(value ?? "").trim();
}

/**
 * Load equipment & functional location yang masih punya order terbuka.
 * @param {string} [filePath]
 * @returns {{
 *   blockedEquipment: Set<string>,
 *   blockedFuncLoc: Set<string>,
 *   totalRows: number,
 *   equipmentCount: number,
 *   funcLocCount: number
 * }}
 */
function loadOpenOrderBlocklist(filePath = DEFAULT_BLOCKLIST_PATH) {
    if (!fs.existsSync(filePath)) {
        throw new Error(
            `File blocklist order belum close tidak ditemukan: ${filePath}`
        );
    }

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        throw new Error(
            `Sheet kosong pada file blocklist order belum close: ${filePath}`
        );
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rows.length < 1) {
        throw new Error(
            `Header blocklist order belum close tidak ditemukan: ${filePath}`
        );
    }

    const headerRow = rows[0] || [];
    const idxEquipment = findHeaderIndex(headerRow, "Equipment");
    const idxFuncLoc = findHeaderIndex(headerRow, "Functional Loc.");

    if (idxEquipment < 0 || idxFuncLoc < 0) {
        throw new Error(
            "Kolom wajib blocklist order belum close tidak lengkap: Equipment, Functional Loc."
        );
    }

    const blockedEquipment = new Set();
    const blockedFuncLoc = new Set();
    let totalRows = 0;

    for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i] || [];
        const equipment = normalizeValue(row[idxEquipment]);
        const funcLoc = normalizeValue(row[idxFuncLoc]);

        if (!equipment && !funcLoc) continue;

        totalRows += 1;
        if (equipment) blockedEquipment.add(equipment);
        if (funcLoc) blockedFuncLoc.add(funcLoc);
    }

    return {
        blockedEquipment,
        blockedFuncLoc,
        totalRows,
        equipmentCount: blockedEquipment.size,
        funcLocCount: blockedFuncLoc.size
    };
}

module.exports = {
    DEFAULT_BLOCKLIST_PATH,
    loadOpenOrderBlocklist
};
