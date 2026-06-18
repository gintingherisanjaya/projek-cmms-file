/**
 * Alias deskripsi funcloc: mapping-alias-desc-func-loc.xlsx
 * NILAI ASAL → ALTERNATIVE NAME (untuk PLTXT / SHTXT* di template LSMW).
 */

const path = require("path");
const XLSX = require("xlsx");

const MAPPING_PATH = path.join(
    __dirname,
    "..",
    "mapping-alias-desc-func-loc.xlsx"
);

let aliasByKey = null;

function normalizeDescKey(value) {
    return String(value ?? "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();
}

function findColumnIndex(headerRow, candidates) {
    const norm = candidates.map(c => normalizeDescKey(c));
    for (let i = 0; i < headerRow.length; i++) {
        const h = normalizeDescKey(headerRow[i]);
        if (norm.includes(h)) return i;
    }
    return -1;
}

function loadAliasMapFromWorkbook() {
    const wb = XLSX.readFile(MAPPING_PATH);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rows.length < 2) return new Map();

    const headerRow = rows[0];
    const idxAsal = findColumnIndex(headerRow, [
        "NILAI ASAL",
        "NILAI_ASAL"
    ]);
    const idxAlt = findColumnIndex(headerRow, [
        "ALTERNATIVE NAME",
        "ALTERNATIVE NAME ",
        "ALTERNATIVE_NAME"
    ]);

    if (idxAsal < 0 || idxAlt < 0) {
        throw new Error(
            "mapping-alias-desc-func-loc.xlsx: kolom NILAI ASAL / ALTERNATIVE NAME tidak ditemukan"
        );
    }

    const map = new Map();
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const asal = String(row[idxAsal] ?? "").trim();
        const alt = String(row[idxAlt] ?? "").trim();
        if (!asal || !alt) continue;
        const key = normalizeDescKey(asal);
        if (!map.has(key)) map.set(key, alt);
    }
    return map;
}

async function ensureFunclocDescAliasMap() {
    if (aliasByKey) return aliasByKey;
    aliasByKey = loadAliasMapFromWorkbook();
    console.log(
        `[funcloc-desc-alias] ${aliasByKey.size} entri dari mapping-alias-desc-func-loc.xlsx`
    );
    return aliasByKey;
}

/**
 * Koreksi typo tetap di file sumber Protect Y (sebelum alias / proses lain).
 * @param {unknown} rawDesc
 * @returns {string}
 */
function fixFunclocDescSourceTypos(rawDesc) {
    const s = String(rawDesc ?? "");
    if (!/\bELECTROMOTR\b/i.test(s)) return s;
    return s.replace(/\bELECTROMOTR\b/gi, "ELECTROMOTOR");
}

/**
 * Untuk isian template LSMW: pakai ALTERNATIVE NAME jika NILAI ASAL cocok, else nilai asal.
 * @param {unknown} rawDesc
 * @returns {string}
 */
function applyFunclocDescAlias(rawDesc) {
    const raw = String(rawDesc ?? "").trim();
    if (!raw) return "";
    if (!aliasByKey) return raw;
    const alt = aliasByKey.get(normalizeDescKey(raw));
    return alt !== undefined ? alt : raw;
}

module.exports = {
    MAPPING_PATH,
    ensureFunclocDescAliasMap,
    fixFunclocDescSourceTypos,
    applyFunclocDescAlias,
    normalizeDescKey
};
