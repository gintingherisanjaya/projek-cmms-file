const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DEFAULT_TRUTH_PATH = path.join(
    __dirname,
    "..",
    "TRUTH_OF_DATA_30_JANUARI.xlsx"
);

function findHeaderIndex(headerRow, expectedName) {
    const target = String(expectedName ?? "").trim().toUpperCase();
    for (let i = 0; i < headerRow.length; i += 1) {
        if (String(headerRow[i] ?? "").trim().toUpperCase() === target) return i;
    }
    return -1;
}

function normalizeFunctionalLocation(value) {
    return String(value ?? "").trim();
}

/**
 * Load TRUTH_OF_DATA: Functional Location → FunctLocDescrip. (first wins).
 * @param {string} [filePath]
 * @returns {{ byFunctionalLocation: Map<string, string>, totalRows: number, duplicateKeys: number }}
 */
function loadFunclocDescByFunctionalLocation(filePath = DEFAULT_TRUTH_PATH) {
    if (!fs.existsSync(filePath)) {
        throw new Error(`File Truth of Data tidak ditemukan: ${filePath}`);
    }

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames.includes("Data")
        ? "Data"
        : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        throw new Error(`Sheet kosong pada file Truth of Data: ${filePath}`);
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rows.length < 1) {
        throw new Error(`Header Truth of Data tidak ditemukan: ${filePath}`);
    }

    const headerRow = rows[0] || [];
    const idxFuncLoc = findHeaderIndex(headerRow, "Functional Location");
    const idxDesc = findHeaderIndex(headerRow, "FunctLocDescrip.");

    if (idxFuncLoc < 0 || idxDesc < 0) {
        throw new Error(
            "Kolom wajib Truth of Data tidak lengkap: Functional Location, FunctLocDescrip."
        );
    }

    const byFunctionalLocation = new Map();
    let totalRows = 0;
    let duplicateKeys = 0;

    for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i] || [];
        const funcLoc = normalizeFunctionalLocation(row[idxFuncLoc]);
        if (!funcLoc) continue;

        const desc = String(row[idxDesc] ?? "").trim();
        totalRows += 1;

        if (byFunctionalLocation.has(funcLoc)) {
            duplicateKeys += 1;
            continue;
        }
        byFunctionalLocation.set(funcLoc, desc);
    }

    return { byFunctionalLocation, totalRows, duplicateKeys };
}

module.exports = {
    DEFAULT_TRUTH_PATH,
    loadFunclocDescByFunctionalLocation
};
