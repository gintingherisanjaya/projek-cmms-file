/**
 * Baca hasil equipment-gathering (Drive / lokal) → indeks baris by FUNCTLOC DESC. AFTER APPLIED.
 */
const fs = require("fs");
const path = require("path");
const { readSpreadsheetRows } = require("./equipment_excel_io.cjs");
const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");
const { normalizeHeader } = require("./lsmw_cell_fill.cjs");
const {
    OUTPUT_HEADERS,
    PROTECT_Y_APPLY_HEADERS,
    FUNCTLOC_DESC_AFTER_APPLIED_HEADER,
    resolveSourceColumnIndices
} = require("./equipment_gathering_columns.cjs");
const { normalizeFunclocDesc } = require("./string_similarity.cjs");

function fileMatchKey(name) {
    return path
        .basename(String(name ?? ""))
        .trim()
        .toLowerCase()
        .replace(/\.xlsx$/i, "")
        .replace(/\s+/g, " ");
}

function resolveLocalPath(inputPath, cwd = process.cwd()) {
    const trimmed = String(inputPath ?? "").trim();
    if (!trimmed) return "";
    return path.normalize(
        path.isAbsolute(trimmed) ? trimmed : path.join(cwd, trimmed)
    );
}

function collectLocalSpreadsheets(dir, out = []) {
    if (!fs.existsSync(dir)) return out;
    for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, ent.name);
        if (ent.isDirectory()) {
            collectLocalSpreadsheets(full, out);
        } else if (
            ent.isFile() &&
            /\.xlsx$/i.test(ent.name) &&
            !ent.name.startsWith("~$")
        ) {
            out.push(full);
        }
    }
    return out;
}

/**
 * @param {string} runRoot — folder run equipment-gathering (berisi REGIONAL*)
 * @param {(name: string) => boolean} isRegionalFolderName
 * @returns {Array<{ regional: string, localPath: string, fileName: string }>}
 */
function listLocalRegionalSpreadsheets(runRoot, isRegionalFolderName) {
    const entries = [];
    if (!fs.existsSync(runRoot)) return entries;

    for (const dirent of fs.readdirSync(runRoot, { withFileTypes: true })) {
        if (!dirent.isDirectory()) continue;
        if (!isRegionalFolderName(dirent.name)) continue;

        const regional = dirent.name.trim();
        const files = collectLocalSpreadsheets(path.join(runRoot, dirent.name));
        for (const localPath of files) {
            const base = path.basename(localPath);
            if (base.toLowerCase() === "all-pks.xlsx") continue;
            entries.push({ regional, localPath, fileName: base });
        }
    }

    entries.sort(
        (a, b) =>
            a.regional.localeCompare(b.regional) ||
            a.fileName.localeCompare(b.fileName)
    );
    return entries;
}

/**
 * @param {string} localPath
 * @returns {{
 *   ok: true,
 *   rows: Array<{ values: unknown[], key: string, isDuplicateKey: boolean }>,
 *   duplicateKeys: string[]
 * } | { ok: false, reason: string }}
 */
async function readGatheringRowsFromFile(localPath) {
    if (!fs.existsSync(localPath)) {
        return { ok: false, reason: "File tidak ditemukan" };
    }

    const read = await readSpreadsheetRows(localPath);
    if (!read.ok) {
        return read;
    }

    const rawRows = read.rawRows;
    if (rawRows.length < 2) {
        return { ok: true, rows: [], duplicateKeys: [] };
    }

    const headerRow = rawRows[0];
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const sourceIndices = resolveSourceColumnIndices(headerRow);
    const idxApplied = colIndex[normalizeHeader(FUNCTLOC_DESC_AFTER_APPLIED_HEADER)];

    if (idxApplied === undefined) {
        return {
            ok: false,
            reason: `Kolom ${FUNCTLOC_DESC_AFTER_APPLIED_HEADER} tidak ada`
        };
    }

    /** @type {Array<{ values: unknown[], key: string, isDuplicateKey: boolean }>} */
    const rows = [];
    const duplicateKeys = [];
    const seenKeys = new Set();

    for (let i = 1; i < rawRows.length; i++) {
        const r = rawRows[i];
        const key = normalizeFunclocDesc(r[idxApplied]);
        let isDuplicateKey = false;

        if (key) {
            if (seenKeys.has(key)) {
                isDuplicateKey = true;
                duplicateKeys.push(key);
            } else {
                seenKeys.add(key);
            }
        }

        const values = OUTPUT_HEADERS.map(h => {
            const idx = sourceIndices.get(h);
            if (idx === undefined || r[idx] === undefined || r[idx] === null) {
                return "";
            }
            return r[idx];
        });

        rows.push({ values, key, isDuplicateKey });
    }

    return { ok: true, rows, duplicateKeys };
}

/**
 * @param {string} localPath
 * @returns {{
 *   ok: true,
 *   index: Map<string, Record<string, unknown>>,
 *   duplicateKeys: string[]
 * } | { ok: false, reason: string }}
 */
async function buildGatheringIndexFromFile(localPath) {
    const parsed = await readGatheringRowsFromFile(localPath);
    if (!parsed.ok) return parsed;

    /** @type {Map<string, Record<string, unknown>>} */
    const index = new Map();

    for (const row of parsed.rows) {
        if (!row.key || row.isDuplicateKey) continue;

        /** @type {Record<string, unknown>} */
        const rowData = {};
        for (const h of PROTECT_Y_APPLY_HEADERS) {
            const idx = OUTPUT_HEADERS.indexOf(h);
            rowData[h] = idx >= 0 ? row.values[idx] : "";
        }
        index.set(row.key, rowData);
    }

    return { ok: true, index, duplicateKeys: parsed.duplicateKeys };
}

const COL_FUNCTLOC_AFTER_APPLIED = OUTPUT_HEADERS.indexOf(
    FUNCTLOC_DESC_AFTER_APPLIED_HEADER
);

/**
 * Indeks apply dari baris output gathering in-memory (tanpa tulis/baca ulang file).
 * @param {Array<Array<unknown>>} outputRows — urutan OUTPUT_HEADERS
 */
function buildGatheringIndexFromOutputRows(outputRows) {
    /** @type {Map<string, Record<string, unknown>>} */
    const index = new Map();
    const duplicateKeys = [];
    const seenKeys = new Set();

    for (const row of outputRows) {
        const key = normalizeFunclocDesc(row[COL_FUNCTLOC_AFTER_APPLIED]);
        if (!key) continue;

        if (seenKeys.has(key)) {
            duplicateKeys.push(key);
            continue;
        }
        seenKeys.add(key);

        /** @type {Record<string, unknown>} */
        const rowData = {};
        for (const h of PROTECT_Y_APPLY_HEADERS) {
            const idx = OUTPUT_HEADERS.indexOf(h);
            rowData[h] = idx >= 0 ? row[idx] : "";
        }
        index.set(key, rowData);
    }

    return { ok: true, index, duplicateKeys };
}

/**
 * @param {Array<Array<unknown>>} outputRows
 * @returns {Array<{ values: unknown[], key: string, isDuplicateKey: boolean }>}
 */
function gatheringOutputRowsToParsedRows(outputRows) {
    const seenKeys = new Set();
    return outputRows.map(values => {
        const key = normalizeFunclocDesc(values[COL_FUNCTLOC_AFTER_APPLIED]);
        let isDuplicateKey = false;
        if (key) {
            if (seenKeys.has(key)) {
                isDuplicateKey = true;
            } else {
                seenKeys.add(key);
            }
        }
        return { values, key, isDuplicateKey };
    });
}

/**
 * @param {Array<{ values: unknown[], key: string, isDuplicateKey: boolean }>} gatheringRows
 * @param {Set<string>} appliedKeys
 * @param {{ allFailed?: boolean }} [options]
 * @returns {Array<{ values: unknown[], key: string, isDuplicateKey: boolean }>}
 */
function collectFailedGatheringRows(gatheringRows, appliedKeys, options = {}) {
    if (options.allFailed) {
        return [...gatheringRows];
    }

    return gatheringRows.filter(
        row => !row.key || (!appliedKeys.has(row.key) && !row.isDuplicateKey)
    );
}

/**
 * Filter baris LPP duplikat (FUNCTLOC DESC. AFTER APPLIED sama).
 * Applied tidak duplikat → keep first LPP row; Applied juga duplikat → keep all.
 * @param {Array<Array<unknown>>} outputRows
 * @param {Map<string, number>} appliedKeyCounts — dari countFunclocAfterKeys
 */
function filterOutputRowsByDuplicateAppliedKey(outputRows, appliedKeyCounts) {
    /** @type {Map<string, number>} */
    const outputKeyCounts = new Map();

    for (const row of outputRows) {
        const key = normalizeFunclocDesc(row[COL_FUNCTLOC_AFTER_APPLIED]);
        if (!key) continue;
        outputKeyCounts.set(key, (outputKeyCounts.get(key) ?? 0) + 1);
    }

    const seenKeepFirst = new Set();
    /** @type {string[]} */
    const duplicateKeysWarn = [];
    let removedLppDuplicates = 0;
    const rows = [];

    for (const row of outputRows) {
        const key = normalizeFunclocDesc(row[COL_FUNCTLOC_AFTER_APPLIED]);
        if (!key) {
            rows.push(row);
            continue;
        }

        if ((outputKeyCounts.get(key) ?? 0) <= 1) {
            rows.push(row);
            continue;
        }

        const appliedCount = appliedKeyCounts.get(key) ?? 0;
        if (appliedCount > 1) {
            rows.push(row);
            if (!duplicateKeysWarn.includes(key)) {
                duplicateKeysWarn.push(key);
            }
            continue;
        }

        if (!seenKeepFirst.has(key)) {
            seenKeepFirst.add(key);
            rows.push(row);
            if (!duplicateKeysWarn.includes(key)) {
                duplicateKeysWarn.push(key);
            }
        } else {
            removedLppDuplicates += 1;
        }
    }

    return { rows, removedLppDuplicates, duplicateKeysWarn };
}

/**
 * Indeks file gathering per fileMatchKey dari daftar regional lokal.
 * @returns {Map<string, string>} fileMatchKey → localPath
 */
function buildLocalGatheringFileIndex(regionalEntries) {
    const map = new Map();
    for (const { localPath, fileName } of regionalEntries) {
        const key = fileMatchKey(fileName);
        if (!map.has(key)) map.set(key, localPath);
    }
    return map;
}

/**
 * @param {Array<{ regional: string, file: object }>} regionalEntries
 * @returns {Map<string, { regional: string, file: object }>}
 */
function buildDriveGatheringFileIndex(regionalEntries) {
    const map = new Map();
    for (const entry of regionalEntries) {
        const key = fileMatchKey(entry.file.name);
        if (!map.has(key)) map.set(key, entry);
    }
    return map;
}

function validateGatheringHeaderRow(headerRow) {
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const missing = OUTPUT_HEADERS.filter(
        h => colIndex[normalizeHeader(h)] === undefined
    );
    if (missing.length > OUTPUT_HEADERS.length / 2) {
        return {
            ok: false,
            reason: "Header tidak sesuai format equipment-gathering"
        };
    }
    return { ok: true };
}

module.exports = {
    fileMatchKey,
    resolveLocalPath,
    listLocalRegionalSpreadsheets,
    readGatheringRowsFromFile,
    buildGatheringIndexFromFile,
    buildGatheringIndexFromOutputRows,
    gatheringOutputRowsToParsedRows,
    filterOutputRowsByDuplicateAppliedKey,
    collectFailedGatheringRows,
    buildLocalGatheringFileIndex,
    buildDriveGatheringFileIndex,
    validateGatheringHeaderRow
};
