/**
 * Kolom equipment-gathering & deteksi FUNCTLOC DESC AFTER LEVEL 1,2,3
 * (logika mengacu pada lsmw_equipment_v0.cjs, tanpa mengubah file tersebut).
 */
const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");
const { normalizeHeader } = require("./lsmw_cell_fill.cjs");
const { normalizeFunclocDesc } = require("./string_similarity.cjs");

/** Header baris output (urutan tetap). */
const OUTPUT_HEADERS = [
    "REGIONAL",
    "POM",
    "EQUIPMENT NUMBER",
    "MAINTENANCE PLANT",
    "COMPANY CODE",
    "COST CENTER BEFORE",
    "COST CENTER AFTER",
    "EQKTU BEFORE",
    "OBJECT NUMBER",
    "EQUIPMENT CATEGORY",
    "EQUIPMENT GROUP BEFORE",
    "CONSTRUCTION YEAR",
    "GROES",
    "MANUFACTURE",
    "MEASURING POINT",
    "MeasurmntRangeUnit",
    "MAINTENANCE PLANNER GROUP",
    "WORK CENTER",
    "FUNCTIONAL LOCATION BEFORE",
    "FUNCTLOC DESC. BEFORE",
    "FUNCTLOC DESC. AFTER LEVEL 1,2,3",
    "FUNCTLOC DESC. AFTER APPLIED"
];

const FUNCTLOC_DESC_AFTER_LEVEL123_HEADER = "FUNCTLOC DESC. AFTER LEVEL 1,2,3";
const FUNCTLOC_DESC_AFTER_APPLIED_HEADER = "FUNCTLOC DESC. AFTER APPLIED";

/** Kolom before dari gathering yang ditulis ke Protect Y (REGIONAL … FUNCTLOC DESC. BEFORE). */
const PROTECT_Y_APPLY_HEADERS = OUTPUT_HEADERS.filter(
    h =>
        h !== "COST CENTER AFTER" &&
        h !== FUNCTLOC_DESC_AFTER_LEVEL123_HEADER &&
        h !== FUNCTLOC_DESC_AFTER_APPLIED_HEADER
);

/** Kolom tetap: nama output → kunci lookup header (normalizeHeader). */
const FIXED_COLUMN_LOOKUPS = {
    REGIONAL: ["REGIONAL"],
    POM: ["POM"],
    "EQUIPMENT NUMBER": ["EQUIPMENT NUMBER"],
    "MAINTENANCE PLANT": ["MAINTENANCE PLANT"],
    "COMPANY CODE": ["COMPANY CODE"],
    "COST CENTER BEFORE": ["COST CENTER BEFORE"],
    "COST CENTER AFTER": ["COST CENTER AFTER"],
    "EQKTU BEFORE": ["EQKTU BEFORE"],
    "OBJECT NUMBER": ["OBJECT NUMBER"],
    "EQUIPMENT CATEGORY": ["EQUIPMENT CATEGORY"],
    "EQUIPMENT GROUP BEFORE": ["EQUIPMENT GROUP BEFORE"],
    "CONSTRUCTION YEAR": ["CONSTRUCTION YEAR"],
    GROES: ["GROES"],
    MANUFACTURE: ["MANUFACTURE"],
    "MEASURING POINT": ["MEASURING POINT"],
    MeasurmntRangeUnit: ["MEASURMNTRANGEUNIT", "MEASURING RANGE UNIT"],
    "MAINTENANCE PLANNER GROUP": ["MAINTENANCE PLANNER GROUP"],
    "WORK CENTER": ["WORK CENTER"],
    "FUNCTIONAL LOCATION BEFORE": ["FUNCTIONAL LOCATION BEFORE"],
    "FUNCTLOC DESC. BEFORE": ["FUNCTLOC DESC. BEFORE"]
};

const KNOWN_FUNCTLOC_DESC_AFTER_LEVEL123 = [
    "FUNCTLOC DESC. AFTER LEVEL 1,2,3",
    "FUNCTLOC DESC AFTRER LEVEL 1,2,3",
    "EFUNCTLOC DESC AFTER LEVEL 1,2,3"
];

/** Header kolom regional — mengandung kata REGIONAL (mis. "Regional I Eks N3"). */
function isRegionalHeader(header) {
    const h = normalizeHeader(header);
    return h.includes("REGIONAL");
}

function findRegionalColumnIndex(headerRow) {
    if (!headerRow) return undefined;

    for (let i = 0; i < headerRow.length; i++) {
        if (isRegionalHeader(headerRow[i])) return i;
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    if (colIndex.REGIONAL !== undefined) return colIndex.REGIONAL;

    return undefined;
}

/** Header kolom company code — mengandung frasa COMPANY CODE. */
function isCompanyCodeHeader(header) {
    const h = normalizeHeader(header);
    return h.includes("COMPANY CODE");
}

function findCompanyCodeColumnIndex(headerRow) {
    if (!headerRow) return undefined;

    for (let i = 0; i < headerRow.length; i++) {
        if (isCompanyCodeHeader(headerRow[i])) return i;
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    if (colIndex["COMPANY CODE"] !== undefined) return colIndex["COMPANY CODE"];

    return undefined;
}

/** Header kolom maintenance plant — mengandung frasa MAINTENANCE PLANT. */
function isMaintenancePlantHeader(header) {
    const h = normalizeHeader(header);
    return h.includes("MAINTENANCE PLANT");
}

function findMaintenancePlantColumnIndex(headerRow) {
    if (!headerRow) return undefined;

    for (let i = 0; i < headerRow.length; i++) {
        if (isMaintenancePlantHeader(headerRow[i])) return i;
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    if (colIndex["MAINTENANCE PLANT"] !== undefined) return colIndex["MAINTENANCE PLANT"];

    return undefined;
}

const MEASURMNTRANGEUNIT_LOOKUP_KEYS = ["MEASURMNTRANGEUNIT", "MEASURING RANGE UNIT"];

/** Header MeasurmntRangeUnit — exact name atau alias "Column1" di posisi tetap. */
function isMeasurmntRangeUnitHeader(header) {
    const h = normalizeHeader(header);
    if (MEASURMNTRANGEUNIT_LOOKUP_KEYS.includes(h)) return true;
    return h === "COLUMN1";
}

function findMeasurmntRangeUnitColumnIndex(headerRow) {
    if (!headerRow) return undefined;

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);

    for (const key of MEASURMNTRANGEUNIT_LOOKUP_KEYS) {
        if (colIndex[key] !== undefined) return colIndex[key];
    }

    const idxMeasuringPoint = colIndex["MEASURING POINT"];
    const idxPlannerGroup = colIndex["MAINTENANCE PLANNER GROUP"];

    if (idxMeasuringPoint !== undefined && idxPlannerGroup !== undefined) {
        const candidate = idxMeasuringPoint + 1;
        if (
            candidate < idxPlannerGroup &&
            normalizeHeader(headerRow[candidate]) === "COLUMN1"
        ) {
            return candidate;
        }
    }

    return undefined;
}

/**
 * Indeks kolom Protect Y untuk header apply (REGIONAL, MAINTENANCE PLANT, COMPANY CODE & MeasurmntRangeUnit dinamis).
 */
function resolveProtectYApplyColumnIndex(headerRow, colByHeader, outputHeaderName) {
    if (outputHeaderName === "REGIONAL") {
        return findRegionalColumnIndex(headerRow);
    }
    if (outputHeaderName === "MAINTENANCE PLANT") {
        return findMaintenancePlantColumnIndex(headerRow);
    }
    if (outputHeaderName === "COMPANY CODE") {
        return findCompanyCodeColumnIndex(headerRow);
    }
    if (outputHeaderName === "MeasurmntRangeUnit") {
        return findMeasurmntRangeUnitColumnIndex(headerRow);
    }
    return colByHeader[normalizeHeader(outputHeaderName)];
}

/** Header kolom deskripsi funcloc "after" — wajib FUNCTLOC + DESC + AFTER; bukan BEFORE. */
function isFunclocDescAfterHeader(header) {
    const h = normalizeHeader(header);
    if (!h) return false;
    if (/\bBEFORE\b/.test(h)) return false;
    if (!h.includes("FUNCTLOC")) return false;
    if (!/\bDESC\b/.test(h)) return false;
    if (!/\b(AFTER|AFTRER)\b/.test(h)) return false;
    return true;
}

function isFunclocDescAfterLevel123Header(header) {
    if (!isFunclocDescAfterHeader(header)) return false;
    const raw = String(header ?? "");
    const h = normalizeHeader(header);
    if (!/\bLEVEL\b/.test(h)) return false;
    if (/1\s*,\s*2\s*,\s*3/.test(raw) || /1\s*,\s*2\s*,\s*3/.test(h)) return true;
    if (/1,2,3/i.test(raw.replace(/\s/g, ""))) return true;
    return false;
}

function findFunclocDescAfterColumnIndex(headerRow) {
    const indices = [];
    for (let i = 0; i < headerRow.length; i++) {
        if (isFunclocDescAfterHeader(headerRow[i])) indices.push(i);
    }
    if (indices.length === 0) return undefined;
    if (indices.length === 1) return indices[0];
    const withLevel = indices.filter(i =>
        /\bLEVEL\b/.test(normalizeHeader(headerRow[i]))
    );
    if (withLevel.length >= 1) return withLevel[0];
    return indices[0];
}

function findFunclocDescAfterLevel123ColumnIndex(headerRow) {
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);

    for (const alias of KNOWN_FUNCTLOC_DESC_AFTER_LEVEL123) {
        const idx = colIndex[normalizeHeader(alias)];
        if (idx !== undefined) return idx;
    }

    for (let i = 0; i < headerRow.length; i++) {
        if (isFunclocDescAfterLevel123Header(headerRow[i])) return i;
    }

    const indices = [];
    for (let i = 0; i < headerRow.length; i++) {
        if (isFunclocDescAfterHeader(headerRow[i])) indices.push(i);
    }
    const withLevel = indices.filter(i =>
        /\bLEVEL\b/.test(normalizeHeader(headerRow[i]))
    );
    if (withLevel.length === 1) return withLevel[0];
    if (withLevel.length > 1) {
        const combined = withLevel.filter(i =>
            isFunclocDescAfterLevel123Header(headerRow[i])
        );
        if (combined.length > 0) return combined[0];
        return withLevel[0];
    }

    return findFunclocDescAfterColumnIndex(headerRow);
}

/** Header persis COST CENTER (bukan BEFORE / AFTER). */
function isPlainCostCenterHeader(header) {
    return normalizeHeader(header) === "COST CENTER";
}

/**
 * Pilih kolom cost center dari header (kiri ke kanan per tier):
 * 1) mengandung COST CENTER AFTER
 * 2) mengandung COST CENTER, tanpa kata BEFORE
 * 3) mengandung COST CENTER BEFORE
 */
function findCostCenterColumnIndex(headerRow) {
    if (!headerRow) return undefined;

    const tiers = [
        h => h.includes("COST CENTER") && h.includes("AFTER"),
        h => h.includes("COST CENTER") && !h.includes("BEFORE"),
        h => h.includes("COST CENTER") && h.includes("BEFORE")
    ];

    for (const match of tiers) {
        for (let i = 0; i < headerRow.length; i++) {
            const h = normalizeHeader(headerRow[i]);
            if (!h) continue;
            if (match(h)) return i;
        }
    }

    return undefined;
}

function findFunclocDescBeforeColumnIndex(headerRow) {
    if (!headerRow) return undefined;

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    if (colIndex["FUNCTLOC DESC. BEFORE"] !== undefined) {
        return colIndex["FUNCTLOC DESC. BEFORE"];
    }

    for (let i = 0; i < headerRow.length; i++) {
        const h = normalizeHeader(headerRow[i]);
        if (!h.includes("FUNCTLOC")) continue;
        if (!/\bDESC\b/.test(h)) continue;
        if (!/\bBEFORE\b/.test(h)) continue;
        return i;
    }

    return undefined;
}

/** Alias: pemilihan kolom cost center untuk LSMW & gathering (prioritas sama). */
function findCostCenterAfterColumnIndex(headerRow) {
    return findCostCenterColumnIndex(headerRow);
}

/**
 * @returns {Map<string, number|undefined>} output header → kolom sumber (0-based)
 */
function resolveSourceColumnIndices(headerRow) {
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const map = new Map();

    for (const outputName of OUTPUT_HEADERS) {
        if (outputName === FUNCTLOC_DESC_AFTER_APPLIED_HEADER) {
            map.set(
                outputName,
                colIndex[normalizeHeader(FUNCTLOC_DESC_AFTER_APPLIED_HEADER)]
            );
            continue;
        }
        if (outputName === FUNCTLOC_DESC_AFTER_LEVEL123_HEADER) {
            map.set(outputName, findFunclocDescAfterLevel123ColumnIndex(headerRow));
            continue;
        }
        if (outputName === "REGIONAL") {
            map.set(outputName, findRegionalColumnIndex(headerRow));
            continue;
        }
        if (outputName === "COMPANY CODE") {
            map.set(outputName, findCompanyCodeColumnIndex(headerRow));
            continue;
        }
        if (outputName === "MAINTENANCE PLANT") {
            map.set(outputName, findMaintenancePlantColumnIndex(headerRow));
            continue;
        }
        if (outputName === "MeasurmntRangeUnit") {
            map.set(outputName, findMeasurmntRangeUnitColumnIndex(headerRow));
            continue;
        }
        if (outputName === "COST CENTER AFTER") {
            map.set(outputName, findCostCenterAfterColumnIndex(headerRow));
            continue;
        }

        const keys = FIXED_COLUMN_LOOKUPS[outputName] ?? [outputName];
        let idx;
        for (const key of keys) {
            if (colIndex[key] !== undefined) {
                idx = colIndex[key];
                break;
            }
        }
        map.set(outputName, idx);
    }

    return map;
}

/**
 * Kolom output yang tidak ter-resolve dari header sumber (LPP / output gathering).
 * FUNCTLOC DESC. AFTER APPLIED dikecualikan saat membaca LPP (belum ada di sumber).
 */
function listMissingSourceColumns(indices, { excludeApplied = true } = {}) {
    return OUTPUT_HEADERS.filter(h => {
        if (excludeApplied && h === FUNCTLOC_DESC_AFTER_APPLIED_HEADER) {
            return false;
        }
        return indices.get(h) === undefined;
    });
}

function cellValue(row, colIndex) {
    if (colIndex === undefined || colIndex < 0) return "";
    const v = row[colIndex];
    if (v === null || v === undefined) return "";
    return v;
}

function normalizeEquipmentNumber(value) {
    if (value === null || value === undefined || value === "") return "";
    return String(value).trim();
}

function normalizeCostCenter(value) {
    if (value === null || value === undefined || value === "") return "";
    return String(value).trim().toUpperCase();
}

function hasCellText(value) {
    if (value === null || value === undefined) return false;
    return String(value).trim() !== "";
}

/** Nilai pertama yang terisi di seluruh baris data (satu file = satu REGIONAL/POM). */
function findFileLevelColumnValue(dataRows, colIndex) {
    if (colIndex === undefined) return "";
    for (const r of dataRows) {
        const v = cellValue(r, colIndex);
        if (hasCellText(v)) return v;
    }
    return "";
}

/**
 * @param {object} [fallbacks]
 * @param {string} [fallbacks.regional] — mis. nama folder REGIONAL 1
 * @param {string} [fallbacks.pom] — mis. dari nama file PKS
 */
function extractGatherRows(headerRow, dataRows, fallbacks = {}) {
    const indices = resolveSourceColumnIndices(headerRow);
    const idxEquip = indices.get("EQUIPMENT NUMBER");
    const idxRegional = indices.get("REGIONAL");
    const idxPom = indices.get("POM");

    const fileRegional =
        findFileLevelColumnValue(dataRows, idxRegional) ||
        String(fallbacks.regional ?? "").trim();
    const filePom =
        findFileLevelColumnValue(dataRows, idxPom) ||
        String(fallbacks.pom ?? "").trim();

    const colRegional = OUTPUT_HEADERS.indexOf("REGIONAL");
    const colPom = OUTPUT_HEADERS.indexOf("POM");
    const idxCostCenterBefore = indices.get("COST CENTER BEFORE");
    const rows = [];
    let skippedNonStas = 0;

    for (const r of dataRows) {
        if (!normalizeEquipmentNumber(cellValue(r, idxEquip))) continue;

        const ccBefore = normalizeCostCenter(cellValue(r, idxCostCenterBefore));
        if (!ccBefore.includes("STAS")) {
            skippedNonStas += 1;
            continue;
        }

        const values = OUTPUT_HEADERS.map(h => cellValue(r, indices.get(h)));

        if (fileRegional) {
            values[colRegional] = fileRegional;
        }
        if (filePom) {
            values[colPom] = filePom;
        }

        rows.push(values);
    }

    return { rows, indices, fileRegional, filePom, skippedNonStas };
}

const COL_REGIONAL_OUTPUT = OUTPUT_HEADERS.indexOf("REGIONAL");
const COL_POM_OUTPUT = OUTPUT_HEADERS.indexOf("POM");

/**
 * Satu nilai REGIONAL + POM per file PKS (dari baris gathering atau fallback).
 * @param {Array<Array<unknown>>} outputRows
 * @param {{ regional?: string, pom?: string }} [fallbacks]
 */
function resolveFileRegionalPom(outputRows, fallbacks = {}) {
    let regional = "";
    let pom = "";

    for (const row of outputRows) {
        if (!regional && hasCellText(row[COL_REGIONAL_OUTPUT])) {
            regional = String(row[COL_REGIONAL_OUTPUT]).trim();
        }
        if (!pom && hasCellText(row[COL_POM_OUTPUT])) {
            pom = String(row[COL_POM_OUTPUT]).trim();
        }
        if (regional && pom) break;
    }

    if (!regional) {
        regional = String(fallbacks.regional ?? "").trim();
    }
    if (!pom) {
        pom = String(fallbacks.pom ?? "").trim();
    }

    return { regional, pom };
}

/**
 * Samakan REGIONAL dan POM di semua baris output gathering.
 * @param {Array<Array<unknown>>} outputRows
 * @param {string} regional
 * @param {string} pom
 */
function normalizeOutputRowsRegionalPom(outputRows, regional, pom) {
    for (const row of outputRows) {
        if (regional) row[COL_REGIONAL_OUTPUT] = regional;
        if (pom) row[COL_POM_OUTPUT] = pom;
    }
}

/**
 * Hitung baris Protect Y/Applied per FUNCTLOC DESC. AFTER (key dinormalisasi).
 * @returns {Map<string, number>}
 */
function countFunclocAfterKeys(headerRow, dataRows) {
    /** @type {Map<string, number>} */
    const counts = new Map();
    const idxFuncloc = findFunclocDescAfterLevel123ColumnIndex(headerRow);
    if (idxFuncloc === undefined) return counts;

    for (const r of dataRows) {
        const key = normalizeFunclocDesc(cellValue(r, idxFuncloc));
        if (!key) continue;
        counts.set(key, (counts.get(key) ?? 0) + 1);
    }

    return counts;
}

/**
 * Satu kandidat per funcloc (baris pertama menang).
 * @param {Array<{ costCenterAfter: string, funclocDescAfter: string }>} candidates
 */
function dedupeAppliedCandidatesByFuncloc(candidates) {
    const seen = new Set();
    const out = [];

    for (const c of candidates) {
        const key = normalizeFunclocDesc(c.funclocDescAfter);
        if (!key || seen.has(key)) continue;
        seen.add(key);
        out.push(c);
    }

    return out;
}

/**
 * Kandidat baris Applied untuk matching.
 * Baris dengan FUNCTLOC DESC AFTER terisi (parent sering tanpa EQUIPMENT NUMBER).
 * @returns {{ ok: true, candidates: Array<{ costCenterAfter: string, funclocDescAfter: string }> } | { ok: false, reason: string }}
 */
function extractAppliedMatchCandidates(headerRow, dataRows) {
    if (!headerRow) {
        return { ok: false, reason: "Header tidak ditemukan" };
    }

    const idxCostCenterAfter = findCostCenterAfterColumnIndex(headerRow);
    if (idxCostCenterAfter === undefined) {
        return { ok: false, reason: "Kolom COST CENTER AFTER tidak ada" };
    }

    const idxFuncloc = findFunclocDescAfterLevel123ColumnIndex(headerRow);
    const candidates = [];

    for (const r of dataRows) {
        const funclocDescAfter = cellValue(r, idxFuncloc);
        if (!hasCellText(funclocDescAfter)) continue;

        candidates.push({
            costCenterAfter: cellValue(r, idxCostCenterAfter),
            funclocDescAfter
        });
    }

    return { ok: true, candidates: dedupeAppliedCandidatesByFuncloc(candidates) };
}

/**
 * @param {Array<Array<unknown>>} lppRows
 * @param {Map<number, string>} appliedMatches — lppRowIndex → nilai Applied
 * @returns {Array<Array<unknown>>}
 */
function appendAppliedColumn(lppRows, appliedMatches) {
    const colApplied = OUTPUT_HEADERS.indexOf(FUNCTLOC_DESC_AFTER_APPLIED_HEADER);

    return lppRows.map((row, idx) => {
        const out = [...row];
        while (out.length < OUTPUT_HEADERS.length) {
            out.push("");
        }
        out[colApplied] = appliedMatches.get(idx) ?? "";
        return out;
    });
}

/** Header FUNCTIONAL LOCATION AFTER — exact atau alias AFTER2; bukan BEFORE. */
function isFunctionalLocationAfterHeader(header) {
    const h = normalizeHeader(header);
    if (!h) return false;
    if (/\bBEFORE\b/.test(h)) return false;
    if (h === "FUNCTIONAL LOCATION AFTER") return true;
    if (/^FUNCTIONAL LOCATION AFTER\d+$/.test(h)) return true;
    return false;
}

function rowHasFunctionalLocationAfterHeader(row) {
    if (!row) return false;
    for (let c = 0; c < row.length; c++) {
        if (isFunctionalLocationAfterHeader(row[c])) return true;
    }
    return false;
}

/** Header FUNCTIONAL LOCATION BEFORE — exact; bukan AFTER. */
function isFunctionalLocationBeforeHeader(header) {
    const h = normalizeHeader(header);
    if (!h) return false;
    if (/\bAFTER\b/.test(h)) return false;
    return h === "FUNCTIONAL LOCATION BEFORE";
}

/**
 * Indeks kolom FUNCTIONAL LOCATION BEFORE (first wins).
 * @param {Array<unknown>} headerRow
 * @returns {number|undefined}
 */
function findFunctionalLocationBeforeColumnIndex(headerRow) {
    if (!headerRow) return undefined;

    for (let i = 0; i < headerRow.length; i += 1) {
        if (isFunctionalLocationBeforeHeader(headerRow[i])) return i;
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    if (colIndex["FUNCTIONAL LOCATION BEFORE"] !== undefined) {
        return colIndex["FUNCTIONAL LOCATION BEFORE"];
    }

    return undefined;
}

/** Baris header khas export LPP (EQUIPMENT NUMBER + REGIONAL/POM/COST CENTER BEFORE). */
function rowLooksLikeLppGatheringHeader(row) {
    if (!row || row.length === 0) return false;

    const colIndex = buildColumnIndexFirstWins(row, normalizeHeader);
    if (colIndex["EQUIPMENT NUMBER"] === undefined) return false;

    const hasRegional =
        findRegionalColumnIndex(row) !== undefined ||
        colIndex.REGIONAL !== undefined;
    const hasPom = colIndex.POM !== undefined;
    const hasCcBefore = colIndex["COST CENTER BEFORE"] !== undefined;

    return hasRegional || hasPom || hasCcBefore;
}

function countNonEmptyHeaderCells(row) {
    let n = 0;
    for (let c = 0; c < row.length; c++) {
        if (hasCellText(row[c])) n += 1;
    }
    return n;
}

/**
 * Deteksi baris header sheet LPP (lewati baris kosong; dukung FUNCTIONAL LOCATION AFTER2).
 * @returns {{ headerRowIndex: number, dataStartExcelRow: number }}
 */
function findLppGatheringLayout(rows) {
    const limit = Math.min(rows.length, 30);

    for (let i = 0; i < limit; i++) {
        if (rowHasFunctionalLocationAfterHeader(rows[i])) {
            return { headerRowIndex: i, dataStartExcelRow: i + 2 };
        }
    }

    for (let i = 0; i < limit; i++) {
        if (rowLooksLikeLppGatheringHeader(rows[i])) {
            return { headerRowIndex: i, dataStartExcelRow: i + 2 };
        }
    }

    let bestIdx = -1;
    let bestScore = 0;
    for (let i = 0; i < limit; i++) {
        const row = rows[i];
        if (!rowLooksLikeLppGatheringHeader(row)) continue;
        const score = countNonEmptyHeaderCells(row);
        if (score > bestScore) {
            bestScore = score;
            bestIdx = i;
        }
    }
    if (bestIdx >= 0) {
        return { headerRowIndex: bestIdx, dataStartExcelRow: bestIdx + 2 };
    }

    return { headerRowIndex: 0, dataStartExcelRow: 2 };
}

module.exports = {
    OUTPUT_HEADERS,
    PROTECT_Y_APPLY_HEADERS,
    FUNCTLOC_DESC_AFTER_LEVEL123_HEADER,
    FUNCTLOC_DESC_AFTER_APPLIED_HEADER,
    isFunclocDescAfterHeader,
    isFunclocDescAfterLevel123Header,
    isRegionalHeader,
    isCompanyCodeHeader,
    isMaintenancePlantHeader,
    isMeasurmntRangeUnitHeader,
    findFunclocDescAfterColumnIndex,
    findFunclocDescAfterLevel123ColumnIndex,
    findRegionalColumnIndex,
    findCompanyCodeColumnIndex,
    findMaintenancePlantColumnIndex,
    findMeasurmntRangeUnitColumnIndex,
    isPlainCostCenterHeader,
    findFunclocDescBeforeColumnIndex,
    findCostCenterColumnIndex,
    findCostCenterAfterColumnIndex,
    countFunclocAfterKeys,
    dedupeAppliedCandidatesByFuncloc,
    resolveProtectYApplyColumnIndex,
    resolveSourceColumnIndices,
    listMissingSourceColumns,
    extractGatherRows,
    extractAppliedMatchCandidates,
    appendAppliedColumn,
    normalizeCostCenter,
    resolveFileRegionalPom,
    normalizeOutputRowsRegionalPom,
    isFunctionalLocationAfterHeader,
    isFunctionalLocationBeforeHeader,
    findFunctionalLocationBeforeColumnIndex,
    rowLooksLikeLppGatheringHeader,
    findLppGatheringLayout
};
