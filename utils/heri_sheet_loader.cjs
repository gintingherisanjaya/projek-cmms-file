/**
 * Baca & parse sheet HERI dari Google Spreadsheet (Data Jam Jalan Peralatan).
 */

const DEFAULT_HERI_SPREADSHEET_ID = "1AAlE1ObboKTF-iuMYx6xWcMUhAEPMgGlEbFCOxSF94Y";
const HERI_SHEET_NAME = "HERI";

/** Baris 2 (1-based) = nama alat; kolom D (index 3) ke kanan. */
const EQUIPMENT_HEADER_ROW_INDEX = 1;
const EQUIPMENT_START_COL_INDEX = 3;

/** Kolom C (index 2) = plant; data mulai baris 4 (index 3). */
const PLANT_COL_INDEX = 2;
const PLANT_DATA_START_ROW_INDEX = 3;

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function normalizePlant(code) {
    return cellText(code).toUpperCase();
}

function isValidPlantCode(code) {
    const s = normalizePlant(code);
    return /^[A-Z0-9]{4}$/.test(s);
}

/**
 * Normalisasi notasi angka HERI: titik = pemisah ribuan (dihapus), koma = desimal.
 * @param {string} text
 * @returns {string}
 */
function normalizeHeriCounterDigits(text) {
    let s = String(text).replace(/\s/g, "");
    const negative = s.startsWith("-");
    if (negative) s = s.slice(1);

    if (s.includes(",")) {
        s = s.replace(/\./g, "").replace(",", ".");
    } else if (s.includes(".")) {
        s = s.replace(/\./g, "");
    }

    if (negative) s = `-${s}`;
    return s;
}

/**
 * @param {unknown} raw
 * @returns {{ ok: true, value: number, raw: string } | { ok: false, raw: string }}
 */
function parseHeriCounter(raw) {
    const text = cellText(raw);
    if (!text) return { ok: false, raw: text };

    const upper = text.toUpperCase();
    if (upper === "HM RUSAK") {
        return { ok: true, value: 0, raw: text };
    }
    if (
        upper === "NO EQUIPPMENT" ||
        upper === "NO EQUIPMENT" ||
        upper === "-"
    ) {
        return { ok: false, raw: text };
    }

    const normalized = normalizeHeriCounterDigits(text);
    const num = Number(normalized);
    if (!Number.isFinite(num)) {
        return { ok: false, raw: text };
    }

    return { ok: true, value: num, raw: text };
}

/**
 * Format nilai counter HERI untuk output Excel (koma = desimal).
 * @param {number | null | undefined} value
 * @returns {string | null}
 */
function formatHeriCounterOutput(value) {
    if (value === null || value === undefined) return null;
    if (!Number.isFinite(value)) return null;
    return String(value).replace(".", ",");
}

/**
 * @param {Array<Array<unknown>>} rows
 * @returns {{
 *   equipmentNames: string[],
 *   equipmentNamesLongestFirst: string[],
 *   byPlant: Map<string, Map<string, { raw: string, numeric: number | null, excelRow: number, colIndex: number }>>
 * }}
 */
function parseHeriRows(rows) {
    const headerRow = rows[EQUIPMENT_HEADER_ROW_INDEX] || [];
    const equipmentCols = [];

    for (let c = EQUIPMENT_START_COL_INDEX; c < headerRow.length; c += 1) {
        const name = cellText(headerRow[c]);
        if (!name) {
            if (equipmentCols.length > 0 && c > equipmentCols.length + 5) break;
            continue;
        }
        equipmentCols.push({ name, colIndex: c });
    }

    if (equipmentCols.length === 0) {
        throw new Error(
            "Sheet HERI: tidak ada nama alat di baris 2 mulai kolom D"
        );
    }

    const byPlant = new Map();

    for (let r = PLANT_DATA_START_ROW_INDEX; r < rows.length; r += 1) {
        const row = rows[r] || [];
        const plant = normalizePlant(row[PLANT_COL_INDEX]);
        if (!plant) continue;
        if (!isValidPlantCode(plant)) continue;

        if (!byPlant.has(plant)) {
            byPlant.set(plant, new Map());
        }
        const byEquipment = byPlant.get(plant);

        for (const { name, colIndex } of equipmentCols) {
            const raw = cellText(row[colIndex]);
            const parsed = parseHeriCounter(raw);
            byEquipment.set(name, {
                raw,
                numeric: parsed.ok ? parsed.value : null,
                excelRow: r + 1,
                colIndex: colIndex + 1
            });
        }
    }

    const equipmentNames = equipmentCols.map(c => c.name);
    const equipmentNamesLongestFirst = [...equipmentNames].sort(
        (a, b) => b.length - a.length
    );

    return {
        equipmentNames,
        equipmentNamesLongestFirst,
        byPlant
    };
}

/**
 * @param {import("googleapis").sheets_v4.Sheets} sheetsApi
 * @param {string} [spreadsheetId]
 */
async function loadHeriSheet(sheetsApi, spreadsheetId = DEFAULT_HERI_SPREADSHEET_ID) {
    const res = await sheetsApi.spreadsheets.values.get({
        spreadsheetId,
        range: `${HERI_SHEET_NAME}!A:ZZ`
    });

    const values = res.data.values;
    if (!values || values.length === 0) {
        throw new Error(`Sheet HERI kosong pada spreadsheet ${spreadsheetId}`);
    }

    const parsed = parseHeriRows(values);
    return {
        spreadsheetId,
        sheetName: HERI_SHEET_NAME,
        ...parsed
    };
}

/**
 * @param {ReturnType<typeof parseHeriRows>} heriData
 * @param {string} plant
 * @param {string} equipmentName
 */
function getHeriCell(heriData, plant, equipmentName) {
    const p = normalizePlant(plant);
    const byEquipment = heriData.byPlant.get(p);
    if (!byEquipment) return null;
    return byEquipment.get(equipmentName) ?? null;
}

const STERILIZER_ALIAS_RE = /^(HORIZONTAL|VERTICAL)\s+STERILIZER\s+NO\.?\s*(\d+)\s*$/i;

/**
 * @param {string} desc
 * @returns {string | null}
 */
function canonicalSterilizerHeriName(desc) {
    const m = cellText(desc).match(STERILIZER_ALIAS_RE);
    if (!m) return null;
    return `STERILIZER NO. ${m[2]}`;
}

/**
 * @param {string} desc
 * @param {string} heriEquipmentName
 * @returns {boolean}
 */
function isSterilizerHeriAlias(desc, heriEquipmentName) {
    const canonical = canonicalSterilizerHeriName(desc);
    if (!canonical) return false;
    return cellText(heriEquipmentName).toUpperCase() === canonical.toUpperCase();
}

/**
 * @param {string} desc
 * @param {string[]} equipmentNamesLongestFirst
 * @param {{ allowSterilizerAlias?: boolean }} [options]
 * @returns {string | null}
 */
function matchHeriEquipmentName(desc, equipmentNamesLongestFirst, options = {}) {
    const upper = cellText(desc).toUpperCase();
    if (!upper) return null;

    for (const name of equipmentNamesLongestFirst) {
        const key = cellText(name).toUpperCase();
        if (!key) continue;
        if (upper === key) {
            return name;
        }
    }

    if (!options.allowSterilizerAlias) return null;

    const canonical = canonicalSterilizerHeriName(desc);
    if (!canonical) return null;

    const canonicalUpper = canonical.toUpperCase();
    for (const name of equipmentNamesLongestFirst) {
        const key = cellText(name).toUpperCase();
        if (!key) continue;
        if (key === canonicalUpper) {
            return name;
        }
    }
    return null;
}

/**
 * @param {ReturnType<typeof parseHeriRows>} heriData
 * @param {string} plant
 * @param {string} equipmentName
 * @returns {{
 *   counterValue?: number | null,
 *   counterIssue?: string,
 *   skipAnchor?: boolean,
 *   skipRow?: boolean,
 *   skipReason?: string
 * }}
 */
function resolveHeriCounterForEquipment(heriData, plant, equipmentName) {
    const heriCell = getHeriCell(heriData, plant, equipmentName);
    if (!heriCell) {
        return {
            counterValue: null,
            counterIssue: "plant tidak ditemukan di HERI"
        };
    }

    const parsed = parseHeriCounter(heriCell.raw);
    if (!parsed.ok) {
        return {
            skipAnchor: true,
            skipRow: true,
            skipReason: `counter HERI non-numerik "${heriCell.raw}"`
        };
    }
    if (parsed.value === 0) {
        return {
            skipAnchor: true,
            skipRow: true,
            skipReason: "counter HERI bernilai 0"
        };
    }

    return { counterValue: parsed.value };
}

/**
 * @param {ReturnType<typeof parseHeriRows> & { equipmentNamesLongestFirst: string[] }} heriData
 * @param {string} plant
 * @param {string} desc
 * @param {{ allowSterilizerAlias?: boolean }} [options]
 */
function resolveHeriCounterForDesc(heriData, plant, desc, options = {}) {
    const equipmentName = matchHeriEquipmentName(
        desc,
        heriData.equipmentNamesLongestFirst,
        options
    );
    if (!equipmentName) {
        return {
            skipRow: true,
            skipAnchor: true,
            skipReason: `tidak ada kolom alat HERI yang cocok untuk "${cellText(desc)}"`
        };
    }

    return resolveHeriCounterForEquipment(heriData, plant, equipmentName);
}

/**
 * HERI untuk maintenance-plan: tidak pernah skip baris; invalid/tidak ada → counter 0.
 * @param {ReturnType<typeof parseHeriRows> & { equipmentNamesLongestFirst: string[] }} heriData
 * @param {string} plant
 * @param {string} desc
 * @param {{ allowSterilizerAlias?: boolean }} [options]
 */
function resolveHeriCounterForMaintenancePlan(heriData, plant, desc, options = {}) {
    if (!heriData) {
        return { counterValue: 0 };
    }

    const equipmentName = matchHeriEquipmentName(
        desc,
        heriData.equipmentNamesLongestFirst,
        options
    );
    if (!equipmentName) {
        return { counterValue: 0 };
    }

    const heriCell = getHeriCell(heriData, plant, equipmentName);
    if (!heriCell) {
        return { counterValue: 0 };
    }

    const parsed = parseHeriCounter(heriCell.raw);
    if (!parsed.ok) {
        return { counterValue: 0 };
    }

    return { counterValue: parsed.value };
}

/**
 * @param {ReturnType<typeof parseHeriRows> & { equipmentNamesLongestFirst: string[] }} heriData
 * @param {string} plant
 * @param {string} desc
 */
function resolveHeriCounter(heriData, plant, desc) {
    const equipmentName = matchHeriEquipmentName(
        desc,
        heriData.equipmentNamesLongestFirst
    );
    if (!equipmentName) {
        return { ok: false, reason: "no_heri_match", raw: "" };
    }

    const cell = getHeriCell(heriData, plant, equipmentName);
    if (!cell) {
        return {
            ok: false,
            reason: "no_plant",
            raw: "",
            equipmentName
        };
    }

    const parsed = parseHeriCounter(cell.raw);
    return {
        ...parsed,
        equipmentName,
        excelRow: cell.excelRow,
        colIndex: cell.colIndex
    };
}

module.exports = {
    DEFAULT_HERI_SPREADSHEET_ID,
    HERI_SHEET_NAME,
    normalizeHeriCounterDigits,
    parseHeriCounter,
    formatHeriCounterOutput,
    parseHeriRows,
    loadHeriSheet,
    getHeriCell,
    canonicalSterilizerHeriName,
    isSterilizerHeriAlias,
    matchHeriEquipmentName,
    resolveHeriCounterForEquipment,
    resolveHeriCounterForDesc,
    resolveHeriCounterForMaintenancePlan,
    resolveHeriCounter,
    normalizePlant
};
