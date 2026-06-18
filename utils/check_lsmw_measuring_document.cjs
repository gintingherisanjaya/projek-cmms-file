/**
 * Util check measuring document vs CDS (Measuring point + Equipment).
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { readSpreadsheetRows } = require("./equipment_excel_io.cjs");
const { findColumnIndexByPrefix } = require("./measuring_document_counter_loader.cjs");

const HEADER_ROW_INDEX = 0;
const DATA_START_ROW_INDEX = 1;

const CHECK_HEADERS = [
    "nama regional",
    "nama pks",
    "measuring point",
    "equipment",
    "status"
];

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

/**
 * @param {string} measuringPoint
 * @param {string} equipment
 */
function pairKey(measuringPoint, equipment) {
    return `${cellText(measuringPoint)}|${cellText(equipment)}`;
}

/**
 * @param {string} cdsPath
 * @returns {Promise<Set<string>>}
 */
async function loadCdsMeasuringPointEquipmentIndex(cdsPath) {
    const resolved = path.resolve(cdsPath);
    if (!fs.existsSync(resolved)) {
        throw new Error(`File CDS tidak ditemukan: ${resolved}`);
    }

    const sheet = await readSpreadsheetRows(resolved);
    if (!sheet.ok) {
        throw new Error(`Gagal baca CDS: ${sheet.reason}`);
    }

    const headerRow = sheet.rawRows[HEADER_ROW_INDEX] || [];
    const idxPoint = findColumnIndexByPrefix(headerRow, "Measuring point");
    const idxEquipment = findColumnIndexByPrefix(headerRow, "Equipment");

    if (idxPoint < 0 || idxEquipment < 0) {
        throw new Error(
            "Kolom wajib CDS tidak ditemukan: Measuring point, Equipment"
        );
    }

    const index = new Set();
    for (let r = DATA_START_ROW_INDEX; r < sheet.rawRows.length; r += 1) {
        const row = sheet.rawRows[r] || [];
        const point = cellText(row[idxPoint]);
        const equipment = cellText(row[idxEquipment]);
        if (!point || !equipment) continue;
        index.add(pairKey(point, equipment));
    }

    return index;
}

/**
 * @param {Array<Array<unknown>>} rawRows
 * @returns {Array<{ measuringPoint: string, equipment: string }>}
 */
function extractMeasuringDocRows(rawRows) {
    if (!rawRows.length) return [];

    const headerRow = rawRows[HEADER_ROW_INDEX] || [];
    const idxPoint = findColumnIndexByPrefix(headerRow, "Measuring point");
    const idxEquipment = findColumnIndexByPrefix(headerRow, "Equipment Number");

    if (idxPoint < 0 || idxEquipment < 0) {
        return [];
    }

    /** @type {Array<{ measuringPoint: string, equipment: string }>} */
    const pairs = [];

    for (let r = DATA_START_ROW_INDEX; r < rawRows.length; r += 1) {
        const row = rawRows[r] || [];
        const measuringPoint = cellText(row[idxPoint]);
        const equipment = cellText(row[idxEquipment]);
        if (!measuringPoint || !equipment) continue;
        pairs.push({ measuringPoint, equipment });
    }

    return pairs;
}

/**
 * @param {Array<{ measuringPoint: string, equipment: string }>} pairs
 * @returns {Array<{ measuringPoint: string, equipment: string }>}
 */
function dedupePairs(pairs) {
    const seen = new Set();
    /** @type {Array<{ measuringPoint: string, equipment: string }>} */
    const out = [];

    for (const pair of pairs) {
        const key = pairKey(pair.measuringPoint, pair.equipment);
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(pair);
    }

    return out;
}

/**
 * @param {string} regional
 * @param {string} pksName
 * @param {Array<{ measuringPoint: string, equipment: string }>} pairs
 * @param {Set<string>} cdsIndex
 */
function buildCheckRows(regional, pksName, pairs, cdsIndex) {
    return pairs.map(pair => {
        const key = pairKey(pair.measuringPoint, pair.equipment);
        return {
            regional,
            pksName,
            measuringPoint: pair.measuringPoint,
            equipment: pair.equipment,
            status: cdsIndex.has(key) ? "finded" : "missing"
        };
    });
}

/**
 * @param {string} outPath
 * @param {Array<{ regional: string, pksName: string, measuringPoint: string, equipment: string, status: string }>} rows
 */
async function writeCheckMeasuringDocumentExcel(outPath, rows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Check", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(CHECK_HEADERS);
    ws.getRow(1).font = { bold: true };

    for (const row of rows) {
        ws.addRow([
            row.regional,
            row.pksName,
            row.measuringPoint,
            row.equipment,
            row.status
        ]);
    }

    ws.columns = [
        { width: 18 },
        { width: 36 },
        { width: 20 },
        { width: 18 },
        { width: 12 }
    ];

    if (rows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + rows.length, column: CHECK_HEADERS.length }
        };
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await wb.xlsx.writeFile(outPath);
}

module.exports = {
    CHECK_HEADERS,
    pairKey,
    loadCdsMeasuringPointEquipmentIndex,
    extractMeasuringDocRows,
    dedupePairs,
    buildCheckRows,
    writeCheckMeasuringDocumentExcel
};
