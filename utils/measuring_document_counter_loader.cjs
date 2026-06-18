/**
 * Index Counter Reading dari output lsmw-measuring-document (by Equipment Number).
 */

const { formatHeriCounterOutput } = require("./heri_sheet_loader.cjs");

const HEADER_ROW_INDEX = 0;
const DATA_START_ROW_INDEX = 1;

const SKIP_FILE_NAMES = new Set([
    "validation.xlsx",
    "all-pks.xlsx",
    "missing.xlsx"
]);

const SKIP_FOLDER_NAMES = new Set(["Missing Measuring Document"]);

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

/**
 * @param {string} fileName
 */
function shouldSkipMeasuringDocAggregateFile(fileName) {
    const name = cellText(fileName);
    if (!name || name.startsWith("~$")) return true;
    if (SKIP_FILE_NAMES.has(name.toLowerCase())) return true;
    if (/^REGIONAL\s+\d+\.xlsx$/i.test(name)) return true;
    return false;
}

/**
 * @param {Array<unknown>} headerRow
 * @param {string} prefix
 */
function findColumnIndexByPrefix(headerRow, prefix) {
    const upper = cellText(prefix).toUpperCase();
    if (!upper) return -1;

    for (let i = 0; i < headerRow.length; i += 1) {
        const h = cellText(headerRow[i]).toUpperCase();
        if (!h) continue;
        if (h.startsWith(upper)) return i;
    }
    for (let i = 0; i < headerRow.length; i += 1) {
        const h = cellText(headerRow[i]).toUpperCase();
        if (!h) continue;
        if (h.includes(upper)) return i;
    }
    return -1;
}

/**
 * @param {unknown} raw
 */
function counterCellToDisplay(raw) {
    if (raw === null || raw === undefined || raw === "") return "0";
    if (typeof raw === "number" && Number.isFinite(raw)) {
        return formatHeriCounterOutput(raw) ?? "0";
    }
    const text = cellText(raw);
    return text || "0";
}

/**
 * @param {unknown} value
 */
function normalizeCounterDisplayForCompare(value) {
    return counterCellToDisplay(value);
}

/**
 * @param {Array<Array<unknown>>} rawRows
 * @param {Map<string, string>} counterByEqunr
 * @returns {number}
 */
function mergeCounterRowsIntoIndex(rawRows, counterByEqunr) {
    if (!rawRows.length) return 0;

    const headerRow = rawRows[HEADER_ROW_INDEX] || [];
    const idxEqunr = findColumnIndexByPrefix(headerRow, "Equipment Number");
    const idxCounter = findColumnIndexByPrefix(headerRow, "Counter Reading");

    if (idxEqunr < 0 || idxCounter < 0) return 0;

    let merged = 0;
    for (let r = DATA_START_ROW_INDEX; r < rawRows.length; r += 1) {
        const row = rawRows[r] || [];
        const equnr = cellText(row[idxEqunr]);
        if (!equnr) continue;
        if (counterByEqunr.has(equnr)) continue;

        counterByEqunr.set(equnr, counterCellToDisplay(row[idxCounter]));
        merged += 1;
    }
    return merged;
}

/**
 * @param {{
 *   drive: import("googleapis").drive_v3.Drive,
 *   rootFolderId: string,
 *   isProcessableSpreadsheet: (file: object) => boolean,
 *   listFolderChildren: (folderId: string) => Promise<object[]>,
 *   downloadSpreadsheetToTemp: (file: object) => Promise<string>,
 *   readSpreadsheetRows: (localFile: string) => Promise<{ rawRows: Array<Array<unknown>> }>
 * }} deps
 */
async function loadMeasuringDocumentCounterFromDrive(deps) {
    const {
        drive,
        rootFolderId,
        isProcessableSpreadsheet,
        listFolderChildren,
        downloadSpreadsheetToTemp,
        readSpreadsheetRows
    } = deps;

    /** @type {Array<object>} */
    const files = [];

    async function walk(folderId) {
        const children = await listFolderChildren(folderId);
        for (const child of children) {
            if (child.mimeType === "application/vnd.google-apps.folder") {
                if (SKIP_FOLDER_NAMES.has(cellText(child.name))) continue;
                await walk(child.id);
            } else if (
                isProcessableSpreadsheet(child) &&
                !shouldSkipMeasuringDocAggregateFile(child.name)
            ) {
                files.push(child);
            }
        }
    }

    await walk(rootFolderId);

    const counterByEqunr = new Map();
    let filesRead = 0;

    for (const file of files) {
        const localFile = await downloadSpreadsheetToTemp(file);
        const sheet = await readSpreadsheetRows(localFile);
        const added = mergeCounterRowsIntoIndex(sheet.rawRows, counterByEqunr);
        if (added > 0) filesRead += 1;
    }

    return {
        counterByEqunr,
        filesScanned: files.length,
        filesWithData: filesRead,
        duplicateEquipment: files.length > 0 ? files.length - filesRead : 0
    };
}

module.exports = {
    shouldSkipMeasuringDocAggregateFile,
    findColumnIndexByPrefix,
    counterCellToDisplay,
    normalizeCounterDisplayForCompare,
    mergeCounterRowsIntoIndex,
    loadMeasuringDocumentCounterFromDrive
};
