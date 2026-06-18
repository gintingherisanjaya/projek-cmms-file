/**
 * Hitung baris data nyata (template-agnostik) dari sheet rows 2D.
 */

const { findDataLayout, normalizeHeader } = require("./lsmw_cell_fill.cjs");

const HEADER_SCAN_ROWS = 30;
const KNOWN_HEADERS = [
    "FUNCTIONAL LOCATION AFTER",
    "TPLNR",
    "STRNO",
    "EQUNR"
];
const KEY_COLUMN_ORDER = [
    "FUNCTIONAL LOCATION AFTER",
    "TPLNR",
    "STRNO",
    "EQUNR"
];
const LSMW_HEADER_MARKERS = new Set([
    "TPLNR",
    "STRNO",
    "SHTXT*",
    "PLTXT",
    "EQART",
    "EQUNR"
]);

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function rowNonEmptyCount(row) {
    let n = 0;
    for (const cell of row || []) {
        if (cellText(cell)) n += 1;
    }
    return n;
}

function buildColumnIndex(headerRow) {
    const colIndex = {};
    for (let c = 0; c < (headerRow || []).length; c++) {
        const key = normalizeHeader(headerRow[c]);
        if (key && colIndex[key] === undefined) colIndex[key] = c;
    }
    return colIndex;
}

function scoreHeaderRow(row) {
    const colIndex = buildColumnIndex(row);
    let score = 0;
    for (const name of KNOWN_HEADERS) {
        if (colIndex[name] !== undefined) score += 1;
    }
    if (score === 0) return 0;

    let lsmwMarkers = 0;
    for (const key of Object.keys(colIndex)) {
        if (LSMW_HEADER_MARKERS.has(key)) lsmwMarkers += 1;
    }
    if (colIndex.EQUNR !== undefined && lsmwMarkers < 2) {
        score -= 0.5;
    }
    return score;
}

/**
 * @param {Array<Array<unknown>>} rows
 * @returns {number}
 */
function findHeaderRowIndex(rows) {
    const layout = findDataLayout(rows);
    const layoutRow = rows[layout.headerRowIndex] || [];
    if (
        buildColumnIndex(layoutRow)["FUNCTIONAL LOCATION AFTER"] !== undefined
    ) {
        return layout.headerRowIndex;
    }

    let bestIdx = 0;
    let bestScore = 0;
    const limit = Math.min(rows.length, HEADER_SCAN_ROWS);

    for (let i = 0; i < limit; i++) {
        const score = scoreHeaderRow(rows[i]);
        if (score > bestScore) {
            bestScore = score;
            bestIdx = i;
        }
    }

    if (bestScore > 0) return bestIdx;

    let maxCells = 0;
    let maxIdx = 0;
    for (let i = 0; i < limit; i++) {
        const n = rowNonEmptyCount(rows[i]);
        if (n >= 3 && n > maxCells) {
            maxCells = n;
            maxIdx = i;
        }
    }
    return maxIdx;
}

/**
 * @param {Array<Array<unknown>>} rows
 * @returns {{ count: number, headerRowIndex: number, keyColumn: string|null }}
 */
function countRealDataRows(rows) {
    if (!rows || rows.length === 0) {
        return { count: 0, headerRowIndex: 0, keyColumn: null };
    }

    const headerRowIndex = findHeaderRowIndex(rows);
    const headerRow = rows[headerRowIndex] || [];
    const colIndex = buildColumnIndex(headerRow);

    let keyColumn = null;
    let keyIdx;
    for (const name of KEY_COLUMN_ORDER) {
        if (colIndex[name] !== undefined) {
            keyColumn = name;
            keyIdx = colIndex[name];
            break;
        }
    }

    let count = 0;
    const namedCols = [];
    for (let c = 0; c < headerRow.length; c++) {
        if (normalizeHeader(headerRow[c])) namedCols.push(c);
    }

    for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const row = rows[r] || [];

        if (keyIdx !== undefined) {
            if (cellText(row[keyIdx])) count += 1;
            continue;
        }

        let any = false;
        const colsToCheck = namedCols.length > 0 ? namedCols : row.map((_, i) => i);
        for (const c of colsToCheck) {
            if (cellText(row[c])) {
                any = true;
                break;
            }
        }
        if (any) count += 1;
    }

    return { count, headerRowIndex, keyColumn };
}

module.exports = {
    countRealDataRows,
    findHeaderRowIndex,
    cellText
};

if (require.main === module) {
    const protectYHeader = [
        "REGIONAL",
        "FUNCTIONAL LOCATION AFTER",
        "FUNCTLOC DESC. AFTER LEVEL 1,2,3"
    ];
    const protectYRows = [
        protectYHeader,
        ["R1", "PALM-1F01-0001", "PUMP NO.1"],
        ["R1", "PALM-1F01-0002", "PUMP NO.2"],
        ["R1", "", ""],
        ["R1", "PALM-1F01-0003", "PUMP NO.3"]
    ];
    const py = countRealDataRows(protectYRows);
    if (py.count !== 3) {
        console.error("FAIL protectY expected 3 got", py.count);
        process.exit(1);
    }

    const lsmwRows = [
        ["", "", ""],
        ["No", "EQART", "TPLNR"],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "ELMO", "PALM-1F01-0001"],
        ["", "ELMO", "PALM-1F01-0002"],
        ["", "", ""]
    ];
    const lsmw = countRealDataRows(lsmwRows);
    if (lsmw.count !== 2 || lsmw.keyColumn !== "TPLNR") {
        console.error("FAIL lsmw expected 2/TPLNR got", lsmw);
        process.exit(1);
    }

    console.log("OK count_real_data_rows smoke");
}
