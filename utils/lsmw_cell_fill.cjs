/**
 * Deteksi background merah (Excel fill, xlsx styles, Google Sheets).
 */

function normalizeHeader(name) {
    if (!name) return "";
    return String(name).trim().toUpperCase();
}

function argbFromColor(color) {
    if (!color) return null;
    if (color.argb) return String(color.argb).replace(/^#/, "").toUpperCase();
    return null;
}

function rgbFromArgb(argb) {
    const s = String(argb ?? "").toUpperCase();
    const hex = s.length === 8 ? s.slice(2) : s;
    if (hex.length !== 6) return null;
    return {
        r: parseInt(hex.slice(0, 2), 16),
        g: parseInt(hex.slice(2, 4), 16),
        b: parseInt(hex.slice(4, 6), 16)
    };
}

/** True jika RGB terlihat merah (termasuk merah terang / pink validasi). */
function isRedRgb(r, g, b) {
    if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return false;
    if (r >= 250 && g >= 250 && b >= 250) return false;
    if (r >= 200 && g <= 195 && b <= 195) return true;
    if (r >= 120 && r > g + 25 && r > b + 25) return true;
    return false;
}

function isRedArgb(argb) {
    if (!argb) return false;
    const u = String(argb).toUpperCase();
    if (u === "FFFF0000" || u === "FF0000") return true;
    if (u === "FFFFC7CE" || u === "FFC7CE") return true;
    if (u === "FFF8CBAD" || u === "F8CBAD") return true;
    if (u === "FFE6B8B7" || u === "E6B8B7") return true;
    const rgb = rgbFromArgb(u);
    return rgb ? isRedRgb(rgb.r, rgb.g, rgb.b) : false;
}

/** rgb string dari xlsx Styles (mis. FFFFC7CE atau FF0000). */
function isRedXlsxRgbString(rgb) {
    if (!rgb) return false;
    return isRedArgb(String(rgb).replace("#", ""));
}

/** Google Sheets backgroundColor { red, green, blue } (0–1). */
function isRedGoogleBackground(bg) {
    if (!bg || typeof bg !== "object") return false;
    const r = bg.red ?? 0;
    const g = bg.green ?? 0;
    const b = bg.blue ?? 0;
    if (r >= 0.97 && g >= 0.97 && b >= 0.97) return false;
    if (r >= 0.72 && g <= 0.82 && b <= 0.82) return true;
    if (r >= 0.45 && r > g + 0.08 && r > b + 0.08) return true;
    return false;
}

/**
 * @param {import('exceljs').Cell | undefined} cell
 */
function isCellRedFill(cell) {
    if (!cell || !cell.fill) return false;
    const fill = cell.fill;

    if (fill.type === "pattern" || !fill.type) {
        if (isRedArgb(argbFromColor(fill.fgColor))) return true;
        if (isRedArgb(argbFromColor(fill.bgColor))) return true;
    }

    return false;
}

/** Index kolom 0-based → huruf Excel (A, B, …, Z, AA, …). */
function columnIndexToLetter(colIndex) {
    let n = colIndex + 1;
    let s = "";
    while (n > 0) {
        const rem = (n - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

/**
 * Cari baris header & kolom EQUIPMENT GROUP AFTER dari sheet rows (header: 1).
 * @returns {{ headerRowIndex: number, dataStartExcelRow: number, idxEquipGroupAfter: number|undefined }}
 */
function findDataLayout(rows) {
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const row = rows[i] || [];
        let idxFuncLoc;
        let idxEquipGroup;
        for (let c = 0; c < row.length; c++) {
            const k = normalizeHeader(row[c]);
            if (k === "FUNCTIONAL LOCATION AFTER") idxFuncLoc = c;
            if (k === "EQUIPMENT GROUP AFTER") idxEquipGroup = c;
            if (
                idxEquipGroup === undefined &&
                k.includes("EQUIPMENT GROUP") &&
                k.includes("AFTER") &&
                !k.includes("BEFORE")
            ) {
                idxEquipGroup = c;
            }
        }
        if (idxFuncLoc !== undefined) {
            return {
                headerRowIndex: i,
                dataStartExcelRow: i + 2,
                idxEquipGroupAfter: idxEquipGroup
            };
        }
    }
    return {
        headerRowIndex: 1,
        dataStartExcelRow: 3,
        idxEquipGroupAfter: undefined
    };
}

module.exports = {
    isCellRedFill,
    isRedGoogleBackground,
    isRedRgb,
    isRedXlsxRgbString,
    isRedArgb,
    columnIndexToLetter,
    normalizeHeader,
    findDataLayout
};
