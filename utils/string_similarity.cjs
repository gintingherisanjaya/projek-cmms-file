/**
 * Similarity score for cell text (Levenshtein-based, 0–1).
 * Ported from utils/excelHelpers.js for CommonJS scripts.
 */

function normalizeCellText(value) {
    return String(value ?? "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();
}

/** Normalisasi FUNCTLOC DESC: spasi, NO.3 vs NO. 3, tanda baca. */
function normalizeFunclocDesc(value) {
    return String(value ?? "")
        .trim()
        .replace(/[`'"\u2018\u2019\u201c\u201d]/g, "")
        .replace(/\bNO\.\s*/gi, "NO ")
        .replace(/[.,;:]+/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function levenshtein(a, b) {
    const m = a.length;
    const n = b.length;
    const d = Array(m + 1)
        .fill(null)
        .map(() => Array(n + 1).fill(0));
    for (let i = 0; i <= m; i += 1) d[i][0] = i;
    for (let j = 0; j <= n; j += 1) d[0][j] = j;
    for (let i = 1; i <= m; i += 1) {
        for (let j = 1; j <= n; j += 1) {
            const cost = a[i - 1] === b[j - 1] ? 0 : 1;
            d[i][j] = Math.min(
                d[i - 1][j] + 1,
                d[i][j - 1] + 1,
                d[i - 1][j - 1] + cost
            );
        }
    }
    return d[m][n];
}

/** Similarity score 0–1 (1 = identical). */
function similarity(a, b) {
    const an = normalizeCellText(a);
    const bn = normalizeCellText(b);
    if (!an || !bn) return 0;
    const d = levenshtein(an, bn);
    const maxLen = Math.max(an.length, bn.length);
    return maxLen === 0 ? 1 : 1 - d / maxLen;
}

/** Similarity khusus FUNCTLOC DESC (normalisasi NO./spasi, deskripsi gabungan LPP). */
function similarityFunclocDesc(a, b) {
    const an = normalizeFunclocDesc(a);
    const bn = normalizeFunclocDesc(b);
    if (!an || !bn) return 0;
    if (an === bn) return 1;

    let containScore = 0;
    if (an.includes(bn) || bn.includes(an)) {
        const shorter = an.length <= bn.length ? an : bn;
        const longer = an.length <= bn.length ? bn : an;
        const ratio = shorter.length / longer.length;
        containScore = 0.85 + 0.15 * ratio;
        if (longer.indexOf(shorter) === 0) {
            containScore = Math.min(1, containScore + 0.05);
        }
    }

    const d = levenshtein(an, bn);
    const maxLen = Math.max(an.length, bn.length);
    const levScore = maxLen === 0 ? 1 : 1 - d / maxLen;

    return Math.max(containScore, levScore);
}

module.exports = {
    normalizeCellText,
    normalizeFunclocDesc,
    levenshtein,
    similarity,
    similarityFunclocDesc
};
