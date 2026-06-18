/**
 * Greedy 1-to-1 matching: baris LPP ↔ baris Applied.
 * Prioritas: COST CENTER cocok (AFTER atau BEFORE tergantung mode baris), lalu similarity FUNCTLOC DESC.
 * Baris LPP dengan FUNCTLOC DESC AFTER terisi: CC AFTER + FUNCTLOC AFTER.
 * Baris LPP tanpa AFTER: CC BEFORE + FUNCTLOC BEFORE vs Applied AFTER.
 * Jika FUNCTLOC AFTER & BEFORE kosong: CC BEFORE + EQKTU BEFORE vs Applied AFTER.
 * Fallback: baris LPP yang belum terpasang wajib dapat Applied paling mirip.
 */
const { similarityFunclocDesc } = require("./string_similarity.cjs");
const { normalizeCostCenter } = require("./equipment_gathering_columns.cjs");

function normalizeDesc(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function hasDesc(value) {
    return normalizeDesc(value) !== "";
}

function costCentersMatch(lppCostCenter, appliedCostCenterAfter) {
    const lpp = normalizeCostCenter(lppCostCenter);
    const applied = normalizeCostCenter(appliedCostCenterAfter);
    if (!lpp || !applied) return false;
    return lpp === applied;
}

function resolveLppMatchFields(
    row,
    lppFunclocAfterCol,
    lppFunclocBeforeCol,
    lppCostCenterAfterCol,
    lppCostCenterBeforeCol,
    lppEqktuBeforeCol
) {
    if (hasDesc(row[lppFunclocAfterCol])) {
        return {
            desc: row[lppFunclocAfterCol],
            costCenter: row[lppCostCenterAfterCol]
        };
    }
    if (hasDesc(row[lppFunclocBeforeCol])) {
        return {
            desc: row[lppFunclocBeforeCol],
            costCenter: row[lppCostCenterBeforeCol]
        };
    }
    if (hasDesc(row[lppEqktuBeforeCol])) {
        return {
            desc: row[lppEqktuBeforeCol],
            costCenter: row[lppCostCenterBeforeCol]
        };
    }
    return null;
}

function greedyAssign(pairs, appliedCandidates, matches, assignedLpp, assignedApplied) {
    for (const { lppIdx, appliedIdx } of pairs) {
        if (assignedLpp.has(lppIdx) || assignedApplied.has(appliedIdx)) continue;
        assignedLpp.add(lppIdx);
        assignedApplied.add(appliedIdx);
        matches.set(
            lppIdx,
            appliedCandidates[appliedIdx].funclocDescAfter ?? ""
        );
    }
}

function buildPairs(
    lppRows,
    appliedCandidates,
    lppFunclocAfterCol,
    lppFunclocBeforeCol,
    lppCostCenterAfterCol,
    lppCostCenterBeforeCol,
    lppEqktuBeforeCol
) {
    /** @type {Array<{ lppIdx: number, appliedIdx: number, score: number, ccMatch: number }>} */
    const pairs = [];

    for (let lppIdx = 0; lppIdx < lppRows.length; lppIdx++) {
        const fields = resolveLppMatchFields(
            lppRows[lppIdx],
            lppFunclocAfterCol,
            lppFunclocBeforeCol,
            lppCostCenterAfterCol,
            lppCostCenterBeforeCol,
            lppEqktuBeforeCol
        );
        if (!fields) continue;

        for (let appliedIdx = 0; appliedIdx < appliedCandidates.length; appliedIdx++) {
            const candidate = appliedCandidates[appliedIdx];
            const appliedDesc = candidate.funclocDescAfter;
            if (!hasDesc(appliedDesc)) continue;

            pairs.push({
                lppIdx,
                appliedIdx,
                score: similarityFunclocDesc(fields.desc, appliedDesc),
                ccMatch: costCentersMatch(fields.costCenter, candidate.costCenterAfter)
                    ? 1
                    : 0
            });
        }
    }

    return pairs;
}

function sortPairs(pairs) {
    pairs.sort(
        (a, b) =>
            b.ccMatch - a.ccMatch ||
            b.score - a.score ||
            a.appliedIdx - b.appliedIdx ||
            a.lppIdx - b.lppIdx
    );
}

/**
 * @param {Array<Array<unknown>>} lppRows
 * @param {Array<{ costCenterAfter: string, funclocDescAfter: string }>} appliedCandidates
 * @param {number} lppCostCenterAfterCol
 * @param {number} lppFunclocAfterCol
 * @param {number} lppCostCenterBeforeCol
 * @param {number} lppFunclocBeforeCol
 * @param {number} lppEqktuBeforeCol
 * @returns {Map<number, string>} lppRowIndex → nilai FUNCTLOC DESC AFTER Applied
 */
function matchLppRowsToApplied(
    lppRows,
    appliedCandidates,
    lppCostCenterAfterCol,
    lppFunclocAfterCol,
    lppCostCenterBeforeCol,
    lppFunclocBeforeCol,
    lppEqktuBeforeCol
) {
    if (appliedCandidates.length === 0) {
        return new Map();
    }

    const pairs = buildPairs(
        lppRows,
        appliedCandidates,
        lppFunclocAfterCol,
        lppFunclocBeforeCol,
        lppCostCenterAfterCol,
        lppCostCenterBeforeCol,
        lppEqktuBeforeCol
    );
    sortPairs(pairs);

    const assignedLpp = new Set();
    const assignedApplied = new Set();
    /** @type {Map<number, string>} */
    const matches = new Map();

    greedyAssign(
        pairs.filter(p => p.ccMatch === 1),
        appliedCandidates,
        matches,
        assignedLpp,
        assignedApplied
    );

    const fallbackPairs = pairs.filter(
        p => !assignedLpp.has(p.lppIdx) && !assignedApplied.has(p.appliedIdx)
    );
    sortPairs(fallbackPairs);
    greedyAssign(
        fallbackPairs,
        appliedCandidates,
        matches,
        assignedLpp,
        assignedApplied
    );

    for (let lppIdx = 0; lppIdx < lppRows.length; lppIdx++) {
        if (assignedLpp.has(lppIdx)) continue;

        const fields = resolveLppMatchFields(
            lppRows[lppIdx],
            lppFunclocAfterCol,
            lppFunclocBeforeCol,
            lppCostCenterAfterCol,
            lppCostCenterBeforeCol,
            lppEqktuBeforeCol
        );
        if (!fields) continue;

        let bestIdx = -1;
        let bestScore = -1;
        for (let appliedIdx = 0; appliedIdx < appliedCandidates.length; appliedIdx++) {
            if (assignedApplied.has(appliedIdx)) continue;
            const appliedDesc = appliedCandidates[appliedIdx].funclocDescAfter;
            if (!hasDesc(appliedDesc)) continue;
            const score = similarityFunclocDesc(fields.desc, appliedDesc);
            if (score > bestScore) {
                bestScore = score;
                bestIdx = appliedIdx;
            }
        }

        if (bestIdx >= 0) {
            assignedLpp.add(lppIdx);
            assignedApplied.add(bestIdx);
            matches.set(
                lppIdx,
                appliedCandidates[bestIdx].funclocDescAfter ?? ""
            );
        }
    }

    return matches;
}

module.exports = { matchLppRowsToApplied, costCentersMatch, resolveLppMatchFields };
