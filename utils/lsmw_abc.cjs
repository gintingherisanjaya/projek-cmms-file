/**
 * Lookup ABC Indicators dari utils/abc.csv (pola funcloc dengan wildcard).
 */

const fs = require("fs");
const path = require("path");

const ABC_CSV_PATH = path.join(__dirname, "abc.csv");
const ABC_DEFAULT = "2";

let abcRulesCache = null;

function isEmpty(v) {
    return v === null || v === undefined || String(v).trim() === "";
}

function loadAbcRules() {
    if (abcRulesCache) return abcRulesCache;

    const raw = fs.readFileSync(ABC_CSV_PATH, "utf8");
    const lines = raw.trim().split(/\r?\n/).filter(Boolean);
    if (lines.length < 2) {
        abcRulesCache = [];
        return abcRulesCache;
    }

    const rules = [];
    for (let i = 1; i < lines.length; i++) {
        const parts = lines[i].split(",");
        const abc = (parts[0] || "").trim();
        const pattern = (parts[1] || "").trim();
        if (!abc || !pattern) continue;
        const segments = pattern.split("-").map(s => s.trim().toUpperCase());
        const literalCount = segments.filter(s => !s.includes("*")).length;
        rules.push({ abc, segments, literalCount });
    }

    rules.sort((a, b) => b.literalCount - a.literalCount);
    abcRulesCache = rules;
    return abcRulesCache;
}

function ensureAbcRules() {
    return loadAbcRules();
}

function segmentMatches(patternSeg, funclocSeg) {
    const p = patternSeg.toUpperCase();
    const f = funclocSeg.toUpperCase();
    if (!p.includes("*")) return p === f;
    if (p.includes("F")) return /F/.test(f);
    return true;
}

function matchAbcPattern(funclocKey, patternSegments) {
    const flSegs = String(funclocKey)
        .trim()
        .split("-")
        .map(s => s.trim())
        .filter(Boolean);
    if (flSegs.length < patternSegments.length) return false;
    for (let i = 0; i < patternSegments.length; i++) {
        if (!segmentMatches(patternSegments[i], flSegs[i])) return false;
    }
    return true;
}

function lookupAbcFromCsv(funclocKey, rules) {
    for (const rule of rules) {
        if (matchAbcPattern(funclocKey, rule.segments)) return rule.abc;
    }
    return null;
}

function resolveAbcIndicator(r, funcLocKey, valFn, rules) {
    const src = valFn(r, "ABC INDICATORS");
    if (!isEmpty(src)) return src;
    const mapped = lookupAbcFromCsv(funcLocKey, rules);
    if (mapped !== null) return mapped;
    return ABC_DEFAULT;
}

module.exports = {
    ensureAbcRules,
    loadAbcRules,
    resolveAbcIndicator,
    matchAbcPattern,
    ABC_DEFAULT
};
