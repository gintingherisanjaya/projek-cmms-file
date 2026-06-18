/**
 * Kunci urutan numerik folder REGIONAL* dan file PKS (prefix "N. ").
 */

function parseRegionalNumber(name) {
    const m = String(name ?? "").match(/REGIONAL\s*(\d+)/i);
    if (!m) return Number.MAX_SAFE_INTEGER;
    const n = parseInt(m[1], 10);
    return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
}

function parsePksFilePrefix(name) {
    const base = String(name ?? "").replace(/\.xlsx$/i, "").trim();
    const m = base.match(/^(\d+)\./);
    if (!m) return Number.MAX_SAFE_INTEGER;
    const n = parseInt(m[1], 10);
    return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
}

/**
 * @param {{ regional: string, file: { name: string } }} a
 * @param {{ regional: string, file: { name: string } }} b
 */
function compareRegionalFileEntries(a, b) {
    const byRegional = parseRegionalNumber(a.regional) - parseRegionalNumber(b.regional);
    if (byRegional !== 0) return byRegional;

    const byFile = parsePksFilePrefix(a.file.name) - parsePksFilePrefix(b.file.name);
    if (byFile !== 0) return byFile;

    return String(a.file.name).localeCompare(String(b.file.name), "id");
}

module.exports = {
    parseRegionalNumber,
    parsePksFilePrefix,
    compareRegionalFileEntries
};
