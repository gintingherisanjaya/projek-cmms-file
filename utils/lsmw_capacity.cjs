/**
 * Normalisasi nilai kolom CAPACITY untuk template LSMW (GROES).
 * Jika ada teks dalam kurung, bagian kurung dihapus.
 * Contoh: "19 Pintu (@ 15 Ton. TBS)" → "19 Pintu"
 */

function normalizeCapacityValue(value) {
    if (value === null || value === undefined || value === "") return value;

    let s = String(value).trim();
    if (!s) return value;

    const stripped = s.replace(/\s*\([^)]*\)/g, "").trim();
    return stripped === "" ? null : stripped;
}

module.exports = { normalizeCapacityValue };
