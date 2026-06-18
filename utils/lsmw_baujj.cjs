/**
 * Resolusi nilai BAUJJ: umur teknis 4 digit, else CONSTRUCTION YEAR.
 */

function isFourDigitYear(value) {
    if (value === null || value === undefined) return false;
    return /^\d{4}$/.test(String(value).trim());
}

function cellTextOrNull(value) {
    if (value === null || value === undefined) return null;
    const text = String(value).trim();
    return text === "" ? null : text;
}

/**
 * @param {Array<unknown>} r
 * @param {(row: Array<unknown>, columnName: string) => unknown} getValueWithParent
 */
function resolveBaujjYear(r, getValueWithParent) {
    const umur = getValueWithParent(r, "STANDART UMUR TEKNIS (TAHUN)");
    if (isFourDigitYear(umur)) return String(umur).trim();

    const construction = getValueWithParent(r, "CONSTRUCTION YEAR");
    return cellTextOrNull(construction);
}

module.exports = {
    isFourDigitYear,
    resolveBaujjYear
};
