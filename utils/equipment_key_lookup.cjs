/**
 * Lookup map Equipment-key dengan primary + fallback key (pola EQUNR maintenance-plan).
 */

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

/**
 * @param {string} primaryKey
 * @param {string} fallbackKey
 * @param {Map<string, unknown> | undefined} map
 * @param {{ isEmpty: (hit: unknown) => boolean }} options
 * @returns {{ hit: unknown, fromPrimary: boolean }}
 */
function lookupMapByEquipmentKeys(primaryKey, fallbackKey, map, options) {
    const { isEmpty } = options;
    if (!map) {
        return { hit: null, fromPrimary: false };
    }

    const primary = cellText(primaryKey);
    if (primary && map.has(primary)) {
        const hit = map.get(primary);
        if (!isEmpty(hit)) {
            return { hit, fromPrimary: true };
        }
    }

    const fallback = cellText(fallbackKey);
    if (fallback && fallback !== primary && map.has(fallback)) {
        const hit = map.get(fallback);
        if (!isEmpty(hit)) {
            return { hit, fromPrimary: false };
        }
    }

    return { hit: null, fromPrimary: false };
}

module.exports = {
    lookupMapByEquipmentKeys
};
