/**
 * Kolom regional di cwc.csv / planner-group.csv dari maintenance plant (4 char, e.g. 8F01 → "8").
 */

/** @param {Iterable<string>} availableColumns Header kolom CSV (1, 2, …, K, 8, …). */
export function normalizeMaintenancePlantCode(raw) {
  const s = String(raw ?? '')
    .replace(/[`]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toUpperCase();
  if (!s) return null;
  return s.length >= 4 ? s.slice(0, 4) : s;
}

/** @param {Iterable<string>} availableColumns */
export function regionColumnFromPlant(plantCode, availableColumns, regionOverride = null) {
  const allow = availableColumns instanceof Set ? availableColumns : new Set(availableColumns);

  if (regionOverride) {
    return String(regionOverride).trim().toUpperCase();
  }

  const code = String(plantCode ?? '').trim().toUpperCase();
  const first = code[0] ?? '';
  if (/^[1-9]$/.test(first) || /^[A-Z]$/.test(first)) {
    if (allow.has(first)) return first;
  }

  const second = code[1] ?? '';
  return second === 'F' ? 'K' : '1';
}

/**
 * @param {string} maintenancePlant SWERK / MAINTENANCE PLAN / plant per file
 * @param {Iterable<string>} availableColumns
 */
export function getRegionalColumnForMaintenancePlant(
  maintenancePlant,
  availableColumns,
  regionOverride = null,
) {
  const plantCode = normalizeMaintenancePlantCode(maintenancePlant);
  if (!plantCode) {
    return { plantCode: null, regionColumn: null };
  }
  const regionColumn = regionColumnFromPlant(plantCode, availableColumns, regionOverride);
  return { plantCode, regionColumn };
}

/**
 * Suffix baris cwc.csv / planner-group.csv (STAS01..STAS14), tanpa tergantung prefix plant di string CC.
 */
export function extractCostCenterSuffixForCsv(costCenter) {
  const cleaned = String(costCenter ?? '')
    .replace(/[`]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toUpperCase();
  if (!cleaned) return '';

  const stasMatch = cleaned.match(/STAS(\d{2,3})/);
  if (stasMatch) {
    return `STAS${String(parseInt(stasMatch[1], 10)).padStart(2, '0')}`;
  }

  const noStars = cleaned.replace(/^\*{4}/, '');
  if (/^STAS\d{2}$/i.test(noStars)) return noStars.toUpperCase();

  if (/^[A-Z0-9]{4}/.test(noStars) && noStars.length > 4) {
    return noStars.slice(4).replace(/^\*{4}/, '');
  }

  return noStars;
}
