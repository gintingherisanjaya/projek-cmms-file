import { toCleanString } from './excelHelpers.js';

/** Plant code (kodering plant) must be exactly this length for safety. */
export const PLANT_CODE_LENGTH = 4;

/**
 * Returns true only if the value is a valid 4-char plant code (trimmed, uppercase).
 * Avoids using header text or malformed values as plant.
 */
export function isValidPlantCode(value) {
  if (value == null) return false;
  const s = String(value).trim().toUpperCase();
  return s.length === PLANT_CODE_LENGTH;
}

/**
 * Get plant prefix (first 2 segments) from func loc string.
 * Example: PALM-2F01-0005-... -> PALM-2F01
 */
export function getFuncLocPrefix(raw) {
  if (!raw) return null;
  const segs = String(raw).toUpperCase().split('-').filter(Boolean);
  if (segs.length < 2) return null;
  return `${segs[0]}-${segs[1]}`;
}

/**
 * Adapt COST CENTER AFTER from template to target plant prefix.
 * Only applies replacement when targetPlantCode is a valid 4-char plant code.
 */
export function adaptCostCenter(templateCc, targetPlantCode) {
  if (!templateCc) return '';
  const cc = toCleanString(templateCc).toUpperCase();
  if (!targetPlantCode || !isValidPlantCode(targetPlantCode)) return cc;
  const plant = String(targetPlantCode).trim().toUpperCase();
  const replaceLen = plant.length;
  if (cc.length <= replaceLen) return plant.slice(0, cc.length);
  return plant + cc.slice(replaceLen);
}
