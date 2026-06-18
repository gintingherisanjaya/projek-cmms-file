/**
 * Second segment of functional location → plant prefix hint (before strict validation).
 * @param {string} raw
 */
export function getFuncLocPrefix(raw) {
  const parts = String(raw ?? '')
    .split('-')
    .map((p) => p.trim())
    .filter(Boolean);
  return parts.length >= 2 ? parts[1] : '';
}

/** 4-char alphanumeric plant codes used in PalmCo sheets. */
export function isValidPlantCode(code) {
  const s = String(code ?? '').trim().toUpperCase();
  return /^[A-Z0-9]{4}$/.test(s);
}
