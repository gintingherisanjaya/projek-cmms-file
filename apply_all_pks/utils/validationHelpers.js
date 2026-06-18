/**
 * Validation helper functions shared across validation scripts.
 */

// Color definitions (ARGB format for exceljs)
export const COLOR_GREEN = 'FFC6EFCE'; // Light green
export const COLOR_YELLOW = 'FFFFEB9C'; // Light yellow
export const COLOR_RED = 'FFFFC7CE'; // Light red
export const COLOR_ORANGE = 'FFF4B183'; // Light orange
export const COLOR_HEADER_BLUE = 'FFDCE6F1'; // Soft blue for column headers

const EMPTY_AFTER_COST_CENTER_MARKERS = new Set([
  'POSISIALATTIDAKDIKETAHUI',
  'POSISIALATDILUARPABIK',
  'POSISIALATDILUARPABRIK',
]);

export const normalizeAfterCostCenterValue = (value) => {
  const cleaned = String(value ?? '').trim();
  if (!cleaned) return '';
  const markerKey = cleaned.toUpperCase().replace(/\s+/g, '');
  return EMPTY_AFTER_COST_CENTER_MARKERS.has(markerKey) ? '' : cleaned;
};

/** Plant-code pattern (e.g. 2F01, 1F02): exclude PALM-<this> from key so we don't treat plant-only as template row. */
const PLANT_CODE_PATTERN = /^[0-9A-Z]F[0-9]{2}$/;

/**
 * Normalize FUNCLoc for combo matching: skip first 2 segments (split by '-').
 * Includes 1- and 2-segment FUNCLOCs except "PALM-XFXX" (plant-only, e.g. PALM-2F01).
 * @param {string} raw - FUNCLoc string (cleaned, uppercase)
 * @returns {string} Key: segments after first 2 for 3+ segments; full string for 1 seg or 2 seg (unless PALM-XFXX); '' for PALM-XFXX
 */
export const funclocKeyAfterFirst2 = (raw) => {
  if (!raw || typeof raw !== 'string') return '';
  const segs = raw.trim().toUpperCase().split('-').filter(Boolean);
  if (segs.length >= 3) return segs.slice(2).join('-');
  if (segs.length === 1) return segs[0];
  if (segs.length === 2) {
    if (segs[0] === 'PALM' && PLANT_CODE_PATTERN.test(segs[1])) return '';
    return segs.join('-');
  }
  return '';
};

/**
 * Key for BAH JAMBI template and source combo: strip leading FUNCLoc segments via {@link funclocKeyAfterFirst2}
 * after uppercasing and removing whitespace (handles odd Excel spacing).
 * @param {string|null|undefined} raw - Raw FUNCLoc cell value
 * @returns {string}
 */
export const funclocTemplateLookupKey = (raw) => {
  if (raw === null || raw === undefined) return '';
  const u = String(raw).trim().toUpperCase().replace(/\s+/g, '');
  return u ? funclocKeyAfterFirst2(u) : '';
};

/**
 * Normalize DESC/EQKTU values for comparison: trim and normalize whitespace.
 * This handles cases where values appear identical but have different whitespace.
 * @param {string} value - Raw DESC or EQKTU value
 * @returns {string} Normalized value (trimmed, whitespace collapsed to single spaces)
 */
export const normalizeDescEqktu = (value) => {
  if (value === null || value === undefined) return '';
  const s = String(value);

  // Spelling autofix: normalize "fibre" -> "fiber" (case-preserving for the replacement).
  const withSpellingFixed = s.replace(/fibre/gi, (m) => {
    if (m === m.toUpperCase()) return 'FIBER';
    if (m[0] && m[0] === m[0].toUpperCase() && m.slice(1) === m.slice(1).toLowerCase()) return 'Fiber';
    return 'fiber';
  });

  // Spacing normalization:
  // - unify variants into "NO. 1": NO.X / NO. X / NO .X / NO . X / NO X / NO..X
  // - also accepts common typo "N0" (zero) for "NO"
  // - no word-boundary needed so it also fixes cases like "DIGESTERNO.4"
  //   and double-dot typos like "NO..5".
  const withNoSpacingFixed = withSpellingFixed.replace(/(n[o0])\s*(?:\.\s*)*(\d+)/gi, 'NO. $2');

  // Remove invisible/control unicode that can break key equality.
  const withSafeWhitespace = withNoSpacingFixed
    .replace(/[\u200B-\u200F\u202A-\u202E\u2060\u2066-\u2069\uFEFF\u00AD]/g, '')
    .replace(/[\u0000-\u001F\u007F\u0080-\u009F]/g, ' ');

  return withSafeWhitespace.trim().replace(/\s+/g, ' '); // Collapse whitespace to single space
};

/**
 * Split a reference combo pair key ("DESC|||EQKTU") from source-funcloc.xlsx.
 * @param {string} pairKey
 * @returns {{ desc: string, eqktu: string }}
 */
export const splitSourceComboPairKey = (pairKey) => {
  const s = String(pairKey ?? '');
  const idx = s.indexOf('|||');
  if (idx < 0) return { desc: s.trim(), eqktu: '' };
  return { desc: s.slice(0, idx).trim(), eqktu: s.slice(idx + 3).trim() };
};

/**
 * Pick canonical DESC/EQKTU from the source reference set for template auto-fix.
 * Prefers the pair whose normalized DESC matches the row (fixes EQKTU typos when DESC matches source).
 * @param {Set<string>|undefined} validPairsRaw
 * @param {string} targetDescRaw
 * @param {string} targetEqktuRaw
 * @returns {{ desc: string, eqktu: string } | null}
 */
export const pickCanonicalSourceComboPair = (validPairsRaw, targetDescRaw, targetEqktuRaw) => {
  const arr = validPairsRaw ? Array.from(validPairsRaw) : [];
  if (arr.length === 0) return null;
  const norm = (v) => normalizeDescEqktu(v).toUpperCase();
  const nd = norm(targetDescRaw);
  for (const pk of arr) {
    const { desc, eqktu } = splitSourceComboPairKey(pk);
    if (norm(desc) === nd) {
      return { desc: normalizeDescEqktu(desc), eqktu: normalizeDescEqktu(eqktu) };
    }
  }
  const sorted = [...arr].sort();
  const { desc, eqktu } = splitSourceComboPairKey(sorted[0]);
  return { desc: normalizeDescEqktu(desc), eqktu: normalizeDescEqktu(eqktu) };
};

/** Truncate long strings for description cell (maxLen each, add "…" if truncated) */
export const truncateForDesc = (s, maxLen = 80) => {
  const t = String(s ?? '').trim();
  if (t.length <= maxLen) return t;
  return t.slice(0, maxLen) + '…';
};

/**
 * Get functional location level by counting segments after splitting by "-".
 * Level 5: PALM-XFXX-AAAA-BBBB-CCCC (5 segments). Level 6: ...-DDDD (6 segments).
 * @param {string} funcloc - Functional location string
 * @returns {number|null} 5, 6, or null if invalid/unknown
 */
export const getFunclocLevel = (funcloc) => {
  if (!funcloc || typeof funcloc !== 'string') return null;
  const raw = String(funcloc).trim().toUpperCase();
  if (!raw) return null;
  const segments = raw.split('-').filter(Boolean);
  if (segments.length === 5) return 5;
  if (segments.length === 6) return 6;
  return null;
};

/**
 * Get parent (level 5) functional location from a level 6 string.
 * Returns first 5 segments joined by "-".
 * @param {string} funcloc - Level 6 functional location
 * @returns {string|null} Parent level 5 funcloc or null
 */
export const getFunclocParent = (funcloc) => {
  if (!funcloc || typeof funcloc !== 'string') return null;
  const raw = String(funcloc).trim().toUpperCase();
  if (!raw) return null;
  const segments = raw.split('-').filter(Boolean);
  if (segments.length < 6) return null;
  return segments.slice(0, 5).join('-');
};

/**
 * Analyze differences between two strings to help debug mismatches
 * @param {string} str1 - First string
 * @param {string} str2 - Second string
 * @returns {string} Description of differences
 */
const analyzeStringDifference = (str1, str2) => {
  const s1 = String(str1 ?? '');
  const s2 = String(str2 ?? '');

  if (s1 === s2) return ' (identical)';

  const len1 = s1.length;
  const len2 = s2.length;
  const lenDiff = len1 !== len2 ? ` [Length: ${len1} vs ${len2}]` : '';

  // Find first difference
  let firstDiffPos = -1;
  const minLen = Math.min(len1, len2);
  for (let i = 0; i < minLen; i++) {
    if (s1[i] !== s2[i]) {
      firstDiffPos = i;
      break;
    }
  }
  if (firstDiffPos === -1 && len1 !== len2) {
    firstDiffPos = minLen;
  }

  let diffInfo = '';
  if (firstDiffPos >= 0) {
    const char1 = firstDiffPos < len1 ? s1[firstDiffPos] : '(end)';
    const char2 = firstDiffPos < len2 ? s2[firstDiffPos] : '(end)';
    const code1 = firstDiffPos < len1 ? s1.charCodeAt(firstDiffPos) : null;
    const code2 = firstDiffPos < len2 ? s2.charCodeAt(firstDiffPos) : null;

    // Show character and code
    const showChar = (c, code) => {
      if (code === null) return c === '(end)' ? '(end)' : `'${c}' (N/A)`;
      const codeHex = code.toString(16).toUpperCase().padStart(4, '0');
      if (c === ' ') return `' ' (space, U+${codeHex})`;
      if (c === '\t') return `'\\t' (tab, U+${codeHex})`;
      if (c === '\n') return `'\\n' (newline, U+${codeHex})`;
      if (c === '\r') return `'\\r' (carriage return, U+${codeHex})`;
      if (code < 32 || code > 126) return `'${c}' (U+${codeHex})`;
      return `'${c}' (U+${codeHex})`;
    };

    diffInfo = ` [First diff at pos ${firstDiffPos}: ${showChar(char1, code1)} vs ${showChar(char2, code2)}]`;
  }

  // Check for non-printable characters
  const nonPrintable1 = Array.from(s1).some((c) => {
    const code = c.charCodeAt(0);
    return (code < 32 && code !== 9 && code !== 10 && code !== 13) || code > 126;
  });
  const nonPrintable2 = Array.from(s2).some((c) => {
    const code = c.charCodeAt(0);
    return (code < 32 && code !== 9 && code !== 10 && code !== 13) || code > 126;
  });

  let nonPrintableInfo = '';
  if (nonPrintable1 || nonPrintable2) {
    nonPrintableInfo = ' [Contains non-printable characters]';
  }

  return lenDiff + diffInfo + nonPrintableInfo;
};

/**
 * Generate description string from validation results
 * @param {Object} validation - Validation results for a row
 * @returns {string} Description of validation issues
 */
export const generateDescription = (validation) => {
  const issues = [];

  if (validation.cwcOk === false) {
    issues.push('No CWC found');
  }

  if (validation.duplicateFunclocOk === false) {
    const firstRow = validation.duplicateFunclocFirstRow;
    const rowRef = firstRow != null ? ` - first at row ${firstRow}` : '';
    issues.push(
      `Duplicate Functional Location AFTER with different DESC/EQKTU than first occurrence${rowRef}`,
    );
  }

  if (validation.duplicateDescOk === false) {
    const firstRow = validation.duplicateDescFirstRow;
    const rowRef = firstRow != null ? ` - first at row ${firstRow}` : '';
    issues.push(
      `Duplicate FUNCTLOC DESC with different Functional Location than first occurrence${rowRef}`,
    );
  }

  if (validation.duplicateEqktuOk === false) {
    const firstRow = validation.duplicateEqktuFirstRow;
    const rowRef = firstRow != null ? ` - first at row ${firstRow}` : '';
    issues.push(
      `Duplicate EQKTU AFTER with different Functional Location or DESC than first occurrence${rowRef}`,
    );
  }

  if (validation.duplicateEquipmentObjectCostCenterBeforeOk === false) {
    const firstRow = validation.duplicateEquipmentObjectCostCenterBeforeFirstRow;
    const rowRef = firstRow != null ? ` - first at row ${firstRow}` : '';
    issues.push(
      `Duplicate combination of EQUIPMENT NUMBER, OBJECT NUMBER, COST CENTER BEFORE, FUNCTIONAL LOCATION BEFORE, and EQKTU BEFORE${rowRef}`,
    );
  }

  if (validation.cwcBeforeOk === false) {
    issues.push('No CWC mapping for COST CENTER BEFORE');
  }

  if (validation.beforeTruthOk === false) {
    issues.push('BEFORE data combination not found in TRUTH_OF_DATA');
  }

  if (validation.plannerGroupOk === false) {
    issues.push('No Planner Group found');
  }

  if (validation.funcOk === false) {
    issues.push('Functional Location plant mismatch');
  }

  if (validation.costOk === false) {
    issues.push('Cost Center mismatch');
  }

  if (validation.comboOk === false) {
    const d = validation.comboMismatchDetail;
    if (d && (d.sourcePairs?.length || d.targetDesc !== undefined || d.targetEqktu !== undefined)) {
      const targetPart = `Target: DESC "${truncateForDesc(d.targetDesc)}" | EQKTU "${truncateForDesc(d.targetEqktu)}"`;
      const fmt = (p) => `DESC "${truncateForDesc(p.desc)}" | EQKTU "${truncateForDesc(p.eqktu)}"`;
      const sourcePart = d.sourcePairs?.length
        ? d.sourcePairs.length === 1
          ? `Source expects: ${fmt(d.sourcePairs[0])}`
          : `Source expects one of: ${d.sourcePairs.map((p, i) => (i ? '; ' : '') + fmt(p)).join('')}`
        : 'Source expects: (see reference)';

      // Add detailed difference analysis
      let diffDetails = [];
      if (d.sourcePairs?.length === 1) {
        const source = d.sourcePairs[0];
        const descDiff = analyzeStringDifference(source.desc, d.targetDesc);
        const eqktuDiff = analyzeStringDifference(source.eqktu, d.targetEqktu);
        if (descDiff !== ' (identical)') {
          diffDetails.push(`DESC diff:${descDiff}`);
        }
        if (eqktuDiff !== ' (identical)') {
          diffDetails.push(`EQKTU diff:${eqktuDiff}`);
        }
      }

      const diffInfo = diffDetails.length > 0 ? ` [${diffDetails.join('; ')}]` : '';
      issues.push(`Combo mismatch. ${sourcePart}. ${targetPart}${diffInfo}`);
    } else {
      issues.push('FUNCTLOC DESC/EQKTU mismatch vs source');
    }
  }

  return issues.length > 0 ? issues.join('; ') : 'OK';
};

/**
 * Determine row color based on validation results
 * @param {Object} validation - Validation results for a row
 * @returns {string} Color ARGB value
 */
export const determineRowColor = (validation) => {
  // Hard failures (CWC / Planner Group / COST CENTER BEFORE / key DESC/EQKTU duplicates) if enabled
  if (validation.cwcOk === false) return COLOR_RED;
  if (validation.plannerGroupOk === false) return COLOR_RED;
  if (validation.cwcBeforeOk === false) return COLOR_RED;
  if (validation.beforeTruthOk === false) return COLOR_RED;
  if (validation.duplicateFunclocOk === false) return COLOR_YELLOW;
  if (validation.duplicateDescOk === false) return COLOR_RED;
  if (validation.duplicateEqktuOk === false) return COLOR_RED;
  if (validation.duplicateEquipmentObjectCostCenterBeforeOk === false) return COLOR_RED;

  // Field validations: RED only if *all present* validations fail, YELLOW if partial, GREEN if all pass.
  const checks = [];
  if (validation.funcOk !== undefined) checks.push(validation.funcOk);
  if (validation.costOk !== undefined) checks.push(validation.costOk);
  if (validation.comboOk !== undefined) checks.push(validation.comboOk);

  if (!checks.length) return COLOR_GREEN;
  const allFail = checks.every((v) => v === false);
  const anyFail = checks.some((v) => v === false);

  if (allFail) return COLOR_RED;
  if (anyFail) return COLOR_YELLOW;
  return COLOR_GREEN;
};
