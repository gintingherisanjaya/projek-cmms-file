import { normalizeDescEqktu } from './validationHelpers.js';
import { getFuncLocPrefix, isValidPlantCode } from './funclocPlantAdapt.js';

export const fixNoDotNumberSpacing = (value) => {
  const cleaned = String(value ?? '')
    .replace(/[\u200B-\u200F\u202A-\u202E\u2060\u2066-\u2069\uFEFF\u00AD]/g, '')
    .replace(/[\u0000-\u001F\u007F\u0080-\u009F]/g, ' ');
  return cleaned.replace(/(n[o0])\s*\.?\s*(\d+)/gi, 'NO. $2');
};

export const normalizeWhitespace = (value) => fixNoDotNumberSpacing(value).replace(/\s+/g, ' ').trim();

export const fixFibreToFiber = (value) => {
  const s = String(value ?? '');
  return s.replace(/fibre/gi, (m) => {
    if (m === m.toUpperCase()) return 'FIBER';
    if (m[0] && m[0] === m[0].toUpperCase() && m.slice(1) === m.slice(1).toLowerCase()) {
      return 'Fiber';
    }
    return 'fiber';
  });
};

export const rowEffectivePlantCode = (
  worksheet,
  r,
  maintenancePlantCol,
  funclocRaw,
  fileMaintenancePlant,
  { encodeCell, toCleanString },
) => {
  if (maintenancePlantCol !== undefined) {
    const addr = encodeCell({ r, c: maintenancePlantCol });
    const raw = toCleanString(worksheet[addr]?.v);
    if (raw) {
      const upper = raw.trim().toUpperCase();
      if (isValidPlantCode(upper)) return upper;
    }
  }
  if (funclocRaw) {
    const segs = funclocRaw.toUpperCase().split('-').filter(Boolean);
    if (segs.length >= 2 && isValidPlantCode(segs[1])) return segs[1].trim().toUpperCase();
  }
  return fileMaintenancePlant && isValidPlantCode(fileMaintenancePlant) ? fileMaintenancePlant : null;
};

export const adaptFuncLocFromTemplateToRow = (funclocRaw, fileMaintenancePlant, tpl) => {
  const funcUpper = String(funclocRaw ?? '')
    .toUpperCase()
    .trim();
  let rowPrefix = getFuncLocPrefix(funcUpper);
  if (!rowPrefix) {
    const segs = funcUpper.split('-').filter(Boolean);
    if (segs.length >= 2 && isValidPlantCode(segs[1])) {
      rowPrefix = `${segs[0]}-${segs[1]}`;
    } else if (fileMaintenancePlant && isValidPlantCode(fileMaintenancePlant)) {
      const tplSeg0 = String(tpl.funcLoc ?? '')
        .toUpperCase()
        .split('-')
        .filter(Boolean);
      if (tplSeg0.length >= 1) rowPrefix = `${tplSeg0[0]}-${fileMaintenancePlant}`;
    }
  }
  const tplSegs = String(tpl.funcLoc ?? '')
    .toUpperCase()
    .split('-')
    .filter(Boolean);
  if (tplSegs.length < 3 || !rowPrefix) return null;
  const rest = tplSegs.slice(2).join('-');
  return `${rowPrefix}-${rest}`;
};

export const fuzzyBigramSimilarity = (left, right) => {
  const a = normalizeDescEqktu(left ?? '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
  const b = normalizeDescEqktu(right ?? '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
  if (!a || !b) return 0;
  if (a === b) return 1;
  const toBigrams = (value) => {
    const normalized = value.replace(/\s+/g, ' ');
    if (normalized.length < 2) return [normalized];
    const out = [];
    for (let i = 0; i < normalized.length - 1; i += 1) out.push(normalized.slice(i, i + 2));
    return out;
  };
  const aBigrams = toBigrams(a);
  const bBigrams = toBigrams(b);
  const bCounts = new Map();
  bBigrams.forEach((gram) => {
    bCounts.set(gram, (bCounts.get(gram) ?? 0) + 1);
  });
  let intersection = 0;
  aBigrams.forEach((gram) => {
    const count = bCounts.get(gram) ?? 0;
    if (count > 0) {
      intersection += 1;
      bCounts.set(gram, count - 1);
    }
  });
  return (2 * intersection) / (aBigrams.length + bBigrams.length);
};
