/**
 * Mirrors webapp draft JSON import (`use-master-data-draft.ts` + `master-data-dendrogram-draft.ts`
 * + `master-data-compact-gaps.ts`) so Node apply scripts produce the same logical FUNCLOC/desc
 * as the dendrogram when given the same workbook + mapping JSON.
 */

const normalizeMappingText = (value) => String(value ?? '').trim().toUpperCase();

const parseFuncLocParts = (value) =>
  String(value)
    .trim()
    .toUpperCase()
    .split('-')
    .filter(Boolean);

/** Same ordering as `compareFuncLoc` in master-data-dendrogram-draft.ts */
export const compareFuncLocDraft = (a, b) => {
  const aa = parseFuncLocParts(a);
  const bb = parseFuncLocParts(b);
  const len = Math.max(aa.length, bb.length);
  for (let i = 0; i < len; i += 1) {
    const av = aa[i];
    const bv = bb[i];
    if (av == null) return -1;
    if (bv == null) return 1;
    const an = Number.parseInt(av, 10);
    const bn = Number.parseInt(bv, 10);
    const bothNumeric = Number.isFinite(an) && Number.isFinite(bn);
    if (bothNumeric) {
      if (an !== bn) return an - bn;
      continue;
    }
    const cmp = av.localeCompare(bv);
    if (cmp !== 0) return cmp;
  }
  return 0;
};

const getParentIdFromFl = (id) => {
  const parts = String(id).split('-').filter(Boolean);
  if (parts.length <= 1) return null;
  return parts.slice(0, -1).join('-');
};

/**
 * Same as `normalizeImportedMappings` in use-master-data-draft.ts (dedupe + sort by beforeFuncLoc only).
 */
export const normalizeImportedDraftChanges = (input) => {
  if (!Array.isArray(input)) return [];
  const out = [];

  for (const raw of input) {
    const beforeFuncLoc = normalizeMappingText(raw.beforeFuncLoc ?? raw.BEFORE_FUNCLOC);
    const afterFuncLoc = normalizeMappingText(raw.afterFuncLoc ?? raw.AFTER_FUNCLOC);
    const descBefore = String(raw.descBefore ?? raw.DESC_BEFORE ?? '').trim();
    const descAfter = String(raw.descAfter ?? raw.DESC_AFTER ?? raw.desc ?? raw.DESC ?? '').trim();
    const rawAction = String(raw.action ?? raw.ACTION ?? '').trim().toLowerCase();

    const action =
      rawAction === 'create'
        ? 'create'
        : rawAction === 'compact-gap' || rawAction === 'compact_gap'
          ? 'compact-gap'
          : afterFuncLoc === ''
            ? 'delete'
            : beforeFuncLoc !== afterFuncLoc && descBefore !== descAfter
              ? 'move+rename'
              : beforeFuncLoc !== afterFuncLoc
                ? 'move'
                : 'rename';

    if (!beforeFuncLoc && action !== 'create') continue;
    if (!afterFuncLoc && action === 'create') continue;

    out.push({
      beforeFuncLoc,
      afterFuncLoc,
      descBefore,
      descAfter,
      action,
      parentBefore: getParentIdFromFl(beforeFuncLoc) ?? '',
      parentAfter: getParentIdFromFl(afterFuncLoc) ?? '',
    });
  }

  const deduped = new Map();
  for (const row of out) {
    const key = row.action === 'create' ? `create:${row.afterFuncLoc}` : `before:${row.beforeFuncLoc}`;
    deduped.set(key, row);
  }

  return Array.from(deduped.values()).sort((a, b) => compareFuncLocDraft(a.beforeFuncLoc, b.beforeFuncLoc));
};

const toSuffix4 = (n) => String(n).padStart(4, '0');

const depth = (id) => id.split('-').filter(Boolean).length;

const lastSegment = (id) => id.split('-').filter(Boolean).at(-1) ?? '';

const lastSegmentNum = (id) => {
  const n = Number.parseInt(lastSegment(id), 10);
  return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
};

export const isLoryDescException = (desc) => /\blory/i.test(String(desc).trim());

export const pickCompactSuffixes = (count, reserved) => {
  const out = [];
  let n = 1;
  for (let i = 0; i < count; i += 1) {
    while (reserved.has(n)) n += 1;
    out.push(n);
    n += 1;
  }
  return out;
};

const replaceFuncLocPrefix = (full, oldRoot, newRoot) => {
  if (full === oldRoot) return newRoot;
  if (full.startsWith(`${oldRoot}-`)) return `${newRoot}${full.slice(oldRoot.length)}`;
  return full;
};

const toDuplicateCountMap = (ids) => {
  const counts = new Map();
  for (const id of ids) counts.set(id, (counts.get(id) ?? 0) + 1);
  const dup = new Map();
  for (const [id, count] of counts.entries()) {
    if (count > 1) dup.set(id, count);
  }
  return dup;
};

const wouldIncreaseDuplicateCounts = (beforeIds, afterIds) => {
  const beforeDup = toDuplicateCountMap(beforeIds);
  const afterDup = toDuplicateCountMap(afterIds);
  for (const [id, afterCount] of afterDup.entries()) {
    const beforeCount = beforeDup.get(id) ?? 0;
    if (afterCount > beforeCount) return true;
  }
  return false;
};

const getParentIdCompact = (id) => {
  const parts = id.split('-').filter(Boolean);
  if (parts.length <= 1) return null;
  return parts.slice(0, -1).join('-');
};

/**
 * Same as `compactGapFuncLocDraft` in master-data-compact-gaps.ts
 */
export const compactGapFuncLocDraft = (rows) => {
  const working = rows.map((r) => ({ ...r }));
  const initialDuplicateCounts = toDuplicateCountMap(working.map((r) => r.funcLocAfter));

  const maxDepth = Math.max(0, ...working.map((r) => depth(r.funcLocAfter)));

  for (let childDepth = 4; childDepth <= maxDepth; childDepth += 1) {
    const parentDepth = childDepth - 1;
    if (parentDepth < 3) continue;

    const childrenByParent = new Map();
    for (const r of working) {
      const p = getParentIdCompact(r.funcLocAfter);
      if (!p) continue;
      if (depth(p) !== parentDepth) continue;
      const arr = childrenByParent.get(p);
      if (arr) arr.push(r);
      else childrenByParent.set(p, [r]);
    }

    const sortedParents = [...childrenByParent.keys()].sort((a, b) => {
      const da = depth(a);
      const db = depth(b);
      if (da !== db) return da - db;
      return a.localeCompare(b);
    });

    for (const P of sortedParents) {
      const children = childrenByParent.get(P);
      if (!children || children.length <= 1) continue;

      const fixedKids = children.filter((r) => isLoryDescException(r.desc));
      const rest = children.filter((r) => !isLoryDescException(r.desc));
      if (rest.length === 0) continue;
      rest.sort(
        (a, b) =>
          lastSegmentNum(a.funcLocAfter) - lastSegmentNum(b.funcLocAfter) ||
          a.funcLocAfter.localeCompare(b.funcLocAfter),
      );

      const reserved = new Set();
      for (const r of fixedKids) {
        const n = Number.parseInt(lastSegment(r.funcLocAfter), 10);
        if (Number.isFinite(n)) reserved.add(n);
      }
      const suffixes = pickCompactSuffixes(rest.length, reserved);
      const remap = new Map();

      for (let i = 0; i < rest.length; i += 1) {
        const row = rest[i];
        const oldId = row.funcLocAfter;
        const newId = `${P}-${toSuffix4(suffixes[i])}`;
        if (oldId !== newId) remap.set(oldId, newId);
      }

      if (remap.size === 0) continue;

      const toTemp = new Map();
      const tempToNew = new Map();
      let tmpIdx = 0;
      for (const [oldId, newId] of remap.entries()) {
        const tempId = `${P}-TMP-COMPACT-${String(tmpIdx).padStart(4, '0')}`;
        tmpIdx += 1;
        toTemp.set(oldId, tempId);
        tempToNew.set(tempId, newId);
      }

      const beforeIds = working.map((r) => r.funcLocAfter);
      const tentative = working.map((r) => ({ ...r }));
      const remapOldRoots = [...toTemp.keys()];
      const rowTouchesAnyRemapRoot = (id) => {
        for (const oldId of remapOldRoots) {
          if (id === oldId || id.startsWith(`${oldId}-`)) return true;
        }
        return false;
      };
      const tempRoots = [...tempToNew.keys()];
      const rowTouchesAnyTemp = (id) => {
        for (const t of tempRoots) {
          if (id === t || id.startsWith(`${t}-`)) return true;
        }
        return false;
      };

      for (const r of tentative) {
        if (!rowTouchesAnyRemapRoot(r.funcLocAfter)) continue;
        let id = r.funcLocAfter;
        for (const [oldId, tempId] of toTemp.entries()) {
          id = replaceFuncLocPrefix(id, oldId, tempId);
        }
        r.funcLocAfter = id;
      }
      for (const r of tentative) {
        if (!rowTouchesAnyTemp(r.funcLocAfter)) continue;
        let id = r.funcLocAfter;
        for (const [tempId, newId] of tempToNew.entries()) {
          id = replaceFuncLocPrefix(id, tempId, newId);
        }
        r.funcLocAfter = id;
      }

      const afterIds = tentative.map((r) => r.funcLocAfter);
      if (wouldIncreaseDuplicateCounts(beforeIds, afterIds)) {
        continue;
      }
      for (let i = 0; i < working.length; i += 1) {
        working[i] = tentative[i];
      }
    }
  }

  const finalDuplicateCounts = toDuplicateCountMap(working.map((r) => r.funcLocAfter));
  let introducedDuplicateConflict = false;
  for (const [id, finalCount] of finalDuplicateCounts.entries()) {
    const initialCount = initialDuplicateCounts.get(id) ?? 0;
    if (finalCount > initialCount) {
      introducedDuplicateConflict = true;
      break;
    }
  }
  if (introducedDuplicateConflict) {
    return [...rows]
      .map((r) => ({
        ...r,
        hiddenBySimilarNumber: r.hiddenBySimilarNumber ?? false,
        equipmentGroupAfterRedBg: r.equipmentGroupAfterRedBg ?? false,
      }))
      .sort((a, b) => a.funcLocAfter.localeCompare(b.funcLocAfter));
  }

  return working
    .map((r) => ({
      ...r,
      hiddenBySimilarNumber: r.hiddenBySimilarNumber ?? false,
      equipmentGroupAfterRedBg: r.equipmentGroupAfterRedBg ?? false,
    }))
    .sort((a, b) => a.funcLocAfter.localeCompare(b.funcLocAfter));
};

const toLinearRows = (rows) => [...rows].sort((a, b) => compareFuncLocDraft(a.funcLocAfter, b.funcLocAfter));

/**
 * Same as `applyDraftMappingsToRows` in master-data-dendrogram-draft.ts
 */
export const applyDraftMappingsToRows = (baseRows, mappings, opts = {}) => {
  if (mappings.length === 0) return baseRows;
  const replaceExistingOnCreate = opts.replaceExistingOnCreate === true;

  const byOriginal = new Map();
  for (const row of baseRows) byOriginal.set(row.originalFuncLocAfter, { ...row });

  const deleted = new Set();
  const created = new Map();

  for (const mapRow of mappings) {
    if (mapRow.action === 'create') {
      if (!mapRow.afterFuncLoc) continue;
      if (replaceExistingOnCreate) {
        for (const current of byOriginal.values()) {
          if (current.funcLocAfter !== mapRow.afterFuncLoc) continue;
          current.desc = mapRow.descAfter ?? '';
          break;
        }
      }
      created.set(mapRow.afterFuncLoc, {
        funcLocAfter: mapRow.afterFuncLoc,
        desc: mapRow.descAfter ?? '',
        originalFuncLocAfter: '',
        hiddenBySimilarNumber: false,
        equipmentGroupAfterRedBg: false,
      });
      continue;
    }

    const current = byOriginal.get(mapRow.beforeFuncLoc);
    if (!current) continue;

    if (!mapRow.afterFuncLoc) {
      deleted.add(mapRow.beforeFuncLoc);
      continue;
    }

    current.funcLocAfter = mapRow.afterFuncLoc;
    if (mapRow.descAfter !== '') current.desc = mapRow.descAfter;
  }

  const nextRows = Array.from(byOriginal.entries())
    .filter(([original]) => !deleted.has(original))
    .map(([, row]) => row)
    .concat(Array.from(created.values()))
    .filter((row, idx, arr) => arr.findIndex((r) => r.funcLocAfter === row.funcLocAfter) === idx);

  return toLinearRows(nextRows).map((row) => ({
    ...row,
    hiddenBySimilarNumber: row.hiddenBySimilarNumber ?? false,
    equipmentGroupAfterRedBg: row.equipmentGroupAfterRedBg ?? false,
  }));
};
