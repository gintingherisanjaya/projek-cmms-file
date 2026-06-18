import fs from 'node:fs';

export const EQUIPMENT_GROUP_AFTER_MAX_LEN = 10;
export const STATION_FUNCLOC_SEGMENT_COUNT = 4;

const normalizeEquipmentGroupValue = (value) =>
  String(value ?? '')
    .replace(/\s+/g, '')
    .slice(0, EQUIPMENT_GROUP_AFTER_MAX_LEN);

const normalizeDescForMatch = (desc) =>
  String(desc ?? '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');

const getFuncLocDepth = (funcLocNorm) => {
  const parts = String(funcLocNorm ?? '')
    .trim()
    .toUpperCase()
    .split('-')
    .filter(Boolean);
  return parts.length;
};

export const getDirectParentFuncLocNorm = (funcLocNorm) => {
  const parts = String(funcLocNorm ?? '')
    .trim()
    .toUpperCase()
    .split('-')
    .filter(Boolean);
  if (parts.length < 2) return '';
  return parts.slice(0, -1).join('-');
};

/** True when FUNCLOC AFTER is exactly four hyphen segments (stasiun), e.g. PALM-2F01-0005-0001. */
export const isStationDepthFuncLoc = (funcLocNorm) =>
  getFuncLocDepth(funcLocNorm) === STATION_FUNCLOC_SEGMENT_COUNT;

const buildOrderedRules = (section) => {
  const rules = [];
  for (const [phrase, mappedValue] of Object.entries(section ?? {})) {
    const key = String(phrase ?? '').trim();
    if (!key) continue;
    rules.push({
      phrase: key.toUpperCase(),
      value: normalizeEquipmentGroupValue(mappedValue),
    });
  }
  rules.sort((a, b) => b.phrase.length - a.phrase.length || a.phrase.localeCompare(b.phrase));
  return rules;
};

const buildChildFollowParentPhrases = (rawList) => {
  const phrases = (rawList ?? [])
    .map((phrase) => String(phrase ?? '').trim().toUpperCase())
    .filter(Boolean);
  phrases.sort((a, b) => b.length - a.length || a.localeCompare(b));
  return phrases;
};

const filterRulesExcludingChildFollow = (rules, childFollowPhraseSet, sectionLabel) => {
  const warned = new Set();
  return (rules ?? []).filter(rule => {
    if (!childFollowPhraseSet.has(rule.phrase)) return true;
    if (!warned.has(rule.phrase)) {
      warned.add(rule.phrase);
      console.warn(
        `[equipment-group-mapping] Frasa "${rule.phrase}" diabaikan di tier ${sectionLabel} ` +
          '(hanya untuk child_follow_parent_group).'
      );
    }
    return false;
  });
};

/**
 * @param {string} mappingPath
 */
export const loadEquipmentGroupMapping = (mappingPath) => {
  const raw = JSON.parse(fs.readFileSync(mappingPath, 'utf8'));
  const childFollowParentGroup = buildChildFollowParentPhrases(
    raw.child_follow_parent_group
  );
  const childFollowPhraseSet = new Set(childFollowParentGroup);

  return {
    childrens: filterRulesExcludingChildFollow(
      buildOrderedRules(raw.childrens),
      childFollowPhraseSet,
      'childrens'
    ),
    childFollowParentGroup,
    childFollowPhraseSet,
    parents: filterRulesExcludingChildFollow(
      buildOrderedRules(raw.parents),
      childFollowPhraseSet,
      'parents'
    ),
  };
};

/**
 * child_follow_parent_group: hanya inherit jika parent sudah punya equipment group.
 * @returns {{ value: string, tier: 'childFollowParent' } | null}
 */
const tryChildFollowParent = (haystack, funcLocNorm, egByFuncLoc, phrases) => {
  const parentFl = getDirectParentFuncLocNorm(funcLocNorm);
  if (!parentFl) return null;

  const parentEg = normalizeEquipmentGroupValue(egByFuncLoc.get(parentFl) ?? '');
  if (!parentEg) return null;

  for (const phrase of phrases ?? []) {
    if (haystack.includes(phrase)) {
      return { value: parentEg, tier: 'childFollowParent' };
    }
  }

  return null;
};

const resolveRowEquipmentGroup = (row, egByFuncLoc, mapping) => {
  if (isStationDepthFuncLoc(row.funcLocNorm)) {
    return { value: '', tier: 'station' };
  }

  const haystack = normalizeDescForMatch(row.desc);
  if (!haystack) return { value: '', tier: 'none' };

  for (const rule of mapping.childrens ?? []) {
    if (haystack.includes(rule.phrase)) {
      return { value: rule.value, tier: 'children' };
    }
  }

  const childFollow1 = tryChildFollowParent(
    haystack,
    row.funcLocNorm,
    egByFuncLoc,
    mapping.childFollowParentGroup
  );
  if (childFollow1) return childFollow1;

  for (const rule of mapping.parents ?? []) {
    if (haystack.includes(rule.phrase)) {
      return { value: rule.value, tier: 'parents' };
    }
  }

  const childFollow2 = tryChildFollowParent(
    haystack,
    row.funcLocNorm,
    egByFuncLoc,
    mapping.childFollowParentGroup
  );
  if (childFollow2) return childFollow2;

  return { value: '', tier: 'none' };
};

/**
 * @param {{ rowIndex: number, funcLocNorm: string, desc: string }[]} rows
 * @param {{ childrens?: { phrase: string, value: string }[], childFollowParentGroup?: string[], parents?: { phrase: string, value: string }[] }} mapping
 */
export const resolveEquipmentGroupAssignmentsForRows = (rows, mapping) => {
  const sorted = [...rows].sort((a, b) => {
    const depthDiff = getFuncLocDepth(a.funcLocNorm) - getFuncLocDepth(b.funcLocNorm);
    if (depthDiff !== 0) return depthDiff;
    const funcLocDiff = String(a.funcLocNorm ?? '').localeCompare(String(b.funcLocNorm ?? ''));
    if (funcLocDiff !== 0) return funcLocDiff;
    return Number(a.rowIndex ?? 0) - Number(b.rowIndex ?? 0);
  });

  const egByFuncLoc = new Map();
  const byRowIndex = new Map();
  for (const row of sorted) {
    const resolved = resolveRowEquipmentGroup(row, egByFuncLoc, mapping);
    if (row.funcLocNorm) egByFuncLoc.set(row.funcLocNorm, resolved.value);
    byRowIndex.set(row.rowIndex, resolved);
  }
  return byRowIndex;
};

/**
 * @param {string} desc
 * @param {{ childrens?: { phrase: string, value: string }[], childFollowParentGroup?: string[], parents?: { phrase: string, value: string }[] }} mapping
 */
export const resolveEquipmentGroupFromDescWithTier = (desc, mapping) =>
  resolveRowEquipmentGroup({ desc, funcLocNorm: '' }, new Map(), mapping);

export const resolveEquipmentGroupFromDesc = (desc, mapping) =>
  resolveEquipmentGroupFromDescWithTier(desc, mapping).value;

export { normalizeEquipmentGroupValue };
