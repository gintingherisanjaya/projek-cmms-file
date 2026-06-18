import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'node:url';

import {
  excelJsValueToPlainText,
  findHeadersExcelJs,
  normalizeHeader,
  stripFormulasFromWorksheet,
  toCleanString,
} from './utils/excelHelpers.js';
import { normalizeDescEqktu } from './utils/validationHelpers.js';
import {
  downloadFile,
  extractFolderIdFromUrl,
  getOrCreateFolder,
  initGoogleDrive,
  listFilesRecursively,
  uploadXlsxReplacing,
} from './utils/googleDrive.js';
import { filterDriveSourceExcelFiles } from './utils/driveSourceFileFilter.js';
import { createTempDir, cleanupTempDir, runWithConcurrency } from './utils/concurrencyHelpers.js';
import { DEFAULT_SOURCE_DRIVE_FOLDER_URL } from './utils/config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const FUNCLOC_HEADER = 'FUNCTIONAL LOCATION AFTER';
const FUNCLOC_DESC_HEADERS = ['FUNCTLOC DESC. AFTER LEVEL 1,2,3', 'FUNCTLOC DESC AFTRER LEVEL 1,2,3'];
const EQKTU_HEADERS = ['EQKTU AFTER LEVEL 2 dan 3', 'EQKTU AFTER LEVEL 2 AND 3'];
const EQUIPMENT_GROUP_AFTER_HEADER = 'EQUIPMENT GROUP AFTER';
const DEFAULT_OUT_DIR = path.join(__dirname, 'Output', 'BahJambi_Level5');
const DEFAULT_FIX_MAPPING_PATH = path.join(__dirname, 'mappings', 'bah-jambi-fix-mapping.json');
const OAUTH_JSON = path.join(__dirname, 'oauth.json');

const DEFAULT_DRIVE_FOLDER_URL = DEFAULT_SOURCE_DRIVE_FOLDER_URL;

const DEFAULT_CONCURRENCY = Math.max(
  1,
  Math.min(8, os.availableParallelism?.() ?? os.cpus().length ?? 8),
);

const safeSlug = (value) =>
  String(value ?? '')
    .trim()
    .replace(/[^\w.-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 80) || 'file';

const isRegionalPath = (filePath) => {
  const firstSegment = String(filePath ?? '')
    .split('/')
    .find(Boolean);
  return /^REGIONAL/i.test(String(firstSegment || '').trim());
};

/** Built-in FUNCLOC moves merged before fix-map (webapp no longer applies bundled defaults). */
const BUILTIN_FUNCL_LOC_MOVES = [
  {
    beforeFuncLoc: 'PALM-2F01-0005-0007-00027-0004-0001',
    afterFuncLoc: 'PALM-2F01-0005-0007-0027-0004-0001',
  },
  {
    beforeFuncLoc: 'PALM-2F01-0005-0007-00027-0004-0002',
    afterFuncLoc: 'PALM-2F01-0005-0007-0027-0004-0002',
  },
];

const COLOR_HEADER = 'FF1F2937';
const COLOR_HEADER_TEXT = 'FFFFFFFF';
const COLOR_WINNER = 'FFC6EFCE';
const COLOR_TIE_SAME = 'FFDBEAFE';
const COLOR_TIE_DIFF = 'FFFFF2CC';
const COLOR_DISABLED = 'FFF2F2F2';
const COLOR_DISABLED_TEXT = 'FF9CA3AF';
const COLOR_ENABLED = 'FFD9EAD3';

const parseArgs = () => {
  const args = process.argv.slice(2);
  let driveFolderUrl = DEFAULT_DRIVE_FOLDER_URL;
  let outDir = DEFAULT_OUT_DIR;
  let fixMapPath = DEFAULT_FIX_MAPPING_PATH;
  let help = false;

  for (let i = 0; i < args.length; i += 1) {
    const a = args[i];
    if (a === '--source' && args[i + 1]) {
      driveFolderUrl = args[i + 1];
      i += 1;
    } else if (a === '--out' && args[i + 1]) {
      outDir = args[i + 1];
      i += 1;
    } else if (a === '--fix-map' && args[i + 1]) {
      fixMapPath = args[i + 1];
      i += 1;
    } else if (a === '--help' || a === '-h') {
      help = true;
    }
  }

  return {
    mode: 'drive-folder',
    driveFolderUrl,
    outDir: path.isAbsolute(outDir) ? outDir : path.join(process.cwd(), outDir),
    fixMapPath: path.isAbsolute(fixMapPath) ? fixMapPath : path.join(process.cwd(), fixMapPath),
    help,
  };
};

const getJakartaStamp = () => {
  const fmt = new Intl.DateTimeFormat('sv-SE', {
    timeZone: 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  });
  const parts = fmt.formatToParts(new Date());
  const map = new Map(parts.map((p) => [p.type, p.value]));
  return `${map.get('year')}${map.get('month')}${map.get('day')}_${map.get('hour')}${map.get('minute')}${map.get('second')}`;
};

const ensureFolderPath = async (rootFolderId, relativeDir, folderCache) => {
  const normalized = String(relativeDir || '')
    .split('/')
    .filter(Boolean)
    .join('/');
  if (!normalized || normalized === '.') return rootFolderId;
  if (folderCache.has(normalized)) return folderCache.get(normalized);

  const segments = normalized.split('/');
  let parentId = rootFolderId;
  let built = '';
  for (const seg of segments) {
    built = built ? `${built}/${seg}` : seg;
    if (folderCache.has(built)) {
      parentId = folderCache.get(built);
      continue;
    }
    const folder = await getOrCreateFolder(seg, parentId);
    folderCache.set(built, folder.id);
    parentId = folder.id;
  }
  return parentId;
};

const getLevel = (funcLoc) => String(funcLoc).split('-').filter(Boolean).length;
const getParentKey = (funcLoc) => {
  const parts = String(funcLoc).split('-').filter(Boolean);
  if (parts.length <= 1) return '';
  return parts.slice(0, -1).join('-');
};

const extractNoNumbers = (text) => {
  const out = [];
  const re = /N[O0]\s*(?:\.\s*)*(\d+)/gi;
  let m;
  while ((m = re.exec(String(text ?? ''))) !== null) {
    const n = Number.parseInt(m[1], 10);
    if (Number.isFinite(n)) out.push(n);
  }
  return out;
};

const parseLastTrailingNumber = (value) => {
  const m = String(value || '')
    .trim()
    .match(/(\d+)\s*$/);
  if (!m) return null;
  const n = Number.parseInt(m[1], 10);
  return Number.isFinite(n) ? n : null;
};

const parseParentUnitNumber = (desc) => {
  const m = String(desc || '')
    .toUpperCase()
    .match(/NO\.?\s*(\d+)\s*$/);
  return m?.[1] ?? null;
};

/**
 * Raw strip of trailing digits only (legacy helper).
 * Prefer canonicalDescBaseForGrouping + canonicalSiblingGroupKey for analysis.
 */
const normalizeDescBase = (value) =>
  String(value || '')
    .trim()
    .replace(/\d+\s*$/, '')
    .replace(/\s+/g, ' ')
    .toUpperCase()
    .trim();

const loadFixMapping = (mappingPath) => {
  const empty = {
    strictMode: true,
    requireScopedDescriptionRules: false,
    descriptionReplacements: {},
    funcLocReplacements: {},
    scopedDescriptionRules: [],
    specialNormalizationRules: [],
    specialAdaptationRules: [],
    /** When true, clone left sibling subtree into a single missing 4-digit slot (see applyMiddleGapSubtreeFill). */
    middleGapFill: true,
    /**
     * When true, if a parent’s 4-digit children have exactly one missing integer in [min..max],
     * shift every direct child with suffix greater than that hole down by 1 (full subtrees), two-pass temp
     * IDs — like 1,2,3,5,6 → 1,2,3,4,5. Runs before middle gap fill.
     */
    suffixShiftCompact: true,
  };
  if (!mappingPath || !fs.existsSync(mappingPath)) {
    // Always fall back to default mapping so running without flags applies full logic.
    if (mappingPath !== DEFAULT_FIX_MAPPING_PATH && fs.existsSync(DEFAULT_FIX_MAPPING_PATH)) {
      mappingPath = DEFAULT_FIX_MAPPING_PATH;
    } else {
      return { mapping: empty, mappingPath };
    }
  }
  try {
    const raw = JSON.parse(fs.readFileSync(mappingPath, 'utf8'));
    return {
      mapping: {
        descriptionReplacements: raw?.descriptionReplacements ?? {},
        funcLocReplacements: raw?.funcLocReplacements ?? {},
        scopedDescriptionRules: Array.isArray(raw?.scopedDescriptionRules)
          ? raw.scopedDescriptionRules
          : [],
        specialNormalizationRules: Array.isArray(raw?.specialNormalizationRules)
          ? raw.specialNormalizationRules
          : [],
        specialAdaptationRules: Array.isArray(raw?.specialAdaptationRules)
          ? raw.specialAdaptationRules
          : [],
        strictMode: raw?.strictMode !== false,
        requireScopedDescriptionRules: raw?.requireScopedDescriptionRules === true,
        middleGapFill: raw?.middleGapFill !== false,
        suffixShiftCompact: raw?.suffixShiftCompact !== false,
      },
      mappingPath,
    };
  } catch {
    return { mapping: empty, mappingPath };
  }
};

const loadBuiltinFuncLocMoves = () =>
  BUILTIN_FUNCL_LOC_MOVES.map((m) => ({
    beforeFuncLoc: m.beforeFuncLoc.toUpperCase(),
    afterFuncLoc: m.afterFuncLoc.toUpperCase(),
  }));

/**
 * Merge built-in FUNCLOC moves into `funcLocReplacements`. Explicit Bah Jambi mapping keys win on collision.
 * Pass `draftMoves` to avoid recomputing when you already called `loadBuiltinFuncLocMoves`.
 */
const mergeBuiltinFuncLocMovesIntoMapping = (mapping, draftMoves = loadBuiltinFuncLocMoves()) => {
  const fromDraft = {};
  for (const m of draftMoves) {
    fromDraft[m.beforeFuncLoc] = m.afterFuncLoc;
  }
  return {
    ...mapping,
    funcLocReplacements: {
      ...fromDraft,
      ...(mapping?.funcLocReplacements ?? {}),
    },
  };
};

/**
 * Apply FUNCLOC replacements on row objects so `analyzeLevel5` sees the same tree as the written workbook.
 */
const applyFuncLocReplacementsToRowsInPlace = (rows, funcLocReplacements) => {
  const map = funcLocReplacements ?? {};
  for (const r of rows) {
    const next = map[r.funcLocAfter];
    if (!next || next === r.funcLocAfter) continue;
    r.funcLocAfter = next;
    r.parentKey = getParentKey(next);
    r.level = getLevel(next);
  }
};

const escapeRegex = (value) => String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const replaceTokenStrict = (input, from, to) => {
  const src = String(input ?? '');
  const find = String(from ?? '').trim().toUpperCase();
  if (!find) return { next: src, changed: false };
  const out = String(to ?? '').toUpperCase();
  // Strict boundary replacement: avoid accidental partial-word substitutions.
  const re = new RegExp(`(^|[^A-Z0-9])(${escapeRegex(find)})(?=[^A-Z0-9]|$)`, 'g');
  let changed = false;
  const next = src.replace(re, (m, p1) => {
    changed = true;
    return `${p1}${out}`;
  });
  return { next, changed };
};

/**
 * Normalize description for sibling grouping (level-5 “same name family”):
 * - NFKC + collapse whitespace (incl. NBSP)
 * - normalize NO / NR token shapes so NO.4 vs NO 4 align
 * - strip trailing numeric instance (same idea as Bah Jambi prune / web dendrogram)
 */
const canonicalDescBaseForGrouping = (value) => {
  let t = String(value ?? '')
    .normalize('NFKC')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
  t = t.replace(/\bNO\.?\b/g, 'NO').replace(/\bNR\.?\b/g, 'NO');
  t = t.replace(/\d+\s*$/, '').trim();
  return t.replace(/\s+/g, ' ').trim();
};

/** Descriptions that mention LORY are grouped separately (do not compete with non-LORY siblings). */
const isLoryTaggedDesc = (desc) => /\blory\b/i.test(String(desc ?? '').trim());

/**
 * Apply global descriptionReplacements (strict token) before base normalization so mapping can merge known variants.
 */
const applyReplacementsForGroupingKey = (desc, mapping) => {
  let normalized = String(desc ?? '')
    .normalize('NFKC')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
  const replacements = mapping?.descriptionReplacements ?? {};
  for (const [from, to] of Object.entries(replacements)) {
    if (!from) continue;
    const { next } =
      mapping?.strictMode !== false
        ? replaceTokenStrict(normalized, from, to)
        : (() => {
            const re = new RegExp(escapeRegex(String(from).toUpperCase()), 'g');
            const after = normalized.replace(re, String(to).toUpperCase());
            return { next: after, changed: after !== normalized };
          })();
    normalized = next;
  }
  return canonicalDescBaseForGrouping(normalized);
};

/**
 * Stable key for level-5 sibling groups: same parent FUNCLOC + canonical desc base + LORY lane.
 */
const canonicalSiblingGroupKey = (parentKey, desc, mapping) => {
  const base = applyReplacementsForGroupingKey(desc, mapping);
  const lane = isLoryTaggedDesc(desc) ? 'LORY' : 'STD';
  return `${parentKey}||${base}||${lane}`;
};

/** Same prefix replace idea as webapp `replaceFuncLocPrefix` in master-data-compact-gaps. */
const replaceFuncLocPrefixForGapFill = (full, oldRoot, newRoot) => {
  const f = String(full).trim().toUpperCase();
  const o = String(oldRoot).trim().toUpperCase();
  const n = String(newRoot).trim().toUpperCase();
  if (f === o) return n;
  if (f.startsWith(`${o}-`)) return `${n}${f.slice(o.length)}`;
  return f;
};

const bumpTrailingNoForGapFill = (desc, delta) => {
  const s = String(desc ?? '').trim();
  const m = s.match(/^(.*?)(\bNO\.?\s*)(\d+)(\s*)$/i);
  if (!m) return null;
  const num = Number.parseInt(m[3], 10) + delta;
  if (!Number.isFinite(num) || num < 1) return null;
  return m[1] + m[2] + String(num) + m[4];
};

const normalizeGapFillDesc = (d) =>
  normalizeDescEqktu(String(d ?? ''))
    .replace(/\s+/g, ' ')
    .trim();

const replaceRootAnchorInDesc = (sourceDesc, oldAnchor, newAnchor) => {
  const src = normalizeGapFillDesc(sourceDesc);
  const oldN = normalizeGapFillDesc(oldAnchor);
  const newN = normalizeGapFillDesc(newAnchor);
  if (!oldN || oldN === newN) return src;
  const flex = escapeRegex(oldN).replace(/ /g, '\\s+');
  return src.replace(new RegExp(flex, 'gi'), newN);
};

const collectDirectNumericChildren = (rows, parentP) => {
  const out = [];
  for (const r of rows) {
    if (r.parentKey !== parentP) continue;
    const parts = String(r.funcLocAfter ?? '')
      .split('-')
      .filter(Boolean);
    const seg = parts[parts.length - 1];
    if (!/^\d{4}$/.test(seg)) continue;
    out.push({ n: Number.parseInt(seg, 10), r });
  }
  out.sort((a, b) => a.n - b.n);
  return out;
};

const suffix4ForShift = (n) => String(Number(n)).padStart(4, '0');

/**
 * If direct numeric children under P have exactly one missing index h in [min..max], shift every
 * n greater than h down by 1 on FUNCLOC (entire subtrees). Two-pass TMP roots avoid collisions (same idea as
 * webapp compact gaps). Skips when any mover is LORY-tagged.
 */
const applySingleHoleSuffixShiftCompact = ({ worksheet, rows, header, mapping }) => {
  if (mapping?.suffixShiftCompact === false) {
    return { operations: [], updatedRowCount: 0, skippedReason: 'suffixShiftCompact disabled in mapping' };
  }

  const operations = [];
  const parentKeys = [...new Set(rows.map((r) => r.parentKey).filter(Boolean))].sort();

  for (const P of parentKeys) {
    const direct = collectDirectNumericChildren(rows, P);
    if (direct.length < 2) continue;

    const present = new Set(direct.map((d) => d.n));
    const min = Math.min(...[...present]);
    const max = Math.max(...[...present]);
    const missing = [];
    for (let k = min; k <= max; k += 1) {
      if (!present.has(k)) missing.push(k);
    }
    if (missing.length !== 1) continue;
    const h = missing[0];

    const movers = direct.filter((d) => d.n > h);
    if (movers.length === 0) continue;
    if (movers.some((m) => isLoryTaggedDesc(m.r.desc))) continue;

    const toTemp = new Map();
    const tempToNew = new Map();
    movers.sort((a, b) => a.n - b.n);
    movers.forEach((m, idx) => {
      const oldRoot = `${P}-${suffix4ForShift(m.n)}`;
      const newRoot = `${P}-${suffix4ForShift(m.n - 1)}`;
      const temp = `${P}-TMP-SUF-${String(idx).padStart(4, '0')}`;
      toTemp.set(oldRoot, temp);
      tempToNew.set(temp, newRoot);
    });

    const toTempEntries = [...toTemp.entries()].sort((a, b) => b[0].length - a[0].length);
    const tempToNewEntries = [...tempToNew.entries()].sort((a, b) => b[0].length - a[0].length);

    const affected = rows.filter((r) => {
      const fl = r.funcLocAfter;
      for (const oldRoot of toTemp.keys()) {
        if (fl === oldRoot || fl.startsWith(`${oldRoot}-`)) return true;
      }
      return false;
    });

    for (const r of affected) {
      let next = r.funcLocAfter;
      for (const [o, t] of toTempEntries) {
        if (next === o || next.startsWith(`${o}-`)) {
          next = replaceFuncLocPrefixForGapFill(next, o, t);
        }
      }
      for (const [t, nRoot] of tempToNewEntries) {
        if (next === t || next.startsWith(`${t}-`)) {
          next = replaceFuncLocPrefixForGapFill(next, t, nRoot);
        }
      }
      r.funcLocAfter = next;
      r.parentKey = getParentKey(next);
      r.level = getLevel(next);
      r.descBase = canonicalDescBaseForGrouping(r.desc);
      r.descNo = parseLastTrailingNumber(r.desc);

      const excelRow = worksheet.getRow(r.rowNumber);
      excelRow.getCell(header.funclocCol).value = next;
    }

    operations.push({
      parentFuncLoc: P,
      holeNumericSuffix: h,
      holeSuffix: suffix4ForShift(h),
      shiftedDirectChildCount: movers.length,
      affectedRowCount: affected.length,
    });
  }

  const updatedRowCount = operations.reduce((s, o) => s + o.affectedRowCount, 0);
  if (updatedRowCount > 0) {
    console.log(
      `Single-hole suffix shift: ${operations.length} parent(s), ${updatedRowCount} row(s) FUNCLOC shifted (n > hole -> n-1; two-pass).`,
    );
  }

  return { operations, updatedRowCount };
};

/**
 * Webapp “compact gaps” renumbers existing siblings; this complements it by **inserting** a missing
 * middle slot when there is exactly one integer gap between two neighbours: clone the **left** sibling’s
 * entire subtree to the empty suffix, remap FUNCLOC prefixes, bump the template root’s trailing `NO. k`,
 * replace that phrase in subtree descriptions, and copy EQUIPMENT GROUP AFTER from each source row.
 *
 * Special case: left root is `NO. n` and right neighbour root is `NO. n+1` (descriptions already
 * consecutive) but FUNCLOC skips one suffix — the hole is filled with **`NO. n+2`** (e.g. …0093 NO.1,
 * …0095 NO.2 → insert …0094 as NO.3), not `NO. n+1` (which would duplicate the right line).
 *
 * Multi-slot: if several integers are missing between left and right **and** the two neighbour roots
 * are different equipment families (canonical desc base), clone the **left** subtree once per
 * missing suffix with `NO. k + (g - n_left)` (e.g. pumps NO.2–4 then a gap then a pond — fill 0034…0038
 * as NO.5…NO.9).
 */
const equipmentFamiliesDifferForGapFill = (descLeft, descRight) =>
  canonicalDescBaseForGrouping(descLeft) !== canonicalDescBaseForGrouping(descRight);

const applyMiddleGapSubtreeFill = ({ worksheet, rows, header, mapping }) => {
  if (mapping?.middleGapFill === false) {
    return { insertedRowCount: 0, fills: [], skippedReason: 'middleGapFill disabled in mapping' };
  }

  const fills = [];
  let insertedRowCount = 0;
  const existingFunc = new Set(rows.map((r) => r.funcLocAfter));
  const parentKeys = [...new Set(rows.map((r) => r.parentKey).filter(Boolean))].sort();

  for (const P of parentKeys) {
    let progressed = true;
    while (progressed) {
      progressed = false;
      const direct = collectDirectNumericChildren(rows, P);
      if (direct.length < 2) break;

      for (let i = 0; i < direct.length - 1; i += 1) {
        const na = direct[i].n;
        const nb = direct[i + 1].n;
        if (nb - na < 2) continue;

        const templateRootRow = direct[i].r;
        const rightRootRow = direct[i + 1].r;
        const templateRoot = templateRootRow.funcLocAfter;
        if (isLoryTaggedDesc(templateRootRow.desc)) continue;

        const subtree = rows
          .filter(
            (x) =>
              x.funcLocAfter === templateRoot ||
              x.funcLocAfter.startsWith(`${templateRoot}-`),
          )
          .sort(
            (a, b) =>
              a.funcLocAfter.length - b.funcLocAfter.length || a.funcLocAfter.localeCompare(b.funcLocAfter),
          );

        if (subtree.length === 0) continue;

        const oldAnchor = normalizeGapFillDesc(templateRootRow.desc);
        const uLStr = parseParentUnitNumber(templateRootRow.desc);
        const uL = uLStr != null ? Number.parseInt(String(uLStr), 10) : Number.NaN;

        const appendOneGap = (g, delta) => {
          const newAnchorRaw = bumpTrailingNoForGapFill(templateRootRow.desc, delta);
          if (!newAnchorRaw) return false;
          const newAnchor = normalizeGapFillDesc(newAnchorRaw);
          if (!newAnchor || newAnchor === oldAnchor) return false;

          const newRootId = `${P}-${String(g).padStart(4, '0')}`;
          if (existingFunc.has(newRootId)) return false;

          for (const sr of subtree) {
            const nf = replaceFuncLocPrefixForGapFill(sr.funcLocAfter, templateRoot, newRootId);
            if (existingFunc.has(nf)) return false;
          }

          for (const sr of subtree) {
            const newFuncLoc = replaceFuncLocPrefixForGapFill(sr.funcLocAfter, templateRoot, newRootId);
            const newDesc = normalizeDescEqktu(replaceRootAnchorInDesc(sr.desc, oldAnchor, newAnchor));
            const eqG = sr.equipmentGroupAfter != null ? String(sr.equipmentGroupAfter) : '';

            const newRow = worksheet.addRow([]);
            const tmpl = worksheet.getRow(sr.rowNumber);
            if (typeof tmpl.height === 'number' && tmpl.height > 0) {
              newRow.height = tmpl.height;
            }
            tmpl.eachCell({ includeEmpty: true }, (cell, col) => {
              const c = newRow.getCell(col);
              c.style = { ...cell.style };
            });

            newRow.getCell(header.funclocCol).value = newFuncLoc;
            newRow.getCell(header.descCol).value = newDesc;
            if (header.eqktuCol != null) {
              newRow.getCell(header.eqktuCol).value = newDesc;
            }
            if (header.equipmentGroupAfterCol != null) {
              newRow.getCell(header.equipmentGroupAfterCol).value = eqG;
            }

            rows.push({
              rowNumber: newRow.number,
              funcLocAfter: newFuncLoc,
              desc: newDesc,
              equipmentGroupAfter: eqG,
              level: getLevel(newFuncLoc),
              parentKey: getParentKey(newFuncLoc),
              descBase: canonicalDescBaseForGrouping(newDesc),
              descNo: parseLastTrailingNumber(newDesc),
            });
            existingFunc.add(newFuncLoc);
            insertedRowCount += 1;
          }

          fills.push({
            parentFuncLoc: P,
            gapSuffix: String(g).padStart(4, '0'),
            templateRootFuncLoc: templateRoot,
            newRootFuncLoc: newRootId,
            clonedRowCount: subtree.length,
          });
          return true;
        };

        if (nb - na === 2) {
          const g = na + 1;
          const uRStr = parseParentUnitNumber(rightRootRow.desc);
          const uR = uRStr != null ? Number.parseInt(String(uRStr), 10) : Number.NaN;
          let delta = g - na;
          if (Number.isFinite(uL) && Number.isFinite(uR) && uR === uL + 1) {
            delta = uR + 1 - uL;
          }
          if (appendOneGap(g, delta)) {
            progressed = true;
            break;
          }
          continue;
        }

        if (nb - na > 2) {
          if (!equipmentFamiliesDifferForGapFill(templateRootRow.desc, rightRootRow.desc)) continue;
          if (!Number.isFinite(uL)) continue;

          const steps = [];
          for (let g = na + 1; g <= nb - 1; g += 1) {
            const delta = g - na;
            const newAnchorRaw = bumpTrailingNoForGapFill(templateRootRow.desc, delta);
            if (!newAnchorRaw) {
              steps.length = 0;
              break;
            }
            const newAnchor = normalizeGapFillDesc(newAnchorRaw);
            if (!newAnchor || newAnchor === oldAnchor) {
              steps.length = 0;
              break;
            }
            const newRootId = `${P}-${String(g).padStart(4, '0')}`;
            if (existingFunc.has(newRootId)) {
              steps.length = 0;
              break;
            }
            let hit = false;
            for (const sr of subtree) {
              const nf = replaceFuncLocPrefixForGapFill(sr.funcLocAfter, templateRoot, newRootId);
              if (existingFunc.has(nf)) {
                hit = true;
                break;
              }
            }
            if (hit) {
              steps.length = 0;
              break;
            }
            steps.push({ g, delta });
          }

          const expectedSteps = nb - na - 1;
          if (steps.length !== expectedSteps) continue;

          let ok = true;
          for (const { g, delta } of steps) {
            if (!appendOneGap(g, delta)) {
              ok = false;
              break;
            }
          }
          if (ok) {
            progressed = true;
            break;
          }
        }
      }
    }
  }

  if (insertedRowCount > 0) {
    console.log(
      `Middle gap subtree fill: ${fills.length} root gap(s), ${insertedRowCount} row(s) appended (single- or multi-slot; left template; different family on right for multi; EQUIPMENT GROUP copied).`,
    );
  }

  return { insertedRowCount, fills };
};

const bumpUsage = (usage, bucket, key) => {
  if (!usage || !bucket || !key) return;
  if (!usage[bucket]) usage[bucket] = {};
  usage[bucket][key] = (usage[bucket][key] ?? 0) + 1;
};

const applySpecialNormalizationRules = (normalized, parentDescNorm, mapping, context, mappingUsage) => {
  let out = normalized;
  const rules = mapping?.specialNormalizationRules ?? [];
  for (const rule of rules) {
    const id = String(rule?.id ?? '').trim();
    const type = String(rule?.type ?? '').trim();
    if (!id || !type) continue;

    if (type === 'inject_parent_marker_if_missing') {
      const parentContains = String(rule?.parentDescContains ?? '')
        .trim()
        .toUpperCase();
      const extractParentMarkerRegex = String(rule?.extractParentMarkerRegex ?? '').trim();
      const targetRegex = String(rule?.targetRegex ?? '').trim();
      const skipIfContains = String(rule?.skipIfContains ?? '')
        .trim()
        .toUpperCase();
      const replaceTemplate = String(rule?.replaceTemplate ?? '').trim();
      if (!parentContains || !extractParentMarkerRegex || !targetRegex || !replaceTemplate) continue;
      if (!parentDescNorm.includes(parentContains)) continue;
      if (skipIfContains && out.includes(skipIfContains)) continue;

      let marker = null;
      try {
        const m = new RegExp(extractParentMarkerRegex, 'i').exec(parentDescNorm);
        marker = m?.[1] ?? null;
      } catch {
        marker = null;
      }
      if (!marker) continue;

      try {
        const re = new RegExp(targetRegex, 'ig');
        const next = out.replace(re, replaceTemplate.replaceAll('{marker}', marker));
        if (next !== out) {
          out = next;
          bumpUsage(mappingUsage, 'specialNormalizationRuleHits', id);
        }
      } catch {}
    }
  }
  return out;
};

const applySpecialAdaptationRules = (
  adaptedDesc,
  winnerParentDesc,
  loserParentDesc,
  mapping,
  context,
  mappingUsage,
) => {
  let out = adaptedDesc;
  const rules = mapping?.specialAdaptationRules ?? [];
  const winnerParentNorm = String(winnerParentDesc || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
  const loserParentNorm = String(loserParentDesc || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
  const loserNo = parseParentUnitNumber(loserParentNorm);

  for (const rule of rules) {
    const id = String(rule?.id ?? '').trim();
    const type = String(rule?.type ?? '').trim();
    if (!id || !type) continue;

    if (type === 'replace_no_with_loser_parent_no') {
      const parentDescContains = String(rule?.parentDescContains ?? '')
        .trim()
        .toUpperCase();
      const targetRegex = String(rule?.targetRegex ?? '').trim();
      if (!parentDescContains || !targetRegex || !loserNo) continue;
      if (!winnerParentNorm.includes(parentDescContains) || !loserParentNorm.includes(parentDescContains))
        continue;
      try {
        const re = new RegExp(targetRegex, 'ig');
        const next = out.replace(re, `$1${loserNo}`);
        if (next !== out) {
          out = next;
          bumpUsage(mappingUsage, 'specialAdaptationRuleHits', id);
        }
      } catch {}
    }
  }

  return out;
};

const normalizeChildDescForCompare = (
  childDesc,
  level5Desc,
  mapping,
  context = {},
  mappingUsage = null,
) => {
  let normalized = String(childDesc || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();

  const parentDescNorm = String(level5Desc || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();

  // Keep item numbering (e.g. ELECTROMOTOR NO.1/2/3), but neutralize the
  // parent unit numbering segment (e.g. HOISTING CRANE NO.1 vs NO.2).
  const parentStem = parentDescNorm.replace(/\s*NO\.?\s*\d+\s*$/, '').trim();
  if (parentStem) {
    const parentNoRegex = new RegExp(`${escapeRegex(parentStem)}\\s+NO\\.?\\s*\\d+`, 'g');
    normalized = normalized.replace(parentNoRegex, `${parentStem} NO`);
  }

  // Parent-aware structure bollard normalization:
  // Mapping-driven special normalization rules.
  normalized = applySpecialNormalizationRules(normalized, parentDescNorm, mapping, context, mappingUsage);

  const scopedRules = mapping?.scopedDescriptionRules ?? [];
  let scopedMatched = false;
  for (const rule of scopedRules) {
    const id = String(rule?.id ?? '').trim();
    const find = String(rule?.find ?? '');
    const replace = String(rule?.replace ?? '');
    const skipIfContains = String(rule?.skipIfContains ?? '').trim().toUpperCase();
    if (!id || !find) continue;

    const parentPrefix = String(rule?.parentPrefix ?? '').trim();
    const parentDescContains = String(rule?.parentDescContains ?? '').trim().toUpperCase();
    const childSuffix = String(rule?.childSuffix ?? '').trim();

    if (parentPrefix && !String(context.level5FuncLoc ?? '').startsWith(parentPrefix)) continue;
    if (parentDescContains && !parentDescNorm.includes(parentDescContains)) continue;
    if (childSuffix && String(context.childSuffix ?? '') !== childSuffix) continue;
    if (skipIfContains && normalized.includes(skipIfContains)) continue;

    const { next, changed } =
      mapping?.strictMode !== false
        ? replaceTokenStrict(normalized, find, replace)
        : (() => {
            const re = new RegExp(escapeRegex(find.toUpperCase()), 'g');
            const before = normalized;
            const after = normalized.replace(re, replace.toUpperCase());
            return { next: after, changed: after !== before };
          })();
    normalized = next;
    if (changed) {
      scopedMatched = true;
      bumpUsage(mappingUsage, 'scopedDescriptionRuleHits', id);
    }
  }

  const allowGlobalReplacements =
    mapping?.requireScopedDescriptionRules !== true || scopedMatched;
  if (allowGlobalReplacements) {
    const replacements = mapping?.descriptionReplacements ?? {};
    for (const [from, to] of Object.entries(replacements)) {
      if (!from) continue;
      const { next, changed } =
        mapping?.strictMode !== false
          ? replaceTokenStrict(normalized, from, to)
          : (() => {
              const re = new RegExp(escapeRegex(String(from).toUpperCase()), 'g');
              const before = normalized;
              const after = normalized.replace(re, String(to).toUpperCase());
              return { next: after, changed: after !== before };
            })();
      normalized = next;
      if (changed) bumpUsage(mappingUsage, 'descriptionReplacementHits', String(from));
    }
  }

  return normalized
    .replace(/\bNO\./g, 'NO')
    .replace(/\d+\s*$/, '')
    .replace(/\s+/g, ' ')
    .trim();
};

const getRelativeChildSuffix = (childFuncLoc, level5FuncLoc) => {
  const childParts = String(childFuncLoc || '').split('-').filter(Boolean);
  const rootParts = String(level5FuncLoc || '').split('-').filter(Boolean);
  if (childParts.length <= rootParts.length) return '';
  return childParts.slice(rootParts.length).join('-');
};

const normalizeChildSignature = (
  childFuncLoc,
  level5FuncLoc,
  desc,
  level5Desc,
  mapping,
  mappingUsage = null,
) => {
  let normalizedChildFuncLoc = String(childFuncLoc || '');
  const funcLocMap = mapping?.funcLocReplacements ?? {};
  if (funcLocMap[normalizedChildFuncLoc]) {
    bumpUsage(mappingUsage, 'funcLocReplacementHits', normalizedChildFuncLoc);
    normalizedChildFuncLoc = funcLocMap[normalizedChildFuncLoc];
  }
  const suffix = getRelativeChildSuffix(normalizedChildFuncLoc, level5FuncLoc);
  const descKey = normalizeChildDescForCompare(
    desc,
    level5Desc,
    mapping,
    { level5FuncLoc, childFuncLoc: normalizedChildFuncLoc, childSuffix: suffix },
    mappingUsage,
  );
  // Use both relative FUNCLOC path and normalized desc to compare sibling children.
  return `${suffix}||${descKey}`;
};

const adaptWinnerChildDescToLoserParent = (winnerChildDesc, winnerParentDesc, loserParentDesc) => {
  let out = String(winnerChildDesc || '');
  const winnerParentNorm = String(winnerParentDesc || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
  const loserParentNorm = String(loserParentDesc || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
  const winnerNo = parseParentUnitNumber(winnerParentNorm);
  const loserNo = parseParentUnitNumber(loserParentNorm);
  const parentStem = winnerParentNorm.replace(/\s*NO\.?\s*\d+\s*$/, '').trim();
  if (!winnerNo || !loserNo || !parentStem || winnerNo === loserNo) return out;
  const parentWithWinnerNo = new RegExp(
    `${escapeRegex(parentStem)}\\s+NO\\.?\\s*${escapeRegex(winnerNo)}\\b`,
    'ig',
  );
  out = out.replace(parentWithWinnerNo, `${parentStem} NO.${loserNo}`);

  return out;
};

const pickWinnerByLowestNumber = (siblings) =>
  [...siblings].sort((a, b) => {
    const an = parseLastTrailingNumber(a.desc);
    const bn = parseLastTrailingNumber(b.desc);
    if (an == null && bn == null) return a.funcLocAfter.localeCompare(b.funcLocAfter);
    if (an == null) return 1;
    if (bn == null) return -1;
    return an - bn || a.funcLocAfter.localeCompare(b.funcLocAfter);
  })[0];

const loadWorkbookRows = async (excelPath) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('No worksheet found in source workbook.');

  // Replace every formula with its cached result before header detection and row reads
  // so analysis and output use plain values from master data, not formula objects.
  stripFormulasFromWorksheet(worksheet);

  const found = findHeadersExcelJs(
    worksheet,
    [FUNCLOC_HEADER, ...FUNCLOC_DESC_HEADERS, ...EQKTU_HEADERS, EQUIPMENT_GROUP_AFTER_HEADER],
    {
      // Scan the whole sheet for the header row (not only the first N rows).
      maxScanRows: worksheet.rowCount,
      minSimilarity: 0.6,
    },
  );
  if (!found) throw new Error('Could not find FUNCTIONAL LOCATION AFTER and FUNCTLOC DESC columns.');

  const funclocCol = found.colMap.get(normalizeHeader(FUNCLOC_HEADER));
  const descCol =
    found.colMap.get(normalizeHeader(FUNCLOC_DESC_HEADERS[0])) ??
    found.colMap.get(normalizeHeader(FUNCLOC_DESC_HEADERS[1]));
  const eqktuCol =
    found.colMap.get(normalizeHeader(EQKTU_HEADERS[0])) ??
    found.colMap.get(normalizeHeader(EQKTU_HEADERS[1]));
  const equipmentGroupAfterCol = found.colMap.get(normalizeHeader(EQUIPMENT_GROUP_AFTER_HEADER));
  if (funclocCol == null || descCol == null) {
    throw new Error('Required columns are missing from Bah Jambi sheet.');
  }

  const rows = [];
  const seen = new Set();
  for (let r = found.headerRowNumber + 1; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const funcLocAfter = toCleanString(excelJsValueToPlainText(row.getCell(funclocCol).value)).toUpperCase();
    const desc = normalizeDescEqktu(excelJsValueToPlainText(row.getCell(descCol).value));
    if (!funcLocAfter) continue;
    const key = `${funcLocAfter}||${desc}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const equipmentGroupAfter =
      equipmentGroupAfterCol != null
        ? excelJsValueToPlainText(row.getCell(equipmentGroupAfterCol).value)
        : '';
    rows.push({
      rowNumber: r,
      funcLocAfter,
      desc,
      equipmentGroupAfter,
      level: getLevel(funcLocAfter),
      parentKey: getParentKey(funcLocAfter),
      descBase: canonicalDescBaseForGrouping(desc),
      descNo: parseLastTrailingNumber(desc),
    });
  }

  return {
    workbook,
    worksheet,
    rows,
    header: {
      funclocCol,
      descCol,
      eqktuCol,
      equipmentGroupAfterCol,
      headerRowNumber: found.headerRowNumber,
    },
  };
};

const applyDescMappingForOutput = (rawDesc, parentDesc, mapping, context = {}, mappingUsage = null) => {
  let text = String(rawDesc ?? '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
  const parentDescNorm = String(parentDesc ?? '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();

  text = applySpecialNormalizationRules(text, parentDescNorm, mapping, context, mappingUsage);

  const scopedRules = mapping?.scopedDescriptionRules ?? [];
  let scopedMatched = false;
  for (const rule of scopedRules) {
    const id = String(rule?.id ?? '').trim();
    const find = String(rule?.find ?? '');
    const replace = String(rule?.replace ?? '');
    if (!id || !find) continue;
    const parentPrefix = String(rule?.parentPrefix ?? '').trim();
    const parentDescContains = String(rule?.parentDescContains ?? '').trim().toUpperCase();
    const childSuffix = String(rule?.childSuffix ?? '').trim();
    const skipIfContains = String(rule?.skipIfContains ?? '').trim().toUpperCase();
    if (parentPrefix && !String(context.level5FuncLoc ?? '').startsWith(parentPrefix)) continue;
    if (parentDescContains && !parentDescNorm.includes(parentDescContains)) continue;
    if (childSuffix && String(context.childSuffix ?? '') !== childSuffix) continue;
    if (skipIfContains && text.includes(skipIfContains)) continue;
    const { next, changed } = replaceTokenStrict(text, find, replace);
    text = next;
    if (changed) {
      scopedMatched = true;
      bumpUsage(mappingUsage, 'scopedDescriptionRuleHits', id);
    }
  }

  const allowGlobal = mapping?.requireScopedDescriptionRules !== true || scopedMatched;
  if (allowGlobal) {
    for (const [from, to] of Object.entries(mapping?.descriptionReplacements ?? {})) {
      const { next, changed } = replaceTokenStrict(text, from, to);
      text = next;
      if (changed) bumpUsage(mappingUsage, 'descriptionReplacementHits', String(from));
    }
  }
  return text.replace(/\s+/g, ' ').trim();
};

const buildFixedMasterWorkbook = ({
  worksheet,
  rows,
  header,
  mapping,
  mappingUsage,
  maxChildrenLoserFixes,
}) => {
  const rowByFuncLoc = new Map(rows.map((r) => [r.funcLocAfter, r]));
  const descByFuncLoc = new Map(rows.map((r) => [r.funcLocAfter, r.desc]));

  // Pass 1: apply mapping-driven normalization to existing rows.
  // Process shallow FUNCLOC levels first so parent DESC/EQKTU mapping is never read stale for children.
  const pass1Order = [...rows].sort((a, b) => a.level - b.level || a.funcLocAfter.localeCompare(b.funcLocAfter));
  for (const r of pass1Order) {
    const parentDesc = descByFuncLoc.get(r.parentKey) ?? '';
    const childSuffix = getRelativeChildSuffix(r.funcLocAfter, r.parentKey);
    const oldFuncLoc = r.funcLocAfter;
    const mappedFuncLoc = mapping?.funcLocReplacements?.[r.funcLocAfter] ?? r.funcLocAfter;
    const mappedDesc = normalizeDescEqktu(
      applyDescMappingForOutput(
        r.desc,
        parentDesc,
        mapping,
        { level5FuncLoc: r.parentKey, childFuncLoc: r.funcLocAfter, childSuffix },
        mappingUsage,
      ),
    );
    const excelRow = worksheet.getRow(r.rowNumber);
    excelRow.getCell(header.funclocCol).value = mappedFuncLoc;
    excelRow.getCell(header.descCol).value = mappedDesc;
    if (header.eqktuCol != null) {
      excelRow.getCell(header.eqktuCol).value = mappedDesc;
    }
    r.funcLocAfter = mappedFuncLoc;
    r.desc = mappedDesc;
    if (oldFuncLoc !== mappedFuncLoc) {
      descByFuncLoc.delete(oldFuncLoc);
    }
    descByFuncLoc.set(mappedFuncLoc, mappedDesc);
  }

  // Reindex after pass 1.
  rowByFuncLoc.clear();
  rows.forEach((r) => rowByFuncLoc.set(r.funcLocAfter, r));

  // Pass 2: apply resolved max-children adaptation rows.
  for (const fix of maxChildrenLoserFixes.filter((x) => x.uniqueConstraintOk)) {
    if (fix.currentChildFuncLoc && rowByFuncLoc.has(fix.currentChildFuncLoc)) {
      const current = rowByFuncLoc.get(fix.currentChildFuncLoc);
      const row = worksheet.getRow(current.rowNumber);
      row.getCell(header.funclocCol).value = fix.targetChildFuncLoc;
      row.getCell(header.descCol).value = fix.targetChildDesc;
      if (header.eqktuCol != null) {
        row.getCell(header.eqktuCol).value = fix.targetChildDesc;
      }
      if (header.equipmentGroupAfterCol != null) {
        row.getCell(header.equipmentGroupAfterCol).value = fix.targetChildEquipmentGroupAfter ?? '';
      }
      rowByFuncLoc.delete(current.funcLocAfter);
      current.funcLocAfter = fix.targetChildFuncLoc;
      current.desc = fix.targetChildDesc;
      current.equipmentGroupAfter = fix.targetChildEquipmentGroupAfter ?? '';
      rowByFuncLoc.set(current.funcLocAfter, current);
    } else {
      const template =
        rowByFuncLoc.get(`${fix.loserFuncLoc}-${fix.childSuffix}`) ??
        rowByFuncLoc.get(fix.loserFuncLoc) ??
        rows[rows.length - 1];
      const newRow = worksheet.addRow([]);
      if (template) {
        const tRow = worksheet.getRow(template.rowNumber);
        if (typeof tRow.height === 'number' && tRow.height > 0) {
          newRow.height = tRow.height;
        }
        tRow.eachCell({ includeEmpty: true }, (cell, col) => {
          const c = newRow.getCell(col);
          c.style = { ...cell.style };
        });
      }
      newRow.getCell(header.funclocCol).value = fix.targetChildFuncLoc;
      newRow.getCell(header.descCol).value = fix.targetChildDesc;
      if (header.eqktuCol != null) {
        newRow.getCell(header.eqktuCol).value = fix.targetChildDesc;
      }
      if (header.equipmentGroupAfterCol != null) {
        newRow.getCell(header.equipmentGroupAfterCol).value = fix.targetChildEquipmentGroupAfter ?? '';
      }
      const newEntry = {
        rowNumber: newRow.number,
        funcLocAfter: fix.targetChildFuncLoc,
        desc: fix.targetChildDesc,
        equipmentGroupAfter: fix.targetChildEquipmentGroupAfter ?? '',
        level: getLevel(fix.targetChildFuncLoc),
        parentKey: getParentKey(fix.targetChildFuncLoc),
        descBase: canonicalDescBaseForGrouping(fix.targetChildDesc),
        descNo: parseLastTrailingNumber(fix.targetChildDesc),
      };
      rows.push(newEntry);
      rowByFuncLoc.set(newEntry.funcLocAfter, newEntry);
    }
  }
};

/** Keep formula objects for stripFormulas; flatten richText/hyperlink so we never write "[object Object]". */
const snapshotCellValueForReorder = (v) => {
  if (v === null || v === undefined) return v;
  if (typeof v !== 'object' || v instanceof Date) return v;
  if ('formula' in v || 'sharedFormula' in v) return v;
  return excelJsValueToPlainText(v);
};

const reorderWorksheetByFuncLoc = (worksheet, header) => {
  const startRow = Number(header?.headerRowNumber ?? 1) + 1;
  const maxCol = Math.max(
    worksheet.columnCount,
    Number(header?.funclocCol ?? 0),
    Number(header?.eqktuCol ?? 0),
    Number(header?.descCol ?? 0),
    Number(header?.equipmentGroupAfterCol ?? 0),
  );
  if (maxCol <= 0) return;

  const dataSnapshots = [];
  for (let r = startRow; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const funcLoc = excelJsValueToPlainText(row.getCell(header.funclocCol).value)
      .replace(/\s+/g, '')
      .toUpperCase();
    if (!funcLoc) continue;
    const cells = [];
    for (let c = 1; c <= maxCol; c += 1) {
      const cell = row.getCell(c);
      cells.push({
        value: snapshotCellValueForReorder(cell.value),
        style: cell.style ? { ...cell.style } : undefined,
      });
    }
    dataSnapshots.push({ funcLoc, height: row.height, cells });
  }

  dataSnapshots.sort((a, b) => a.funcLoc.localeCompare(b.funcLoc));

  // Clear original data area first.
  for (let r = startRow; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    for (let c = 1; c <= maxCol; c += 1) {
      const cell = row.getCell(c);
      cell.value = null;
    }
  }

  // Write sorted rows back.
  for (let i = 0; i < dataSnapshots.length; i += 1) {
    const targetRow = worksheet.getRow(startRow + i);
    const snap = dataSnapshots[i];
    targetRow.height = snap.height;
    for (let c = 1; c <= maxCol; c += 1) {
      const targetCell = targetRow.getCell(c);
      const src = snap.cells[c - 1];
      targetCell.value = src?.value ?? null;
      targetCell.style = src?.style ? { ...src.style } : {};
    }
  }
};

/**
 * After FUNCLOC sort, physical row indices change — refresh rowNumber on row objects so they match the sheet.
 */
const resyncRowMetadataFromWorksheet = (worksheet, header, rows) => {
  const startRow = Number(header?.headerRowNumber ?? 1) + 1;
  const funclocCol = Number(header?.funclocCol ?? 0);
  if (!funclocCol || !rows?.length) return;
  const byFuncLoc = new Map();
  for (const rec of rows) {
    const k = String(rec.funcLocAfter ?? '')
      .trim()
      .toUpperCase();
    if (k) byFuncLoc.set(k, rec);
  }
  for (let rn = startRow; rn <= worksheet.rowCount; rn += 1) {
    const row = worksheet.getRow(rn);
    const fl = excelJsValueToPlainText(row.getCell(funclocCol).value).replace(/\s+/g, '').toUpperCase();
    if (!fl) continue;
    const rec = byFuncLoc.get(fl);
    if (rec) rec.rowNumber = rn;
  }
};

const applyFinalCleanStyleBolding = (worksheet, header) => {
  const startRow = Number(header?.headerRowNumber ?? 1) + 1;
  const funclocCol = Number(header?.funclocCol ?? 0);
  const descCol = Number(header?.descCol ?? 0);
  const eqktuCol = Number(header?.eqktuCol ?? 0);
  if (!funclocCol || !descCol) return;

  for (let r = startRow; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const funcloc = excelJsValueToPlainText(row.getCell(funclocCol).value).replace(/\s+/g, '').toUpperCase();
    if (!funcloc) continue;
    const segments = funcloc.split('-').filter(Boolean);
    const n = segments.length;
    if (n !== 4 && n !== 5 && n < 6) continue;

    const boldDescEqktu = n === 4 || n === 5;

    const descCell = row.getCell(descCol);
    const descBase = descCell.style || {};
    const descFont = { ...(descBase.font || {}), bold: boldDescEqktu };
    descCell.style = { ...descBase, font: descFont };
    if (eqktuCol) {
      const eqCell = row.getCell(eqktuCol);
      const eqBase = eqCell.style || {};
      const eqFont = { ...(eqBase.font || {}), bold: boldDescEqktu };
      eqCell.style = { ...eqBase, font: eqFont };
    }
  }
};

const analyzeLevel5 = (rows, mapping) => {
  const level5Rows = rows.filter((r) => r.level === 5);
  const level6Rows = rows.filter((r) => r.level === 6);

  const childrenByLevel5 = new Map();
  for (const child of level6Rows) {
    const p = child.parentKey;
    if (!childrenByLevel5.has(p)) childrenByLevel5.set(p, []);
    childrenByLevel5.get(p).push(child);
  }

  const groups = new Map();
  for (const row of level5Rows) {
    const groupKey = canonicalSiblingGroupKey(row.parentKey, row.desc, mapping);
    if (!groups.has(groupKey)) groups.set(groupKey, []);
    groups.get(groupKey).push(row);
  }

  const decisions = [];
  const childrenAudit = [];
  const summary = {
    totalGroups: 0,
    maxChildrenWinner: 0,
    maxChildrenWinnerRaw: 0,
    maxChildrenResolvedByAdoption: 0,
    tieSameChildren: 0,
    tieDifferentChildren: 0,
  };
  const maxChildrenLoserFixes = [];
  const mappingUsage = {
    descriptionReplacementHits: {},
    scopedDescriptionRuleHits: {},
    funcLocReplacementHits: {},
  };
  const currentFuncLocToDesc = new Map();
  const currentDescToFuncLoc = new Map();
  const existingUniquenessViolations = [];
  const plannedUniquenessConflicts = [];

  for (const r of rows) {
    const func = String(r.funcLocAfter || '').trim();
    const desc = String(r.desc || '').trim();
    if (!func || !desc) continue;
    const prevDesc = currentFuncLocToDesc.get(func);
    if (prevDesc && prevDesc !== desc) {
      existingUniquenessViolations.push({
        type: 'funcLoc_to_multiple_desc',
        funcLoc: func,
        descA: prevDesc,
        descB: desc,
      });
    } else if (!prevDesc) {
      currentFuncLocToDesc.set(func, desc);
    }

    const prevFunc = currentDescToFuncLoc.get(desc);
    if (prevFunc && prevFunc !== func) {
      existingUniquenessViolations.push({
        type: 'desc_to_multiple_funcLoc',
        desc,
        funcLocA: prevFunc,
        funcLocB: func,
      });
    } else if (!prevFunc) {
      currentDescToFuncLoc.set(desc, func);
    }
  }

  const plannedFuncLocToDesc = new Map(currentFuncLocToDesc);
  const plannedDescToFuncLoc = new Map(currentDescToFuncLoc);

  groups.forEach((siblings, groupKey) => {
    if (siblings.length <= 1) return;
    summary.totalGroups += 1;

    const withChildren = siblings.map((s) => {
      const children = childrenByLevel5.get(s.funcLocAfter) ?? [];
      const childSignatures = new Set(
        children.map((c) =>
          normalizeChildSignature(
            c.funcLocAfter,
            s.funcLocAfter,
            c.desc,
            s.desc,
            mapping,
            mappingUsage,
          ),
        ),
      );
      return {
        ...s,
        children,
        childCount: children.length,
        childSignatures,
      };
    });

    const maxCount = Math.max(...withChildren.map((s) => s.childCount));
    const topSiblings = withChildren.filter((s) => s.childCount === maxCount);
    const winner =
      topSiblings.length === 1 ? topSiblings[0] : pickWinnerByLowestNumber(topSiblings.map((s) => s));

    let decisionType = 'max_children';
    let reason = 'Selected sibling with highest level-6 children count.';

    if (topSiblings.length > 1) {
      const signatureKeys = topSiblings.map((s) => [...s.childSignatures].sort().join('|'));
      const sameChildrenPattern = signatureKeys.every((v) => v === signatureKeys[0]);
      if (sameChildrenPattern) {
        decisionType = 'tie_same_children';
        reason = 'Tie on child count and same children pattern; branch kept as valid.';
        summary.tieSameChildren += 1;
      } else {
        decisionType = 'tie_different_children';
        reason =
          'Tie on child count but different children pattern; winner chosen by lowest trailing number.';
        summary.tieDifferentChildren += 1;
      }
    } else {
      summary.maxChildrenWinnerRaw += 1;
    }

    const allChildCandidates = new Set();
    topSiblings.forEach((s) => s.childSignatures.forEach((sig) => allChildCandidates.add(sig)));
    const sharedCandidates = new Set([...allChildCandidates]);
    for (const s of topSiblings) {
      for (const candidate of [...sharedCandidates]) {
        if (!s.childSignatures.has(candidate)) {
          sharedCandidates.delete(candidate);
        }
      }
    }

    const winnerSignatures = winner.childSignatures;
    const winnerSuffixes = new Set(
      winner.children.map((c) => getRelativeChildSuffix(c.funcLocAfter, winner.funcLocAfter)),
    );
    const winnerDescNorms = new Set(
      winner.children.map((c) =>
        normalizeChildDescForCompare(
          c.desc,
          winner.desc,
          mapping,
          {
            level5FuncLoc: winner.funcLocAfter,
            childFuncLoc: c.funcLocAfter,
            childSuffix: getRelativeChildSuffix(c.funcLocAfter, winner.funcLocAfter),
          },
          mappingUsage,
        ),
      ),
    );

    const siblingsSummary = withChildren.map((s) => ({
      funcLocAfter: s.funcLocAfter,
      desc: s.desc,
      childCount: s.childCount,
      isWinner: s.funcLocAfter === winner.funcLocAfter,
      childrenDetail: s.children
        .map((c) => {
          const relativeSuffix = getRelativeChildSuffix(c.funcLocAfter, s.funcLocAfter);
          const normalizedDesc = normalizeChildDescForCompare(
            c.desc,
            s.desc,
            mapping,
            {
              level5FuncLoc: s.funcLocAfter,
              childFuncLoc: c.funcLocAfter,
              childSuffix: relativeSuffix,
            },
            mappingUsage,
          );
          const signature = normalizeChildSignature(
            c.funcLocAfter,
            s.funcLocAfter,
            c.desc,
            s.desc,
            mapping,
            mappingUsage,
          );
          const isDifferent = !winnerSignatures.has(signature);
          const diffBySuffix = !winnerSuffixes.has(relativeSuffix);
          const diffByDesc = !winnerDescNorms.has(normalizedDesc);
          return {
            childFuncLoc: c.funcLocAfter,
            relativeSuffix,
            desc: c.desc,
            normalizedDesc,
            signature,
            isDifferent,
            diffBySuffix,
            diffByDesc,
          };
        })
        .sort((a, b) => a.relativeSuffix.localeCompare(b.relativeSuffix)),
    }));

    decisions.push({
      groupKey,
      parentKey: winner.parentKey,
      descBase: applyReplacementsForGroupingKey(winner.desc, mapping),
      siblingCount: siblings.length,
      winnerFuncLoc: winner.funcLocAfter,
      winnerDesc: winner.desc,
      winnerChildCount: winner.childCount,
      topTieCount: topSiblings.length,
      decisionType,
      reason,
      siblingsSummary,
      loserAdaptation: [],
    });

    const latestDecision = decisions[decisions.length - 1];

    if (decisionType === 'tie_different_children' || decisionType === 'tie_same_children') {
      const candidateList = [...allChildCandidates].sort((a, b) => a.localeCompare(b));
      for (const candidate of candidateList) {
        const enabled = sharedCandidates.has(candidate);
        childrenAudit.push({
          groupKey,
          parentKey: winner.parentKey,
          descBase: applyReplacementsForGroupingKey(winner.desc, mapping),
          winnerFuncLoc: winner.funcLocAfter,
          childCandidate: candidate,
          enabled,
          status: enabled ? 'enabled' : 'disabled',
          reason: enabled
            ? 'Candidate exists in all tied siblings.'
            : 'Candidate missing in one or more tied siblings.',
        });
      }
    } else {
      const uniqueChildren = [...winner.childSignatures].sort((a, b) => a.localeCompare(b));
      for (const candidate of uniqueChildren) {
        childrenAudit.push({
          groupKey,
          parentKey: winner.parentKey,
          descBase: applyReplacementsForGroupingKey(winner.desc, mapping),
          winnerFuncLoc: winner.funcLocAfter,
          childCandidate: candidate,
          enabled: true,
          status: 'enabled',
          reason: 'Winner branch child candidate.',
        });
      }

      // For max-children resolution: propose how loser siblings should adapt to winner child structure.
      const winnerChildrenBySuffix = new Map();
      for (const wc of winner.children) {
        const suffix = getRelativeChildSuffix(wc.funcLocAfter, winner.funcLocAfter);
        winnerChildrenBySuffix.set(suffix, wc);
      }
      const losers = withChildren.filter((s) => s.funcLocAfter !== winner.funcLocAfter);
      for (const loser of losers) {
        const loserChildrenBySuffix = new Map();
        for (const lc of loser.children) {
          const suffix = getRelativeChildSuffix(lc.funcLocAfter, loser.funcLocAfter);
          loserChildrenBySuffix.set(suffix, lc);
        }
        for (const [suffix, winnerChild] of winnerChildrenBySuffix.entries()) {
          const currentLoserChild = loserChildrenBySuffix.get(suffix) ?? null;
          const targetFuncLoc = `${loser.funcLocAfter}-${suffix}`;
          const targetDesc = adaptWinnerChildDescToLoserParent(
            winnerChild.desc,
            winner.desc,
            loser.desc,
          );
          const targetDescFinal = applySpecialAdaptationRules(
            targetDesc,
            winner.desc,
            loser.desc,
            mapping,
            { groupKey, loserFuncLoc: loser.funcLocAfter, childSuffix: suffix },
            mappingUsage,
          );
          const action = currentLoserChild == null ? 'add' : 'update';
          const uniquenessReasons = [];
          const prevPlannedDesc = plannedFuncLocToDesc.get(targetFuncLoc);
          if (prevPlannedDesc && prevPlannedDesc !== targetDescFinal) {
            uniquenessReasons.push(
              `FUNCLOC already mapped to a different DESC: ${prevPlannedDesc}`,
            );
          }
          const prevPlannedFunc = plannedDescToFuncLoc.get(targetDescFinal);
          if (prevPlannedFunc && prevPlannedFunc !== targetFuncLoc) {
            uniquenessReasons.push(
              `DESC already mapped to a different FUNCLOC: ${prevPlannedFunc}`,
            );
          }
          const uniqueConstraintOk = uniquenessReasons.length === 0;
          if (uniqueConstraintOk) {
            plannedFuncLocToDesc.set(targetFuncLoc, targetDescFinal);
            plannedDescToFuncLoc.set(targetDescFinal, targetFuncLoc);
          } else {
            plannedUniquenessConflicts.push({
              groupKey,
              loserFuncLoc: loser.funcLocAfter,
              targetChildFuncLoc: targetFuncLoc,
              targetChildDesc: targetDescFinal,
              reasons: uniquenessReasons,
            });
          }
          const targetEqGroup = String(winnerChild.equipmentGroupAfter ?? '').trim();
          const fixRow = {
            groupKey,
            winnerFuncLoc: winner.funcLocAfter,
            winnerDesc: winner.desc,
            loserFuncLoc: loser.funcLocAfter,
            loserDesc: loser.desc,
            childSuffix: suffix,
            action,
            currentChildFuncLoc: currentLoserChild?.funcLocAfter ?? null,
            currentChildDesc: currentLoserChild?.desc ?? null,
            currentChildEquipmentGroupAfter: String(currentLoserChild?.equipmentGroupAfter ?? '').trim(),
            targetChildFuncLoc: targetFuncLoc,
            targetChildDesc: targetDescFinal,
            targetChildEquipmentGroupAfter: targetEqGroup,
            uniqueConstraintOk,
            uniquenessReasons,
          };
          latestDecision.loserAdaptation.push(fixRow);
          maxChildrenLoserFixes.push(fixRow);
        }
      }

      // Max-children branch is considered resolved once adaptation plan exists.
      summary.maxChildrenResolvedByAdoption += 1;
    }
  });

  summary.maxChildrenWinner = Math.max(
    0,
    summary.maxChildrenWinnerRaw - summary.maxChildrenResolvedByAdoption,
  );

  return {
    decisions,
    childrenAudit,
    summary,
    mappingUsage,
    maxChildrenLoserFixes,
    uniquenessReport: {
      existingViolations: existingUniquenessViolations,
      plannedConflicts: plannedUniquenessConflicts,
      existingViolationCount: existingUniquenessViolations.length,
      plannedConflictCount: plannedUniquenessConflicts.length,
    },
  };
};

const analyzeTwoNoPatterns = (rows) => {
  const byFunc = new Map();
  rows.forEach((r) => byFunc.set(String(r.funcLocAfter || '').trim().toUpperCase(), r));

  const summary = {
    totalRows: rows.length,
    twoNoRows: 0,
    first_unit_second_parent: 0,
    first_parent_second_unit: 0,
    both_equal_parent: 0,
    contains_parent_once_other_diff: 0,
    parent_missing: 0,
    no_parent_number: 0,
    ambiguous: 0,
  };

  const detailRows = [];
  for (const r of rows) {
    const nums = extractNoNumbers(r.desc);
    if (nums.length !== 2) continue;
    summary.twoNoRows += 1;

    const parent = byFunc.get(String(r.parentKey || '').trim().toUpperCase());
    if (!parent) {
      summary.parent_missing += 1;
      detailRows.push({
        class: 'parent_missing',
        funcLoc: r.funcLocAfter,
        desc: r.desc,
        parentFuncLoc: r.parentKey,
        parentDesc: '',
        parentNo: null,
        firstNo: nums[0],
        secondNo: nums[1],
      });
      continue;
    }

    const parentNums = extractNoNumbers(parent.desc);
    const parentNo = parentNums.length > 0 ? parentNums[parentNums.length - 1] : null;
    if (parentNo == null) {
      summary.no_parent_number += 1;
      detailRows.push({
        class: 'no_parent_number',
        funcLoc: r.funcLocAfter,
        desc: r.desc,
        parentFuncLoc: parent.funcLocAfter,
        parentDesc: parent.desc,
        parentNo: null,
        firstNo: nums[0],
        secondNo: nums[1],
      });
      continue;
    }

    const [a, b] = nums;
    let cls = 'ambiguous';
    if (a !== parentNo && b === parentNo) cls = 'first_unit_second_parent';
    else if (a === parentNo && b !== parentNo) cls = 'first_parent_second_unit';
    else if (a === parentNo && b === parentNo) cls = 'both_equal_parent';
    else if (a !== parentNo && b !== parentNo) cls = 'contains_parent_once_other_diff';
    summary[cls] += 1;

    detailRows.push({
      class: cls,
      funcLoc: r.funcLocAfter,
      desc: r.desc,
      parentFuncLoc: parent.funcLocAfter,
      parentDesc: parent.desc,
      parentNo,
      firstNo: a,
      secondNo: b,
    });
  }

  return { summary, detailRows };
};

const applyHeaderStyle = (row) => {
  row.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_HEADER } };
    cell.font = { bold: true, color: { argb: COLOR_HEADER_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  });
};

const writeExcelReport = async ({ decisions, childrenAudit, summary, outputPath, sourceUrl }) => {
  const workbook = new ExcelJS.Workbook();

  const summarySheet = workbook.addWorksheet('Summary');
  summarySheet.addRow(['Metric', 'Value']);
  applyHeaderStyle(summarySheet.getRow(1));
  summarySheet.addRow(['Source URL', sourceUrl]);
  summarySheet.addRow(['Total sibling groups checked', summary.totalGroups]);
  summarySheet.addRow(['Resolved by max children', summary.maxChildrenWinner]);
  summarySheet.addRow(['Tie with same children', summary.tieSameChildren]);
  summarySheet.addRow(['Tie with different children', summary.tieDifferentChildren]);
  summarySheet.columns = [{ width: 36 }, { width: 80 }];

  const decisionsSheet = workbook.addWorksheet('Decisions');
  decisionsSheet.addRow([
    'Parent FUNCLOC',
    'Desc Base',
    'Group Key',
    'Sibling Count',
    'Top Tie Count',
    'Winner Level-5 FUNCLOC',
    'Winner Description',
    'Winner Child Count',
    'Decision Type',
    'Reason',
  ]);
  applyHeaderStyle(decisionsSheet.getRow(1));

  for (const row of decisions) {
    const excelRow = decisionsSheet.addRow([
      row.parentKey,
      row.descBase,
      row.groupKey,
      row.siblingCount,
      row.topTieCount,
      row.winnerFuncLoc,
      row.winnerDesc,
      row.winnerChildCount,
      row.decisionType,
      row.reason,
    ]);
    if (row.decisionType === 'max_children') {
      excelRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_WINNER } };
      });
    } else if (row.decisionType === 'tie_same_children') {
      excelRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_TIE_SAME } };
      });
    } else {
      excelRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_TIE_DIFF } };
      });
    }
  }

  decisionsSheet.columns = [
    { width: 22 },
    { width: 32 },
    { width: 52 },
    { width: 14 },
    { width: 14 },
    { width: 26 },
    { width: 40 },
    { width: 18 },
    { width: 24 },
    { width: 56 },
  ];

  const auditSheet = workbook.addWorksheet('ChildrenAudit');
  auditSheet.addRow([
    'Parent FUNCLOC',
    'Desc Base',
    'Group Key',
    'Winner Level-5 FUNCLOC',
    'Child Candidate (normalized)',
    'Status',
    'Enabled',
    'Reason',
  ]);
  applyHeaderStyle(auditSheet.getRow(1));

  for (const row of childrenAudit) {
    const excelRow = auditSheet.addRow([
      row.parentKey,
      row.descBase,
      row.groupKey,
      row.winnerFuncLoc,
      row.childCandidate,
      row.status,
      row.enabled ? 'YES' : 'NO',
      row.reason,
    ]);
    excelRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: row.enabled ? COLOR_ENABLED : COLOR_DISABLED },
      };
      if (!row.enabled) {
        cell.font = { color: { argb: COLOR_DISABLED_TEXT } };
      }
    });
  }

  auditSheet.columns = [
    { width: 22 },
    { width: 32 },
    { width: 52 },
    { width: 26 },
    { width: 36 },
    { width: 12 },
    { width: 10 },
    { width: 52 },
  ];

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await workbook.xlsx.writeFile(outputPath);
};

const escapeHtml = (str) =>
  String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const writeHtmlReport = ({ decisions, childrenAudit, summary, outputPath, sourceUrl }) => {
  const byParent = new Map();
  const childrenByGroup = new Map();

  for (const d of decisions) {
    // For the grouped drill-down view, only keep:
    // - max_children
    // - tie_different_children
    if (d.decisionType === 'max_children' || d.decisionType === 'tie_different_children') {
      if (!byParent.has(d.parentKey)) byParent.set(d.parentKey, []);
      byParent.get(d.parentKey).push(d);
    }
  }

  for (const row of childrenAudit) {
    const arr = childrenByGroup.get(row.groupKey) ?? [];
    arr.push(row);
    childrenByGroup.set(row.groupKey, arr);
  }

  const parentSections = [...byParent.entries()]
    .sort(([a], [b]) => String(a).localeCompare(String(b)))
    .map(([parentKey, groups], parentIndex) => {
      const groupCards = groups
        .map((g, groupIndex) => {
          const groupId = `g_${parentIndex}_${groupIndex}`;
          const groupChildren = (childrenByGroup.get(g.groupKey) ?? []).sort((a, b) =>
            a.childCandidate.localeCompare(b.childCandidate),
          );
          const decisionBadge =
            g.decisionType === 'max_children'
              ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-200'
              : g.decisionType === 'tie_same_children'
                ? 'bg-sky-100 text-sky-800 dark:bg-sky-900/40 dark:text-sky-200'
                : 'bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-200';

          const decisionLabel =
            g.decisionType === 'max_children'
              ? 'Max children'
              : g.decisionType === 'tie_same_children'
                ? 'Tie – same children'
                : 'Tie – different children';

          const siblingsRows = (g.siblingsSummary ?? [])
            .map(
              (s) => `<tr>
  <td class="font-mono text-[11px]">${escapeHtml(s.funcLocAfter)}</td>
  <td class="text-xs">${escapeHtml(s.desc)}</td>
  <td class="text-xs text-right">${s.childCount}</td>
  <td class="text-[11px] text-right">${s.isWinner ? '<span class="inline-flex items-center rounded-full bg-emerald-500/10 text-emerald-700 dark:text-emerald-200 px-2 py-0.5">Winner</span>' : ''}</td>
</tr>`,
            )
            .join('\n');

          const siblingsChildrenDetailBlocks = (g.siblingsSummary ?? [])
            .map((s) => {
              const rows = (s.childrenDetail ?? [])
                .map(
                  (d) => `<tr class="${d.isDifferent ? 'bg-amber-50/70 dark:bg-amber-950/20' : ''}">
  <td class="font-mono text-[11px] ${d.diffBySuffix ? 'text-amber-700 dark:text-amber-200 font-semibold' : ''}">${escapeHtml(d.relativeSuffix || '(direct)')}</td>
  <td class="font-mono text-[11px] ${d.diffBySuffix ? 'text-amber-700 dark:text-amber-200' : ''}">${escapeHtml(d.childFuncLoc)}</td>
  <td class="text-xs ${d.diffByDesc ? 'text-rose-700 dark:text-rose-200 font-semibold' : ''}">${escapeHtml(d.desc || '-')}</td>
  <td class="text-[11px] ${d.diffByDesc ? 'text-rose-700 dark:text-rose-200' : 'text-zinc-500 dark:text-zinc-400'}">${escapeHtml(d.normalizedDesc || '-')}</td>
  <td class="text-[11px] flex flex-wrap gap-1 items-center">
    ${d.isDifferent ? '<span class="inline-flex items-center rounded-full bg-amber-500/10 text-amber-700 dark:text-amber-200 px-2 py-0.5">Different</span>' : ''}
    ${d.diffBySuffix ? '<span class="inline-flex items-center rounded-full bg-amber-500/10 text-amber-700 dark:text-amber-200 px-2 py-0.5">FUNCLOC diff</span>' : ''}
    ${d.diffByDesc ? '<span class="inline-flex items-center rounded-full bg-rose-500/10 text-rose-700 dark:text-rose-200 px-2 py-0.5">DESC diff</span>' : ''}
  </td>
</tr>`,
                )
                .join('\n');

              return `<div class="rounded-md border ${s.isWinner ? 'border-emerald-300 dark:border-emerald-700 bg-emerald-50/50 dark:bg-emerald-950/20' : 'border-zinc-100 dark:border-zinc-800 bg-zinc-50/70 dark:bg-zinc-900/50'} p-2">
  <h5 class="mb-1 text-[11px] font-semibold text-zinc-700 dark:text-zinc-200">
    Sibling ${escapeHtml(s.desc || '-')}
    <span class="font-mono text-[10px] text-zinc-500 dark:text-zinc-400">(${escapeHtml(s.funcLocAfter)})</span>
    ${s.isWinner ? '<span class="ml-1 inline-flex items-center rounded-full bg-emerald-500/10 text-emerald-700 dark:text-emerald-200 px-2 py-0.5 text-[10px]">Winner</span>' : ''}
  </h5>
  <div class="overflow-x-auto rounded-md border border-zinc-100 dark:border-zinc-800 bg-white/80 dark:bg-zinc-900/60">
    <table class="min-w-full border-collapse text-[11px]">
      <thead class="bg-zinc-100/80 dark:bg-zinc-900/80 text-zinc-600 dark:text-zinc-300">
        <tr>
          <th class="px-2 py-1 text-left font-medium">Relative suffix</th>
          <th class="px-2 py-1 text-left font-medium">Child FUNCLOC</th>
          <th class="px-2 py-1 text-left font-medium">Desc</th>
          <th class="px-2 py-1 text-left font-medium">Normalized desc</th>
          <th class="px-2 py-1 text-left font-medium">Match</th>
        </tr>
      </thead>
      <tbody class="divide-y divide-zinc-100 dark:divide-zinc-800">
        ${rows || '<tr><td colspan="5" class="px-2 py-1 text-[11px] text-zinc-500 dark:text-zinc-400 italic">No children</td></tr>'}
      </tbody>
    </table>
  </div>
</div>`;
            })
            .join('\n');

          const loserAdaptationRows = (g.loserAdaptation ?? [])
            .map(
              (r) => `<tr>
  <td class="font-mono text-[11px]">${escapeHtml(r.loserFuncLoc)}</td>
  <td class="font-mono text-[11px]">${escapeHtml(r.childSuffix)}</td>
  <td class="text-[11px]">${escapeHtml(r.action)}</td>
  <td class="text-[11px]">
    ${
      r.uniqueConstraintOk
        ? '<span class="inline-flex items-center rounded-full bg-emerald-500/10 text-emerald-700 dark:text-emerald-200 px-2 py-0.5">OK</span>'
        : '<span class="inline-flex items-center rounded-full bg-rose-500/10 text-rose-700 dark:text-rose-200 px-2 py-0.5">CONFLICT</span>'
    }
  </td>
  <td class="font-mono text-[11px]">${escapeHtml(r.currentChildFuncLoc ?? '-')}</td>
  <td class="text-xs">${escapeHtml(r.currentChildDesc ?? '-')}</td>
  <td class="font-mono text-[11px]">${escapeHtml(r.targetChildFuncLoc)}</td>
  <td class="text-xs">${escapeHtml(r.targetChildDesc)}</td>
  <td class="text-xs">${escapeHtml(r.targetChildEquipmentGroupAfter ?? '-')}</td>
  <td class="text-xs text-zinc-500 dark:text-zinc-400">${escapeHtml((r.uniquenessReasons ?? []).join(' | ') || '-')}</td>
</tr>`,
            )
            .join('\n');

          const childrenRows =
            groupChildren.length === 0
              ? '<tr><td colspan="4" class="text-[11px] text-zinc-500 dark:text-zinc-400 italic">No child candidates recorded for this group.</td></tr>'
              : groupChildren
                  .map(
                    (c) => `<tr>
  <td class="font-mono text-[11px]">${escapeHtml(c.winnerFuncLoc)}</td>
  <td class="text-xs">${escapeHtml(c.childCandidate)}</td>
  <td class="text-[11px]">${c.enabled ? '<span class="inline-flex items-center rounded-full bg-emerald-500/10 text-emerald-700 dark:text-emerald-200 px-2 py-0.5">Enabled</span>' : '<span class="inline-flex items-center rounded-full bg-zinc-500/10 text-zinc-500 dark:bg-zinc-800 dark:text-zinc-300 px-2 py-0.5">Disabled</span>'}</td>
  <td class="text-[11px] text-zinc-500 dark:text-zinc-400">${escapeHtml(c.reason)}</td>
</tr>`,
                  )
                  .join('\n');

          return `<section class="not-prose rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white/70 dark:bg-zinc-900/60">
  <button
    type="button"
    class="w-full flex items-center justify-between px-3 py-2 text-left hover:bg-zinc-50 dark:hover:bg-zinc-800/70 transition"
    data-group-toggle="${groupId}"
  >
    <div class="flex-1 min-w-0">
      <p class="text-xs font-medium text-zinc-800 dark:text-zinc-100 truncate">
        ${escapeHtml(g.descBase || '(no desc base)')}
      </p>
      <p class="text-[11px] text-zinc-500 dark:text-zinc-400 font-mono truncate">
        ${escapeHtml(g.groupKey)}
      </p>
    </div>
    <div class="flex items-center gap-2 pl-3">
      <span class="inline-flex items-center rounded-full px-2 py-0.5 text-[10px] font-medium ${decisionBadge}">
        ${escapeHtml(decisionLabel)}
      </span>
      <span class="text-[10px] text-zinc-400 dark:text-zinc-500">Sib: ${g.siblingCount}</span>
      <span class="text-[10px] text-zinc-400 dark:text-zinc-500">Tie: ${g.topTieCount}</span>
      <svg data-chevron class="h-3 w-3 text-zinc-400 transition-transform" aria-hidden="true" viewBox="0 0 20 20" fill="currentColor">
        <path fill-rule="evenodd" d="M5.23 7.21a.75.75 0 011.06.02L10 10.94l3.71-3.71a.75.75 0 111.06 1.06l-4.24 4.25a.75.75 0 01-1.06 0L5.21 8.27a.75.75 0 01.02-1.06z" clip-rule="evenodd" />
      </svg>
    </div>
  </button>
  <div id="${groupId}" class="border-t border-zinc-200 dark:border-zinc-800 px-3 py-3 space-y-3 hidden">
    <div>
      <h4 class="text-[11px] font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400 mb-1">Siblings at level 5</h4>
      <div class="overflow-x-auto rounded-md border border-zinc-100 dark:border-zinc-800 bg-zinc-50/80 dark:bg-zinc-900/60">
        <table class="min-w-full border-collapse text-[11px]">
          <thead class="bg-zinc-100/80 dark:bg-zinc-900/80 text-zinc-600 dark:text-zinc-300">
            <tr>
              <th class="px-2 py-1 text-left font-medium">FUNCLOC</th>
              <th class="px-2 py-1 text-left font-medium">Description</th>
              <th class="px-2 py-1 text-right font-medium">Children</th>
              <th class="px-2 py-1 text-right font-medium">Winner</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-zinc-100 dark:divide-zinc-800">
            ${siblingsRows}
          </tbody>
        </table>
      </div>
    </div>
    <div>
      <h4 class="text-[11px] font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400 mb-1">Child candidates</h4>
      <div class="overflow-x-auto rounded-md border border-zinc-100 dark:border-zinc-800 bg-zinc-50/80 dark:bg-zinc-900/60">
        <table class="min-w-full border-collapse text-[11px]">
          <thead class="bg-zinc-100/80 dark:bg-zinc-900/80 text-zinc-600 dark:text-zinc-300">
            <tr>
              <th class="px-2 py-1 text-left font-medium">Winner FUNCLOC</th>
              <th class="px-2 py-1 text-left font-medium">Child (normalized)</th>
              <th class="px-2 py-1 text-left font-medium">Status</th>
              <th class="px-2 py-1 text-left font-medium">Reason</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-zinc-100 dark:divide-zinc-800">
            ${childrenRows}
          </tbody>
        </table>
      </div>
    </div>
    <div>
      <h4 class="text-[11px] font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400 mb-1">Each sibling child details</h4>
      <div class="space-y-2">
        ${siblingsChildrenDetailBlocks}
      </div>
    </div>
    ${
      g.decisionType === 'max_children' && (g.loserAdaptation?.length ?? 0) > 0
        ? `<div>
      <h4 class="text-[11px] font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400 mb-1">Loser adaptation plan (follow winner children)</h4>
      <div class="overflow-x-auto rounded-md border border-zinc-100 dark:border-zinc-800 bg-zinc-50/80 dark:bg-zinc-900/60">
        <table class="min-w-full border-collapse text-[11px]">
          <thead class="bg-zinc-100/80 dark:bg-zinc-900/80 text-zinc-600 dark:text-zinc-300">
            <tr>
              <th class="px-2 py-1 text-left font-medium">Loser FUNCLOC</th>
              <th class="px-2 py-1 text-left font-medium">Suffix</th>
              <th class="px-2 py-1 text-left font-medium">Action</th>
              <th class="px-2 py-1 text-left font-medium">Unique check</th>
              <th class="px-2 py-1 text-left font-medium">Current child FUNCLOC</th>
              <th class="px-2 py-1 text-left font-medium">Current child desc</th>
              <th class="px-2 py-1 text-left font-medium">Target child FUNCLOC</th>
              <th class="px-2 py-1 text-left font-medium">Target child desc</th>
              <th class="px-2 py-1 text-left font-medium">Target EQUIPMENT GROUP AFTER (winner)</th>
              <th class="px-2 py-1 text-left font-medium">Conflict reason</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-zinc-100 dark:divide-zinc-800">
            ${loserAdaptationRows}
          </tbody>
        </table>
      </div>
    </div>`
        : ''
    }
  </div>
</section>`;
        })
        .join('\n');

      return `<section class="mt-4">
  <h3 class="text-sm font-semibold text-zinc-800 dark:text-zinc-100">
    Level-4 parent: <span class="font-mono text-[11px] align-middle">${escapeHtml(parentKey)}</span>
  </h3>
  <div class="mt-2 space-y-2">
    ${groupCards}
  </div>
</section>`;
    })
    .join('\n');

  const html = `<!DOCTYPE html>
<html lang="en" class="h-full">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Bah Jambi Level-5 Checker Report</title>
  <script src="https://cdn.tailwindcss.com?plugins=typography"></script>
  <script>
    (function () {
      const root = document.documentElement;
      try {
        const stored = window.localStorage.getItem('bj-theme');
        const prefersDark =
          window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
        if (stored === 'dark' || (!stored && prefersDark)) root.classList.add('dark');
        else root.classList.remove('dark');
      } catch {}
    })();
  </script>
  <style>
    :root { color-scheme: light dark; }
    body { min-height: 100vh; }
    [data-group-content].is-open { display: block; }
    [data-group-toggle] [data-chevron] { transition: transform .2s ease; }
    [data-group-toggle].is-open [data-chevron] { transform: rotate(180deg); }
  </style>
</head>
<body class="bg-zinc-50 text-zinc-900 dark:bg-zinc-950 dark:text-zinc-100">
  <div class="min-h-screen flex flex-col">
    <header class="border-b border-zinc-200 dark:border-zinc-800 bg-white/80 dark:bg-zinc-900/80 backdrop-blur">
      <div class="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-4">
        <div>
          <h1 class="text-lg font-semibold tracking-tight">Bah Jambi Level-5 Checker</h1>
          <p class="text-xs text-zinc-500 dark:text-zinc-400">Sibling pruning recap by children patterns</p>
        </div>
        <button
          type="button"
          id="theme-toggle"
          class="rounded-full border border-zinc-200 dark:border-zinc-700 px-3 py-1 text-xs font-medium text-zinc-700 dark:text-zinc-200 bg-white/80 dark:bg-zinc-900/80 shadow-sm hover:bg-zinc-50 dark:hover:bg-zinc-800 transition"
        >
          Toggle theme
        </button>
      </div>
    </header>

    <main class="flex-1">
      <div class="max-w-6xl mx-auto px-4 py-6">
        <article class="prose prose-sm prose-zinc dark:prose-invert max-w-none">
          <section>
            <h2>Summary</h2>
            <p class="text-xs text-zinc-500 dark:text-zinc-400">
              Source:
              <a href="${escapeHtml(
                sourceUrl,
              )}" class="break-all text-zinc-900 dark:text-zinc-50 underline underline-offset-2 decoration-zinc-400">
                ${escapeHtml(sourceUrl)}
              </a>
            </p>
            <div class="not-prose mt-3 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
              <div class="rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white/70 dark:bg-zinc-900/70 px-3 py-2">
                <p class="text-xs text-zinc-500 dark:text-zinc-400">Total sibling groups</p>
                <p class="mt-1 text-lg font-semibold">${summary.totalGroups}</p>
              </div>
              <div class="rounded-lg border border-emerald-200 dark:border-emerald-900/70 bg-emerald-50/70 dark:bg-emerald-950/40 px-3 py-2">
                <p class="text-xs text-emerald-700 dark:text-emerald-300/90">Resolved by max children</p>
                <p class="mt-1 text-lg font-semibold">${summary.maxChildrenWinner}</p>
              </div>
              <div class="rounded-lg border border-sky-200 dark:border-sky-900/70 bg-sky-50/70 dark:bg-sky-950/40 px-3 py-2">
                <p class="text-xs text-sky-700 dark:text-sky-300/90">Tie – same children</p>
                <p class="mt-1 text-lg font-semibold">${summary.tieSameChildren}</p>
              </div>
              <div class="rounded-lg border border-amber-200 dark:border-amber-900/70 bg-amber-50/70 dark:bg-amber-950/40 px-3 py-2">
                <p class="text-xs text-amber-700 dark:text-amber-300/90">Tie – different children</p>
                <p class="mt-1 text-lg font-semibold">${summary.tieDifferentChildren}</p>
              </div>
            </div>
          </section>

          <section>
            <h2>Groups by level 4 parent</h2>
            <p class="text-sm text-zinc-500 dark:text-zinc-400">
              Click a parent group to reveal its level-5 siblings and their child candidates.
            </p>
            <div class="mt-3 space-y-4">
              ${parentSections}
            </div>
          </section>

          <section>
            <h2>Raw counts</h2>
            <pre class="text-[11px] leading-relaxed bg-zinc-900/95 text-zinc-100 rounded-lg px-3 py-2 overflow-x-auto border border-zinc-800"><code>${escapeHtml(
              JSON.stringify(summary, null, 2),
            )}</code></pre>
          </section>
        </article>
      </div>
    </main>

    <footer class="border-t border-zinc-200 dark:border-zinc-800 text-[11px] text-zinc-500 dark:text-zinc-400">
      <div class="max-w-6xl mx-auto px-4 py-2 flex flex-wrap items-center justify-between gap-2">
        <span>Bah Jambi Level-5 Checker</span>
        <span class="text-[10px]">Generated at ${escapeHtml(getJakartaStamp())}</span>
      </div>
    </footer>
  </div>

  <script>
    (function () {
      const root = document.documentElement;
      const themeBtn = document.getElementById('theme-toggle');
      if (themeBtn) {
        themeBtn.addEventListener('click', () => {
          const isDark = root.classList.toggle('dark');
          try {
            window.localStorage.setItem('bj-theme', isDark ? 'dark' : 'light');
          } catch {}
        });
      }

      const toggles = document.querySelectorAll('[data-group-toggle]');
      toggles.forEach((btn) => {
        btn.addEventListener('click', () => {
          const id = btn.getAttribute('data-group-toggle');
          if (!id) return;
          const panel = document.getElementById(id);
          if (!panel) return;
          const open = panel.classList.toggle('is-open');
          panel.classList.toggle('hidden', !open);
          btn.classList.toggle('is-open', open);
          btn.setAttribute('aria-expanded', String(open));
        });
      });
    })();
  </script>
</body>
</html>`;

  const outDir = path.dirname(outputPath);
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(outputPath, html, 'utf8');
};

const writeJsonReport = ({
  decisions,
  childrenAudit,
  summary,
  outputPath,
  sourceUrl,
  mappingPath,
  mapping,
  mappingUsage,
  maxChildrenLoserFixes,
  uniquenessReport,
  middleGapFill,
  suffixShiftCompact,
  funclocGapScan,
}) => {
  const payload = {
    meta: {
      sourceUrl,
      generatedAtJakarta: getJakartaStamp(),
      version: 1,
      mappingPath,
      grouping: {
        version: 2,
        notes:
          'Level-5 sibling groups: parent FUNCLOC + canonical desc (NFKC whitespace, NO/NR tokens, strip trailing instance digit) + optional fix-map descriptionReplacements + LORY vs STD lane (word "lory" splits competition).',
      },
    },
    appliedMapping: mapping,
    mappingUsage,
    uniquenessReport,
    middleGapFill,
    suffixShiftCompact,
    funclocGapScan,
    summary,
    decisions,
    childrenAudit,
    maxChildrenLoserFixes,
  };
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, JSON.stringify(payload, null, 2), 'utf8');
};

const writeTwoNoHtmlReport = ({ outputPath, sourceUrl, analysis }) => {
  const clsOrder = [
    'first_unit_second_parent',
    'first_parent_second_unit',
    'both_equal_parent',
    'contains_parent_once_other_diff',
    'parent_missing',
    'no_parent_number',
    'ambiguous',
  ];
  const labels = {
    first_unit_second_parent: 'First NO = unit, second NO = parent',
    first_parent_second_unit: 'First NO = parent, second NO = unit',
    both_equal_parent: 'Both NO equal parent NO',
    contains_parent_once_other_diff: 'Two NO values, neither matches parent NO',
    parent_missing: 'Parent row missing',
    no_parent_number: 'Parent has no NO number',
    ambiguous: 'Ambiguous',
  };
  const byClass = new Map(clsOrder.map((k) => [k, []]));
  for (const row of analysis.detailRows) byClass.get(row.class)?.push(row);

  const esc = (v) =>
    String(v ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

  const sections = clsOrder
    .map((k) => {
      const rows = byClass.get(k) ?? [];
      const tr = rows
        .map(
          (r) => `<tr>
<td><code>${esc(r.funcLoc)}</code></td>
<td>${esc(r.desc)}</td>
<td><code>${esc(r.parentFuncLoc)}</code></td>
<td>${esc(r.parentDesc)}</td>
<td>${esc(r.parentNo)}</td>
<td>${esc(r.firstNo)}</td>
<td>${esc(r.secondNo)}</td>
</tr>`,
        )
        .join('\n');
      return `<details ${rows.length > 0 ? 'open' : ''}>
<summary><strong>${esc(labels[k])}</strong> — ${rows.length} row(s)</summary>
<div class="table-wrap">
<table>
<thead><tr><th>FUNCLOC</th><th>Description</th><th>Parent FUNCLOC</th><th>Parent Description</th><th>Parent NO</th><th>First NO</th><th>Second NO</th></tr></thead>
<tbody>${tr || '<tr><td colspan="7"><em>No rows</em></td></tr>'}</tbody>
</table>
</div>
</details>`;
    })
    .join('\n');

  const s = analysis.summary;
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Bah Jambi Two-NO Report</title>
  <style>
    body{font-family:Inter,Segoe UI,Arial,sans-serif;margin:20px;background:#fafafa;color:#111}
    h1,h2{margin:0 0 8px}
    .meta{color:#555;margin-bottom:16px}
    .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin:12px 0 18px}
    .card{background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:10px}
    .k{font-size:12px;color:#555}.v{font-size:18px;font-weight:700}
    details{background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:8px 10px;margin:10px 0}
    summary{cursor:pointer}
    .table-wrap{overflow:auto;margin-top:8px}
    table{border-collapse:collapse;width:100%;min-width:1100px}
    th,td{border:1px solid #e5e7eb;padding:6px 8px;font-size:12px;vertical-align:top}
    th{background:#f3f4f6;text-align:left}
    code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:11px}
  </style>
</head>
<body>
  <h1>Bah Jambi Two-NO Description Report</h1>
  <div class="meta">Source: ${esc(sourceUrl)}<br/>Generated: ${esc(new Date().toISOString())}</div>
  <div class="cards">
    <div class="card"><div class="k">Total rows</div><div class="v">${s.totalRows}</div></div>
    <div class="card"><div class="k">Rows with exactly 2 NO numbers</div><div class="v">${s.twoNoRows}</div></div>
    <div class="card"><div class="k">First=unit, second=parent</div><div class="v">${s.first_unit_second_parent}</div></div>
    <div class="card"><div class="k">First=parent, second=unit</div><div class="v">${s.first_parent_second_unit}</div></div>
    <div class="card"><div class="k">Both equal parent</div><div class="v">${s.both_equal_parent}</div></div>
    <div class="card"><div class="k">Parent missing/no parent NO/ambiguous</div><div class="v">${s.parent_missing + s.no_parent_number + s.ambiguous}</div></div>
  </div>
  ${sections}
</body>
</html>`;
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, html, 'utf8');
};

/** Same as webapp `isLoryDescException` for FUNCLOC gap scan (numeric gaps ignore LORY-reserved slots). */
const isLoryDescExceptionForGapScan = (text) => /\blory/i.test(String(text ?? '').trim());

const gapScanRowIsLoryTagged = (r) =>
  isLoryDescExceptionForGapScan(r.desc) || isLoryDescExceptionForGapScan(r.eqktu);

/**
 * Read sheet rows for gap analysis (separate workbook read so we scan the **written** fixed master).
 */
const loadGapScanRowsFromFixedXlsx = async (xlsxPath) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(xlsxPath);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('No worksheet found for FUNCLOC gap scan.');

  stripFormulasFromWorksheet(worksheet);

  const found = findHeadersExcelJs(
    worksheet,
    [FUNCLOC_HEADER, ...FUNCLOC_DESC_HEADERS, ...EQKTU_HEADERS],
    {
      maxScanRows: worksheet.rowCount,
      minSimilarity: 0.6,
    },
  );
  if (!found) throw new Error('Could not find FUNCLOC/DESC headers for gap scan.');

  const funclocCol = found.colMap.get(normalizeHeader(FUNCLOC_HEADER));
  const descCol =
    found.colMap.get(normalizeHeader(FUNCLOC_DESC_HEADERS[0])) ??
    found.colMap.get(normalizeHeader(FUNCLOC_DESC_HEADERS[1]));
  const eqktuCol =
    found.colMap.get(normalizeHeader(EQKTU_HEADERS[0])) ??
    found.colMap.get(normalizeHeader(EQKTU_HEADERS[1])) ??
    null;
  if (!funclocCol || !descCol) throw new Error('Gap scan: required columns missing.');

  const rows = [];
  for (let r = found.headerRowNumber + 1; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const funcLoc = toCleanString(excelJsValueToPlainText(row.getCell(funclocCol).value)).toUpperCase();
    if (!funcLoc) continue;
    const desc = normalizeDescEqktu(excelJsValueToPlainText(row.getCell(descCol).value));
    const eqktuRaw =
      eqktuCol != null ? excelJsValueToPlainText(row.getCell(eqktuCol).value) : '';
    const eqktu = normalizeDescEqktu(eqktuRaw);
    const segs = funcLoc.split('-').filter(Boolean);
    rows.push({ row: r, funcLoc, desc, eqktu, segs });
  }
  return rows;
};

/**
 * Numeric sibling gaps + missing immediate parents (LORY reserves suffixes like webapp compact gaps).
 */
const analyzeFunclocGaps = (rows) => {
  const byFunc = new Map(rows.map((r) => [r.funcLoc, r]));

  const siblingGroups = new Map();
  for (const r of rows) {
    if (r.segs.length < 2) continue;
    const parent = r.segs.slice(0, -1).join('-');
    const last = r.segs[r.segs.length - 1];
    if (!/^\d{4}$/.test(last)) continue;
    const n = Number.parseInt(last, 10);
    if (!siblingGroups.has(parent)) siblingGroups.set(parent, []);
    siblingGroups.get(parent).push({
      funcLoc: r.funcLoc,
      desc: r.desc,
      eqktu: r.eqktu,
      n,
    });
  }

  const numericGapGroups = [];
  for (const [parent, arr] of siblingGroups.entries()) {
    const loryNs = new Set();
    for (const x of arr) {
      if (gapScanRowIsLoryTagged(x)) loryNs.add(x.n);
    }
    const stdArr = arr.filter((x) => !gapScanRowIsLoryTagged(x));
    if (stdArr.length < 2) continue;
    stdArr.sort((a, b) => a.n - b.n);
    const nums = [...new Set(stdArr.map((x) => x.n))];
    const min = nums[0];
    const max = nums[nums.length - 1];
    const stdSet = new Set(nums);
    const missing = [];
    for (let k = min; k <= max; k += 1) {
      if (stdSet.has(k)) continue;
      if (loryNs.has(k)) continue;
      missing.push(String(k).padStart(4, '0'));
    }
    if (missing.length === 0) continue;

    numericGapGroups.push({
      parentFuncLoc: parent,
      parentLevel: parent.split('-').filter(Boolean).length,
      presentCount: nums.length,
      min: String(min).padStart(4, '0'),
      max: String(max).padStart(4, '0'),
      missingCount: missing.length,
      missing,
      firstPresentFuncLoc: stdArr[0]?.funcLoc ?? '',
      lastPresentFuncLoc: stdArr[stdArr.length - 1]?.funcLoc ?? '',
    });
  }
  numericGapGroups.sort((a, b) => a.parentFuncLoc.localeCompare(b.parentFuncLoc));

  const missingParentMap = new Map();
  for (const r of rows) {
    if (r.segs.length < 2) continue;
    const parent = r.segs.slice(0, -1).join('-');
    if (!parent || byFunc.has(parent)) continue;

    if (!missingParentMap.has(parent)) {
      missingParentMap.set(parent, {
        missingParentFuncLoc: parent,
        expectedLevel: parent.split('-').filter(Boolean).length,
        childCount: 0,
        childLevels: new Set(),
        children: [],
      });
    }
    const rec = missingParentMap.get(parent);
    rec.childCount += 1;
    rec.childLevels.add(r.segs.length);
    rec.children.push({
      row: r.row,
      funcLoc: r.funcLoc,
      desc: r.desc,
    });
  }

  const missingImmediateParents = [...missingParentMap.values()]
    .map((x) => ({
      ...x,
      childLevels: [...x.childLevels].sort((a, b) => a - b),
      children: x.children.sort((a, b) => a.funcLoc.localeCompare(b.funcLoc)),
    }))
    .sort((a, b) => a.missingParentFuncLoc.localeCompare(b.missingParentFuncLoc));

  return {
    summary: {
      totalRows: rows.length,
      numericGapGroups: numericGapGroups.length,
      missingImmediateParents: missingImmediateParents.length,
    },
    numericGapGroups,
    missingImmediateParents,
  };
};

const escapeHtmlFunclocGap = (v) =>
  String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const writeFunclocGapReportHtml = ({ outputPath, sourceUrl, fixedWorkbookPath, payload }) => {
  const gapsHtml = payload.numericGapGroups
    .map(
      (g) => `<details>
<summary><strong>${escapeHtmlFunclocGap(g.parentFuncLoc)}</strong> - missing ${g.missingCount} item(s)</summary>
<div class="box">
  <p>Range (non-LORY): <code>${escapeHtmlFunclocGap(g.min)}</code> .. <code>${escapeHtmlFunclocGap(g.max)}</code> | Present (distinct non-LORY suffixes): ${
    g.presentCount
  }</p>
  <p>Missing suffixes: <code>${escapeHtmlFunclocGap(g.missing.join(', '))}</code></p>
  <p>Non-LORY min / max FUNCLOC (by suffix): <code>${escapeHtmlFunclocGap(g.firstPresentFuncLoc)}</code> … <code>${escapeHtmlFunclocGap(g.lastPresentFuncLoc)}</code></p>
</div>
</details>`,
    )
    .join('\n');

  const missingParentHtml = payload.missingImmediateParents
    .map((m) => {
      const rowsHtml = m.children
        .map(
          (c) => `<tr>
<td>${c.row}</td>
<td><code>${escapeHtmlFunclocGap(c.funcLoc)}</code></td>
<td>${escapeHtmlFunclocGap(c.desc)}</td>
</tr>`,
        )
        .join('\n');
      return `<details>
<summary><strong>${escapeHtmlFunclocGap(m.missingParentFuncLoc)}</strong> - children found: ${m.childCount}</summary>
<div class="box">
  <p>Expected parent level: ${m.expectedLevel} | Child levels: ${escapeHtmlFunclocGap(m.childLevels.join(', '))}</p>
  <div class="table-wrap">
    <table>
      <thead><tr><th>Row</th><th>Child FUNCLOC</th><th>Description</th></tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  </div>
</div>
</details>`;
    })
    .join('\n');

  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Bah Jambi FUNCLOC Gap Report (fixed master)</title>
  <style>
    body{font-family:Arial,sans-serif;margin:18px;background:#fafafa;color:#111}
    h1,h2{margin:0 0 8px}
    .meta{color:#555;margin:0 0 12px}
    .cards{display:flex;gap:10px;flex-wrap:wrap;margin:12px 0 16px}
    .card{background:#fff;border:1px solid #ddd;border-radius:8px;padding:8px 10px}
    details{margin:8px 0;padding:8px;border:1px solid #ddd;border-radius:8px;background:#fff}
    summary{cursor:pointer}
    .box{margin-top:8px}
    .table-wrap{overflow:auto}
    table{border-collapse:collapse;min-width:900px;width:100%}
    th,td{border:1px solid #ddd;padding:6px 8px;font-size:12px;text-align:left;vertical-align:top}
    th{background:#f1f5f9}
    code{font-family:ui-monospace, SFMono-Regular, Menlo, monospace;font-size:11px}
  </style>
</head>
<body>
  <h1>Bah Jambi FUNCLOC Gap Report</h1>
  <p class="meta">Source sheet: ${escapeHtmlFunclocGap(sourceUrl)}<br/>
  Fixed workbook scanned: <code>${escapeHtmlFunclocGap(fixedWorkbookPath)}</code><br/>
  Generated: ${escapeHtmlFunclocGap(new Date().toISOString())}</p>
  <div class="cards">
    <div class="card"><strong>Total rows:</strong> ${payload.summary.totalRows}</div>
    <div class="card"><strong>Numeric gap groups:</strong> ${payload.summary.numericGapGroups}</div>
    <div class="card"><strong>Missing immediate parents:</strong> ${payload.summary.missingImmediateParents}</div>
  </div>
  <h2>Numeric Sibling Gaps</h2>
  <p class="meta" style="margin:0 0 10px">Same idea as webapp “Fill/compact gaps”: LORY rows (description or EQKTU contains “lory”) reserve their suffix; gaps are for non-LORY siblings only, and a slot is not listed as missing if a LORY row uses that number.</p>
  ${gapsHtml || '<p>No numeric sibling gaps found.</p>'}
  <h2>Missing Parent With Existing Children</h2>
  ${missingParentHtml || '<p>No missing immediate parent detected.</p>'}
</body>
</html>`;

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, html, 'utf8');
};

const processWorkbookFromXlsxPath = async ({ xlsxPath, sourceUrl, outDir, fixMapPath, outputPrefix }) => {
  const loaded = await loadWorkbookRows(xlsxPath);
  const rows = loaded.rows;
  const { mapping, mappingPath } = loadFixMapping(fixMapPath);
  const builtinMoves = loadBuiltinFuncLocMoves();
  const mappingWithDraftDefaults = mergeBuiltinFuncLocMovesIntoMapping(mapping, builtinMoves);
  applyFuncLocReplacementsToRowsInPlace(rows, mappingWithDraftDefaults.funcLocReplacements);
  console.log(
    `Built-in FUNCLOC moves merged: ${builtinMoves.length} (Bah Jambi fix-map wins on key overlap).`,
  );

  const suffixShiftCompact = applySingleHoleSuffixShiftCompact({
    worksheet: loaded.worksheet,
    rows: loaded.rows,
    header: loaded.header,
    mapping: mappingWithDraftDefaults,
  });
  const middleGapFill = applyMiddleGapSubtreeFill({
    worksheet: loaded.worksheet,
    rows: loaded.rows,
    header: loaded.header,
    mapping: mappingWithDraftDefaults,
  });

  // Analysis uses a snapshot of rows + mapping only (does not read the worksheet after this).
  const { decisions, childrenAudit, summary, mappingUsage, maxChildrenLoserFixes, uniquenessReport } =
    analyzeLevel5(rows, mappingWithDraftDefaults);

  const stamp = getJakartaStamp();
  const prefix = outputPrefix ? `${safeSlug(outputPrefix)}_` : '';
  const outputXlsx = path.join(outDir, `${prefix}bah_jambi_level5_master_fixed_${stamp}.xlsx`);
  const outputHtml = path.join(outDir, `${prefix}bah_jambi_level5_checker_${stamp}.html`);
  const outputJson = path.join(outDir, `${prefix}bah_jambi_level5_checker_${stamp}.json`);
  const outputTwoNoHtml = path.join(outDir, `${prefix}bah_jambi_two_no_report_${stamp}.html`);
  const outputGapJson = path.join(outDir, `${prefix}bah_jambi_funcloc_gap_report_${stamp}.json`);
  const outputGapHtml = path.join(outDir, `${prefix}bah_jambi_funcloc_gap_report_${stamp}.html`);

  buildFixedMasterWorkbook({
    worksheet: loaded.worksheet,
    rows: loaded.rows,
    header: loaded.header,
    mapping: mappingWithDraftDefaults,
    mappingUsage,
    maxChildrenLoserFixes,
  });
  reorderWorksheetByFuncLoc(loaded.worksheet, loaded.header);
  resyncRowMetadataFromWorksheet(loaded.worksheet, loaded.header, loaded.rows);
  stripFormulasFromWorksheet(loaded.worksheet);
  applyFinalCleanStyleBolding(loaded.worksheet, loaded.header);
  fs.mkdirSync(path.dirname(outputXlsx), { recursive: true });
  await loaded.workbook.xlsx.writeFile(outputXlsx);

  const gapScanRows = await loadGapScanRowsFromFixedXlsx(outputXlsx);
  const funclocGapPayload = analyzeFunclocGaps(gapScanRows);
  fs.writeFileSync(
    outputGapJson,
    JSON.stringify(
      {
        sourceUrl,
        fixedWorkbookPath: outputXlsx,
        generatedAt: new Date().toISOString(),
        ...funclocGapPayload,
      },
      null,
      2,
    ),
    'utf8',
  );
  writeFunclocGapReportHtml({
    outputPath: outputGapHtml,
    sourceUrl,
    fixedWorkbookPath: outputXlsx,
    payload: funclocGapPayload,
  });

  writeHtmlReport({
    decisions,
    childrenAudit,
    summary,
    outputPath: outputHtml,
    sourceUrl,
  });
  writeJsonReport({
    decisions,
    childrenAudit,
    summary,
    outputPath: outputJson,
    sourceUrl,
    mappingPath,
    mapping: mappingWithDraftDefaults,
    mappingUsage,
    maxChildrenLoserFixes,
    uniquenessReport,
    middleGapFill,
    suffixShiftCompact,
    funclocGapScan: {
      fixedWorkbookPath: outputXlsx,
      summary: funclocGapPayload.summary,
      gapReportJson: outputGapJson,
      gapReportHtml: outputGapHtml,
    },
  });
  const twoNoAnalysis = analyzeTwoNoPatterns(loaded.rows);
  writeTwoNoHtmlReport({
    outputPath: outputTwoNoHtml,
    sourceUrl,
    analysis: twoNoAnalysis,
  });

  console.log('\nDone.');
  console.log(`Excel: ${outputXlsx}`);
  console.log(`HTML : ${outputHtml}`);
  console.log(`JSON : ${outputJson}`);
  console.log(`2NO  : ${outputTwoNoHtml}`);
  console.log(`Gaps JSON: ${outputGapJson}`);
  console.log(`Gaps HTML: ${outputGapHtml}`);
  console.log(`Groups: ${summary.totalGroups}`);
  console.log(`Max-children decisions: ${summary.maxChildrenWinner}`);
  console.log(`Tie same-children: ${summary.tieSameChildren}`);
  console.log(`Tie different-children: ${summary.tieDifferentChildren}`);
  console.log(`Children audit rows: ${childrenAudit.length}`);

  return {
    outputXlsx,
    outputHtml,
    outputJson,
    outputTwoNoHtml,
    outputGapJson,
    outputGapHtml,
  };
};

const main = async () => {
  const args = parseArgs();
  if (args.help) {
    console.log(`
Bah Jambi Level-5 Checker

Flags:
  --source <folderUrl>               Source Drive folder URL (default from config)
  --out <dir>                        Output directory
  --fix-map <path>                   Fix mapping JSON path
  --help, -h                         Show help
`);
    process.exit(0);
  }

  await initGoogleDrive(OAUTH_JSON);

  const { mode, outDir, fixMapPath, driveFolderUrl } = args;
  console.log('Bah Jambi Level-5 Checker');
  console.log(`Fix map: ${fixMapPath}`);
  console.log(`Out dir: ${outDir}`);
  console.log(`Mode: ${mode}`);

  if (mode === 'drive-folder') {
    const folderId = extractFolderIdFromUrl(driveFolderUrl || DEFAULT_DRIVE_FOLDER_URL);
    if (!folderId) throw new Error(`Invalid --drive-folder URL: ${driveFolderUrl}`);
    console.log(`Drive folder: ${driveFolderUrl}`);
    console.log(`Drive path filter : Top-level folder starts with "REGIONAL"`);

    const allFiles = await listFilesRecursively(folderId);
    const { driveFiles: excelFiles } = filterDriveSourceExcelFiles(allFiles, {
      allowedOwnerEmails: [],
    });
    const driveFiles = excelFiles.filter((f) => isRegionalPath(f.path));
    if (!driveFiles.length) throw new Error('No Excel files found in Drive folder.');

    const targetFolderName = `Level 5 Standardization_${getJakartaStamp()}`;
    const uploadRoot = await getOrCreateFolder(targetFolderName, folderId);
    const folderCache = new Map();

    const tempDir = createTempDir();
    try {
      const downloaded = await runWithConcurrency(driveFiles, DEFAULT_CONCURRENCY, async (df) => {
        const localPath = path.join(tempDir, df.path.replace(/\//g, path.sep));
        fs.mkdirSync(path.dirname(localPath), { recursive: true });
        await downloadFile(df.id, localPath);
        return { localPath, driveFile: df };
      });

      for (const { localPath, driveFile } of downloaded) {
        const label = driveFile.path || driveFile.name || driveFile.id;
        console.log(`\n--- Processing Drive file: ${label} ---`);
        const outputs = await processWorkbookFromXlsxPath({
          xlsxPath: localPath,
          sourceUrl: `https://drive.google.com/file/d/${driveFile.id}/view`,
          outDir,
          fixMapPath,
          outputPrefix: label,
        });
        const uploadRelativeDir = path.posix.dirname(String(driveFile.path || '').replace(/\\/g, '/'));
        const uploadFolderId = await ensureFolderPath(uploadRoot.id, uploadRelativeDir, folderCache);
        const uploadName = String(driveFile.name || path.basename(outputs.outputXlsx));
        await uploadXlsxReplacing(
          outputs.outputXlsx,
          uploadName,
          uploadFolderId,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        );
        console.log(`Uploaded now: ${path.posix.join(uploadRelativeDir || '', uploadName)}`);
      }
    } finally {
      cleanupTempDir(tempDir);
    }
    return;
  }
};

main().catch((err) => {
  console.error(`Failed: ${err.message}`);
  process.exit(1);
});
