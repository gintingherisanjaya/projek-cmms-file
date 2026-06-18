import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '..');

/** JSON map FUNCLOC_SUFFIX → ABC indicator (optional Reference/abc-map.json). */
export function loadAbcMap() {
  const fp = path.join(PROJECT_ROOT, 'Reference', 'abc-map.json');
  if (!fs.existsSync(fp)) return new Map();
  try {
    const raw = JSON.parse(fs.readFileSync(fp, 'utf8'));
    if (raw && typeof raw === 'object' && !(raw instanceof Array)) {
      return new Map(
        Object.entries(raw).map(([k, v]) => [String(k).trim().toUpperCase(), String(v).trim()]),
      );
    }
  } catch {
    /* empty */
  }
  return new Map();
}

/** Last segment uppercase for map key. */
export function normalizeFunclocSuffix(funcloc) {
  const parts = String(funcloc ?? '')
    .split('-')
    .map((p) => p.trim())
    .filter(Boolean);
  const last = parts[parts.length - 1];
  return last ? last.toUpperCase() : '';
}
