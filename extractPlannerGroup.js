import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  extractCostCenterSuffixForCsv,
  getRegionalColumnForMaintenancePlant,
  normalizeMaintenancePlantCode,
  regionColumnFromPlant,
} from './utils/regionalPlantMapping.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const csvData = fs.readFileSync(path.join(__dirname, 'utils', 'planner-group.csv'), 'utf8');

const splitCsvLine = (line) =>
  line.split(',').map((cell) => String(cell ?? '').replace(/\r/g, '').trim());

// Parse the raw CSV string into an array of row objects keyed by header names.
const [headers, ...rows] = csvData
  .trim()
  .split(/\r?\n/)
  .map(splitCsvLine);

const availableColumns = new Set(
  headers.map((h) =>
    String(h || '')
      .trim()
      .toUpperCase(),
  ),
);

const plannerGroupRows = rows.map((values) =>
  Object.fromEntries(headers.map((key, idx) => [key, (values[idx] || '').trim()])),
);

// Pre-index cost centers for faster lookup.
// Key: normalized template suffix (e.g. "STAS01" from "****STAS01").
const costCenterIndex = new Map();
plannerGroupRows.forEach((row) => {
  const raw = String(row.COSTCENTER ?? '')
    .trim()
    .toUpperCase();
  if (!raw) return;
  const key = raw.replace(/^\*{4}/, '');
  if (!key) return;
  if (!costCenterIndex.has(key)) costCenterIndex.set(key, row);
});

// Memoization caches
const regionColumnCache = new Map();
const findCostCenterRowCache = new Map();

const regionColumnFromPlantCached = (plantCode, regionOverride) => {
  const cacheKey = `${plantCode}|${regionOverride ?? ''}`;
  if (regionColumnCache.has(cacheKey)) return regionColumnCache.get(cacheKey);
  const result = regionColumnFromPlant(plantCode, availableColumns, regionOverride);
  regionColumnCache.set(cacheKey, result);
  return result;
};

export { getRegionalColumnForMaintenancePlant };

const findCostCenterRow = (plantCode, costCenter) => {
  const cacheKey = `${plantCode}|${costCenter}`;
  if (findCostCenterRowCache.has(cacheKey)) return findCostCenterRowCache.get(cacheKey);
  const suffix = extractCostCenterSuffixForCsv(costCenter);
  if (!suffix) {
    findCostCenterRowCache.set(cacheKey, null);
    return null;
  }
  const row = costCenterIndex.get(suffix) || null;
  findCostCenterRowCache.set(cacheKey, row || null);
  return row;
};

const normalizeCostCenter = (plantCode, costCenter) => {
  // Clean the cost center by removing unwanted characters like backticks and extra spaces
  const cleaned = String(costCenter ?? '')
    .replace(/[`]/g, '')
    .replace(/\s+/g, '');
  const upper = cleaned.trim().toUpperCase();
  if (!upper) return '';
  const plantUpper = String(plantCode ?? '')
    .trim()
    .toUpperCase();
  const withoutPlant =
    plantUpper && upper.startsWith(plantUpper) ? upper.slice(plantUpper.length) : upper;

  // Some LPP exports use station-style codes like "STAS01" which correspond to "STAS01"
  // in our planner group mapping (STAS01..STAS14). Keep as is.
  const stationNormalized = withoutPlant.replace(
    /^STAS(\d{2,3})$/,
    (_m, n) => `STAS${String(n).padStart(2, '0')}`,
  );
  const baseNormalized = plantUpper ? `${plantUpper}${stationNormalized}` : stationNormalized;

  // If the incoming code already embeds the plant code, keep it; otherwise add it.
  if (plantUpper && baseNormalized.startsWith(plantUpper)) return baseNormalized;
  return plantUpper ? `${plantUpper}${baseNormalized.replace(/^(\*{4})?/, '')}` : baseNormalized;
};

const derivePlantCode = (costCenter, funcLocation) => {
  // Clean the cost center by removing unwanted characters like backticks and extra spaces
  const cleaned = String(costCenter ?? '')
    .replace(/[`]/g, '')
    .replace(/\s+/g, '');
  const ccPart = cleaned.trim().toUpperCase();
  if (ccPart.length >= 4) return ccPart.slice(0, 4);
  return extractPlantCode(funcLocation);
};

export const getPlannerGroup = (
  funcLocation,
  costCenter,
  { verbose = false, region = null, maintenancePlant = null } = {},
) => {
  const plantCode = maintenancePlant
    ? normalizeMaintenancePlantCode(maintenancePlant)
    : derivePlantCode(costCenter, funcLocation);
  if (!plantCode) {
    if (verbose) console.log('Planner group: no plant code', { maintenancePlant, costCenter });
    return null;
  }
  let regionColumn = regionColumnFromPlantCached(plantCode, region);
  const row = findCostCenterRow(plantCode, costCenter);
  if (!row) {
    if (verbose) console.log('Cost center not found for plant', { plantCode, costCenter });
    return null;
  }

  // The CSV now has simple column names like "1", "9", "K", "2", etc.
  // We can directly use the regionColumn as the column name
  // If the computed region column doesn't exist in this CSV, fallback to "1".
  if (!Object.prototype.hasOwnProperty.call(row, regionColumn) && regionColumn !== '1') {
    regionColumn = '1';
  }

  const value = row[regionColumn] || '';
  if (verbose) {
    console.log('Selected Planner Group:', value || '(empty)', {
      plantCode,
      regionColumn,
      costCenter,
      maintenancePlant: maintenancePlant ?? '(derived)',
    });
  }
  return value;
};

const extractPlantCode = (funcloc) => String(funcloc ?? '').split('-')[1] || '';
