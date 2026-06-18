import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  extractCostCenterSuffixForCsv,
  getRegionalColumnForMaintenancePlant,
  normalizeMaintenancePlantCode,
  regionColumnFromPlant,
} from '../utils/regionalPlantMapping.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DEFAULT_CWC_PATH = path.join(__dirname, '..', 'utils', 'cwc.csv');
const CWC_SPECIAL_DIR = path.join(__dirname, '..', 'utils', 'cwc_special');

const splitCsvLine = (line) =>
  line.split(',').map((cell) => String(cell ?? '').replace(/\r/g, '').trim());

export const availableColumns = new Set();

// Pre-index cost centers for faster lookup.
// Key: normalized template suffix (e.g. "STA501" from "****STA501").
const costCenterIndex = new Map();

function loadCwcFromFile(filePath) {
  const csvData = fs.readFileSync(filePath, 'utf8');
  const [headers, ...rows] = csvData
    .trim()
    .split(/\r?\n/)
    .map(splitCsvLine);

  availableColumns.clear();
  for (const h of headers) {
    availableColumns.add(
      String(h || '')
        .trim()
        .toUpperCase(),
    );
  }

  const cwcRows = rows.map((values) =>
    Object.fromEntries(headers.map((key, idx) => [key, (values[idx] || '').trim()])),
  );

  costCenterIndex.clear();
  cwcRows.forEach((row) => {
    const raw = String(row.COSTCENTER ?? '')
      .trim()
      .toUpperCase();
    if (!raw) return;
    const key = raw.replace(/^\*{4}/, '');
    if (!key) return;
    if (!costCenterIndex.has(key)) costCenterIndex.set(key, row);
  });
}

export function resolveCwcCsvPath(maintenancePlant) {
  const plantCode = normalizeMaintenancePlantCode(maintenancePlant);
  if (plantCode) {
    const specialPath = path.join(CWC_SPECIAL_DIR, `${plantCode}.csv`);
    if (fs.existsSync(specialPath)) return specialPath;
  }
  return DEFAULT_CWC_PATH;
}

export function setCwcSourceForMaintenancePlant(maintenancePlant) {
  const csvPath = resolveCwcCsvPath(maintenancePlant);
  loadCwcFromFile(csvPath);
  regionColumnCache.clear();
  findCostCenterRowCache.clear();
  return csvPath;
}

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
  // in our CWC mapping (STAS01..STAS14). Keep as is.
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

export const getCWC = (
  funcLocation,
  costCenter,
  { verbose = false, region = null, maintenancePlant = null } = {},
) => {
  const plantCode = maintenancePlant
    ? normalizeMaintenancePlantCode(maintenancePlant)
    : derivePlantCode(costCenter, funcLocation);
  if (!plantCode) {
    if (verbose) console.log('CWC: no plant code', { maintenancePlant, costCenter });
    return null;
  }
  let regionColumn = regionColumnFromPlantCached(plantCode, region);
  const row = findCostCenterRow(plantCode, costCenter);
  if (!row) {
    if (verbose) console.log('Cost center not found for plant', { plantCode, costCenter });
    return null;
  }

  // If the computed region column doesn't exist in this CSV, fallback to "1".
  if (!Object.prototype.hasOwnProperty.call(row, regionColumn) && regionColumn !== '1') {
    regionColumn = '1';
  }

  const value = row[regionColumn] || '';
  if (verbose) {
    console.log('Selected CWC:', value || '(empty)', {
      plantCode,
      regionColumn,
      costCenter,
      maintenancePlant: maintenancePlant ?? '(derived)',
    });
  }
  return value;
};

const extractPlantCode = (funcloc) => String(funcloc ?? '').split('-')[1] || '';

loadCwcFromFile(DEFAULT_CWC_PATH);
