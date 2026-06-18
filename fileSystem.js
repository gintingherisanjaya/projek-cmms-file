import fs from 'node:fs';
import path from 'node:path';

export const ensureDir = (dirPath) => fs.mkdirSync(dirPath, { recursive: true });

export const buildOutputPath = (inputPath) => {
  const swapped = inputPath.replace(
    `${path.sep}Data${path.sep}`,
    `${path.sep}Output${path.sep}Work-Center${path.sep}`,
  );
  if (swapped !== inputPath) return swapped;

  const dir = path.dirname(inputPath);
  const base = path.basename(inputPath);
  const parent = path.dirname(dir);
  return path.join(parent, 'Output', 'Work-Center', base);
};

export const buildValidationOutputPath = (inputPath) => {
  const swapped = inputPath.replace(
    `${path.sep}Data${path.sep}`,
    `${path.sep}Output${path.sep}Validation${path.sep}`,
  );
  if (swapped !== inputPath) return swapped;

  const dir = path.dirname(inputPath);
  const base = path.basename(inputPath);
  const parent = path.dirname(dir);
  return path.join(parent, 'Output', 'Validation', base);
};

export const isExcelFile = (filename) => {
  const lower = filename.toLowerCase();
  const base = path.basename(filename);
  return lower.endsWith('.xlsx') && !base.startsWith('~$');
};

export const collectExcelFiles = (rootDir) => {
  const files = [];
  const queue = [rootDir];

  while (queue.length) {
    const current = queue.shift();
    const entries = fs.readdirSync(current, { withFileTypes: true });

    // Sort entries by name to ensure consistent processing order
    // Use numeric sorting for directory names containing numbers
    const sortedEntries = entries.sort((a, b) => {
      // Extract numbers from directory names for proper numeric sorting
      const numA = a.name.match(/\d+/)?.[0] ? parseInt(a.name.match(/\d+/)[0]) : 0;
      const numB = b.name.match(/\d+/)?.[0] ? parseInt(b.name.match(/\d+/)[0]) : 0;

      // If both have numbers, sort numerically by the number
      if (!isNaN(numA) && !isNaN(numB)) {
        return numA - numB;
      }
      // Otherwise fall back to locale-aware string comparison
      return a.name.localeCompare(b.name);
    });

    sortedEntries.forEach((entry) => {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) queue.push(full);
      else if (entry.isFile() && isExcelFile(entry.name)) files.push(full);
    });
  }

  return files;
};
