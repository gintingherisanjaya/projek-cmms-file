import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export async function runWithConcurrency(items, concurrency, mapper) {
  const list = [...items];
  const limit = Math.max(1, Math.min(Number(concurrency) || 4, list.length || 1));
  const results = new Array(list.length);
  let i = 0;

  async function worker() {
    while (true) {
      const idx = i++;
      if (idx >= list.length) return;
      results[idx] = await mapper(list[idx], idx);
    }
  }

  await Promise.all(Array.from({ length: Math.min(limit, list.length || 1) }, () => worker()));
  return results;
}

export function createTempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'final-clean-'));
}

export function cleanupTempDir(dir) {
  if (!dir || !fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}
