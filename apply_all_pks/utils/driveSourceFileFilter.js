import path from 'node:path';

/**
 * @typedef {{ id: string, path?: string, name?: string, ownerEmail?: string }} DriveListFile
 */

/**
 * Keep files whose owner email matches one of allowedOwnerEmails (case-insensitive).
 * If allowedOwnerEmails is empty, returns all files.
 * @param {DriveListFile[]} allFiles
 * @param {string[] | null | undefined} allowedOwnerEmails
 * @returns {{ files: DriveListFile[], excludedCount: number }}
 */
export function filterDriveFilesByOwners(allFiles, allowedOwnerEmails) {
  const emails = (allowedOwnerEmails ?? [])
    .map((e) => String(e).trim().toLowerCase())
    .filter(Boolean);
  if (emails.length === 0) {
    return { files: allFiles, excludedCount: 0 };
  }
  const allowed = new Set(emails);
  const files = allFiles.filter((file) =>
    allowed.has(String(file.ownerEmail ?? '').trim().toLowerCase()),
  );
  return { files, excludedCount: allFiles.length - files.length };
}

/**
 * Exclude files living under known Compare-LPP output trees (uploaded results, LSMW exports).
 * Output folders often repeat "REGIONAL N" inside LSMW subfolders; those paths must not be treated as source.
 * @param {DriveListFile[]} files
 * @returns {{ files: DriveListFile[], excludedCount: number }}
 */
export function filterOutDrivePipelineOutputPaths(files) {
  const before = files.length;
  const out = files.filter((file) => !isPipelineGeneratedDrivePath(file.path));
  return { files: out, excludedCount: before - out.length };
}

/**
 * True if relative Drive path is under Final/Full validation results, Final process/clean results, or an LSMW_* folder.
 * @param {string | undefined} filePath
 */
export function isPipelineGeneratedDrivePath(filePath) {
  const normalized = String(filePath ?? '').replace(/\\/g, '/');
  const lower = normalized.toLowerCase();
  if (lower.includes('final_clean_results')) return true;
  if (lower.includes('final_process_results')) return true;
  if (lower.includes('full_validation_results')) return true;
  const segments = normalized
    .split('/')
    .map((s) => s.trim())
    .filter(Boolean);
  // Only folder segments — avoid excluding a workbook named e.g. LSMW_notes.xlsx under REGIONAL.
  const dirSegments = segments.length > 1 ? segments.slice(0, -1) : [];
  return dirSegments.some((seg) => /^lsmw_/i.test(seg));
}

/**
 * Keep files whose path's first segment under the source root matches REGIONAL (case-insensitive).
 * Use together with {@link filterOutDrivePipelineOutputPaths}: nested REGIONAL folders under LSMW outputs
 * still have a REGIONAL first segment when the listed root is inside an LSMW subfolder.
 * @param {DriveListFile[]} files
 * @returns {{ files: DriveListFile[], excludedCount: number }}
 */
export function filterDriveFilesRegionalRootOnly(files) {
  const before = files.length;
  const out = files.filter((file) => {
    const segments = String(file.path ?? '')
      .split('/')
      .map((segment) => segment.trim())
      .filter(Boolean);
    const topLevelFolder = segments[0] ?? '';
    return /regional/i.test(topLevelFolder);
  });
  return { files: out, excludedCount: before - out.length };
}

/**
 * Drop Excel lock/temp files (~$*.xlsx) often present when a workbook is open in Drive.
 * @param {DriveListFile[]} files
 * @returns {{ files: DriveListFile[], excludedCount: number }}
 */
export function filterOutDriveExcelTempFiles(files) {
  const before = files.length;
  const out = files.filter((file) => {
    const name = file.name ?? path.basename(String(file.path ?? ''));
    return !name.startsWith('~$');
  });
  return { files: out, excludedCount: before - out.length };
}

/**
 * Pipeline: optional owner filter → drop known result/LSMW output paths → REGIONAL root folder → exclude ~$ temps.
 * @param {DriveListFile[]} allFiles
 * @param {{
 *   allowedOwnerEmails?: string[] | null,
 *   logInfo?: (msg: string) => void,
 *   formatOwnerExcludedCount?: (excludedCount: number) => string,
 * }} options
 * @returns {{
 *   driveFiles: DriveListFile[],
 *   ownerExcluded: number,
 *   pipelineOutputExcluded: number,
 *   regionalExcluded: number,
 *   tempExcluded: number,
 * }}
 */
export function filterDriveSourceExcelFiles(allFiles, options = {}) {
  const logInfo = options.logInfo;

  const { files: afterOwner, excludedCount: ownerExcluded } = filterDriveFilesByOwners(
    allFiles,
    options.allowedOwnerEmails,
  );
  if (ownerExcluded > 0 && logInfo) {
    const msg =
      typeof options.formatOwnerExcludedCount === 'function'
        ? options.formatOwnerExcludedCount(ownerExcluded)
        : `Filtered out ${ownerExcluded} file(s) not matching configured Drive owner filter`;
    logInfo(msg);
  }

  const { files: afterPipeline, excludedCount: pipelineOutputExcluded } =
    filterOutDrivePipelineOutputPaths(afterOwner);
  if (pipelineOutputExcluded > 0 && logInfo) {
    logInfo(
      `Excluded ${pipelineOutputExcluded} Drive file(s) under result or LSMW_* output folders (not source workbooks).`,
    );
  }

  const { files: afterRegional, excludedCount: regionalExcluded } =
    filterDriveFilesRegionalRootOnly(afterPipeline);
  if (regionalExcluded > 0 && logInfo) {
    logInfo(`Excluded ${regionalExcluded} Drive file(s) whose root folder is not REGIONAL.`);
  }

  const { files: afterTemp, excludedCount: tempExcluded } =
    filterOutDriveExcelTempFiles(afterRegional);
  if (tempExcluded > 0 && logInfo) {
    logInfo(`Excluded ${tempExcluded} Excel temp file(s) (~$...) from processing.`);
  }

  return {
    driveFiles: afterTemp,
    ownerExcluded,
    pipelineOutputExcluded,
    regionalExcluded,
    tempExcluded,
  };
}
