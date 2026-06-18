function ownerEmails(file) {
  const owners = file.owners || [];
  return owners.map((o) => String(o.emailAddress || '').toLowerCase()).filter(Boolean);
}

/** Keep only `.xlsx` (not temp `~$`), optionally restricted to Drive owners. */
export function filterDriveSourceExcelFiles(allFiles, options = {}) {
  const {
    allowedOwnerEmails = [],
    logInfo,
    formatOwnerExcludedCount = (n) => `Filtered out ${n} file(s) by owner rules`,
  } = options;

  const allowSet = allowedOwnerEmails.map((e) => String(e).toLowerCase()).filter(Boolean);

  let excludedOwner = 0;
  const driveFiles = [];

  for (const f of allFiles) {
    const name = String(f.name || '');
    if (!/\.xlsx$/i.test(name) || /^~\$/.test(name)) continue;

    if (allowSet.length) {
      const emails = ownerEmails(f);
      const ok = emails.some((e) => allowSet.includes(e));
      if (!ok) {
        excludedOwner += 1;
        continue;
      }
    }

    driveFiles.push(f);
  }

  if (excludedOwner > 0 && logInfo) {
    logInfo(formatOwnerExcludedCount(excludedOwner));
  }

  return { driveFiles, excludedOwnerCount: excludedOwner };
}
