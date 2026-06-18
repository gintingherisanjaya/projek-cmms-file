/**
 * equipment-gathering.cjs
 * Gathering LPP + terapkan kolom before ke Protect Y (Applied / Last Result) dalam satu run.
 * Cocokkan FUNCTLOC DESC dengan file Protect Y (similarity, 1:1), lalu isi kolom before di Protect Y.
 *
 *   pnpm run equipment-gathering
 *
 * Output 1: Output/equipment-gathering/{timestamp WIB}/
 *   REGIONAL N/{nama PKS}.xlsx, all-pks.xlsx
 * Output 2: Output/apply-equipment-gathering-to-last-result/{timestamp WIB}/
 *   REGIONAL N/{nama PKS}.xlsx (Protect Y dengan kolom before terisi)
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { readSpreadsheetRows } = require("./utils/equipment_excel_io.cjs");
const { authenticate } = require("@google-cloud/local-auth");
const { input } = require("@inquirer/prompts");
const {
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    extractFolderIdFromUrl,
    isRegionalFolderName
} = require("./utils/lsmw_cli.cjs");
const { findDataLayout } = require("./utils/lsmw_cell_fill.cjs");
const {
    extractGatherRows,
    extractAppliedMatchCandidates,
    countFunclocAfterKeys,
    appendAppliedColumn,
    findLppGatheringLayout,
    OUTPUT_HEADERS,
    FUNCTLOC_DESC_AFTER_LEVEL123_HEADER,
    listMissingSourceColumns,
    resolveFileRegionalPom,
    normalizeOutputRowsRegionalPom
} = require("./utils/equipment_gathering_columns.cjs");
const { matchLppRowsToApplied } = require("./utils/equipment_gathering_match.cjs");
const { writeGatheringExcel } = require("./utils/equipment_gathering_excel.cjs");
const { applyGatheringToProtectYFile } = require("./utils/apply_equipment_gathering_protecty.cjs");
const {
    buildGatheringIndexFromOutputRows,
    gatheringOutputRowsToParsedRows,
    filterOutputRowsByDuplicateAppliedKey,
    collectFailedGatheringRows
} = require("./utils/apply_equipment_gathering_io.cjs");

const COL_EQUIPMENT_NUMBER = OUTPUT_HEADERS.indexOf("EQUIPMENT NUMBER");
const COL_COST_CENTER_BEFORE = OUTPUT_HEADERS.indexOf("COST CENTER BEFORE");
const COL_COST_CENTER_AFTER = OUTPUT_HEADERS.indexOf("COST CENTER AFTER");
const COL_FUNCTLOC_AFTER_LPP = OUTPUT_HEADERS.indexOf(
    FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
);
const COL_FUNCTLOC_BEFORE = OUTPUT_HEADERS.indexOf("FUNCTLOC DESC. BEFORE");
const COL_EQKTU_BEFORE = OUTPUT_HEADERS.indexOf("EQKTU BEFORE");

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = ["https://www.googleapis.com/auth/drive"];
const GATHERING_OUTPUT_DIR = path.join("Output", "equipment-gathering");
const APPLY_OUTPUT_DIR = path.join(
    "Output",
    "apply-equipment-gathering-to-last-result"
);
const SUMMARY_FILE_NAME = "all-pks.xlsx";

function stopWithError(message) {
    console.error(`\n  [error] ${message}`);
    console.error("Proses dihentikan.");
    process.exit(1);
}

let drive;

function wibTimestampForFilename() {
    const parts = new Intl.DateTimeFormat("en-GB", {
        timeZone: "Asia/Jakarta",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false
    }).formatToParts(new Date());
    const get = type => parts.find(p => p.type === type)?.value ?? "00";
    return `${get("year")}-${get("month")}-${get("day")}_${get("hour")}-${get("minute")}-${get("second")} WIB`;
}

function pksXlsxFileName(sourceFileName) {
    const base = path
        .basename(String(sourceFileName ?? ""))
        .replace(/\.xlsx$/i, "")
        .trim();
    return `${base.replace(/[/\\?%*:|"<>]/g, "_")}.xlsx`;
}

function fileMatchKey(name) {
    return path
        .basename(String(name ?? ""))
        .trim()
        .toLowerCase()
        .replace(/\.xlsx$/i, "")
        .replace(/\s+/g, " ");
}

/** Fallback POM dari nama file (mis. "1.PKS AUR GADING.xlsx" → "PKS AUR GADING"). */
function normalizeCellKey(value) {
    if (value === null || value === undefined || value === "") return "";
    return String(value).trim();
}

/** Kunci duplikat: EQUIPMENT NUMBER + COST CENTER BEFORE. */
function compositeFromRow(row) {
    const equipmentNumber = normalizeCellKey(row[COL_EQUIPMENT_NUMBER]);
    if (!equipmentNumber) return null;
    return {
        equipmentNumber,
        costCenterBefore: normalizeCellKey(row[COL_COST_CENTER_BEFORE])
    };
}

function compositeKeyString({ equipmentNumber, costCenterBefore }) {
    return `${equipmentNumber}\0${costCenterBefore}`;
}

function formatCompositeLabel({ equipmentNumber, costCenterBefore }) {
    const cc = costCenterBefore || "(kosong)";
    return `${equipmentNumber} | COST CENTER BEFORE: ${cc}`;
}

/** Hitung kemunculan composite key per baris output. */
function countCompositeInRows(rows) {
    /** @type {Map<string, { equipmentNumber: string, costCenterBefore: string, count: number }>} */
    const counts = new Map();
    for (const row of rows) {
        const comp = compositeFromRow(row);
        if (!comp) continue;
        const key = compositeKeyString(comp);
        const prev = counts.get(key);
        if (prev) {
            prev.count += 1;
        } else {
            counts.set(key, { ...comp, count: 1 });
        }
    }
    return counts;
}

function trackWithinPksDuplicates(rows, meta, withinPksDuplicates) {
    for (const entry of countCompositeInRows(rows).values()) {
        if (entry.count <= 1) continue;
        withinPksDuplicates.push({
            equipmentNumber: entry.equipmentNumber,
            costCenterBefore: entry.costCenterBefore,
            count: entry.count,
            regional: meta.regional,
            pks: meta.pks
        });
    }
}

function trackCrossPksDuplicates(rows, meta, crossPksIndex) {
    const seenInFile = new Set();
    for (const row of rows) {
        const comp = compositeFromRow(row);
        if (!comp) continue;
        const key = compositeKeyString(comp);
        if (seenInFile.has(key)) continue;
        seenInFile.add(key);

        if (!crossPksIndex.has(key)) {
            crossPksIndex.set(key, {
                equipmentNumber: comp.equipmentNumber,
                costCenterBefore: comp.costCenterBefore,
                files: []
            });
        }
        crossPksIndex.get(key).files.push({
            regional: meta.regional,
            pks: meta.pks
        });
    }
}

function printDuplicateEquipmentReport(withinPksDuplicates, crossPksIndex) {
    const compositeLabel =
        "EQUIPMENT NUMBER + COST CENTER BEFORE";

    console.log(`\n--- Duplikat ${compositeLabel} (PKS yang sama) ---`);
    if (withinPksDuplicates.length === 0) {
        console.log("  (tidak ada)");
    } else {
        const sorted = [...withinPksDuplicates].sort(
            (a, b) =>
                a.equipmentNumber.localeCompare(b.equipmentNumber) ||
                a.costCenterBefore.localeCompare(b.costCenterBefore) ||
                a.regional.localeCompare(b.regional) ||
                a.pks.localeCompare(b.pks)
        );
        for (const item of sorted) {
            console.log(
                `  ${formatCompositeLabel(item)} ×${item.count} — [${item.regional}] ${item.pks}`
            );
        }
        console.log(`  Total: ${sorted.length} kombinasi`);
    }

    const crossPks = [];
    for (const entry of crossPksIndex.values()) {
        const unique = [
            ...new Map(
                entry.files.map(f => [`${f.regional}\0${f.pks}`, f])
            ).values()
        ];
        if (unique.length < 2) continue;
        crossPks.push({
            equipmentNumber: entry.equipmentNumber,
            costCenterBefore: entry.costCenterBefore,
            files: unique
        });
    }

    console.log(`\n--- Duplikat ${compositeLabel} (lintas PKS berbeda) ---`);
    if (crossPks.length === 0) {
        console.log("  (tidak ada)");
    } else {
        crossPks.sort(
            (a, b) =>
                a.equipmentNumber.localeCompare(b.equipmentNumber) ||
                a.costCenterBefore.localeCompare(b.costCenterBefore)
        );
        for (const item of crossPks) {
            const locations = item.files
                .map(f => `[${f.regional}] ${f.pks}`)
                .join(" | ");
            console.log(`  ${formatCompositeLabel(item)} — ${locations}`);
        }
        console.log(`  Total: ${crossPks.length} kombinasi`);
    }
}

function pomFromFileName(fileName) {
    const base = path
        .basename(String(fileName ?? ""))
        .replace(/\.xlsx$/i, "")
        .trim();
    const withoutPrefix = base.replace(/^\d+\.?\s*/, "").trim();
    return withoutPrefix || base;
}

async function gatherFromLocalFile(localPath, { regional, sourceFileName } = {}) {
    const read = await readSpreadsheetRows(localPath);
    if (!read.ok) {
        return { ok: false, reason: read.reason, rows: [] };
    }

    const rawRows = read.rawRows;
    const layout = findLppGatheringLayout(rawRows);
    const headerRow = rawRows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header tidak ditemukan", rows: [] };
    }

    const dataRows = rawRows.slice(layout.headerRowIndex + 1);
    const { rows, indices, skippedNonStas } = extractGatherRows(headerRow, dataRows, {
        regional,
        pom: pomFromFileName(sourceFileName)
    });

    const missingColumns = listMissingSourceColumns(indices);
    if (missingColumns.length > 0) {
        return {
            ok: false,
            reason: `Kolom tidak ditemukan di LPP: ${missingColumns.join(", ")}`,
            missingColumns,
            rows: []
        };
    }

    return {
        ok: true,
        rows,
        rowCount: rows.length,
        skippedNonStas,
        missingColumns: []
    };
}

async function readProtectYSheetRows(localPath) {
    const read = await readSpreadsheetRows(localPath);
    if (!read.ok) {
        return read;
    }

    const rawRows = read.rawRows;
    const layout = findDataLayout(rawRows);
    const headerRow = rawRows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header tidak ditemukan" };
    }

    const dataRows = rawRows.slice(layout.headerRowIndex + 1);
    return { ok: true, headerRow, dataRows };
}

async function readAppliedCandidates(localPath) {
    const sheet = await readProtectYSheetRows(localPath);
    if (!sheet.ok) return sheet;
    return extractAppliedMatchCandidates(sheet.headerRow, sheet.dataRows);
}

async function applyAppliedMatching(lppRows, appliedLocalPath) {
    const applied = await readAppliedCandidates(appliedLocalPath);
    if (!applied.ok) {
        return {
            ok: false,
            reason: applied.reason,
            rows: appendAppliedColumn(lppRows, new Map())
        };
    }

    const matches = matchLppRowsToApplied(
        lppRows,
        applied.candidates,
        COL_COST_CENTER_AFTER,
        COL_FUNCTLOC_AFTER_LPP,
        COL_COST_CENTER_BEFORE,
        COL_FUNCTLOC_BEFORE,
        COL_EQKTU_BEFORE
    );
    const rows = appendAppliedColumn(lppRows, matches);

    let matchedCount = 0;
    for (let i = 0; i < lppRows.length; i++) {
        if (matches.has(i)) matchedCount += 1;
    }

    return {
        ok: true,
        rows,
        matchedCount,
        unmatchedCount: lppRows.length - matchedCount
    };
}

async function buildAppliedFileIndex(rootFolderId) {
    const index = new Map();

    async function walk(folderId) {
        const children = await listFolderChildren(folderId);
        for (const file of children) {
            if (file.mimeType === "application/vnd.google-apps.folder") {
                await walk(file.id);
            } else if (isProcessableSpreadsheet(file)) {
                const key = fileMatchKey(file.name);
                if (!index.has(key)) index.set(key, file);
            }
        }
    }

    await walk(rootFolderId);
    return index;
}

async function promptFolderUrl(message) {
    return input({
        message,
        validate: value => {
            const trimmed = String(value ?? "").trim();
            if (!trimmed) return "URL tidak boleh kosong";
            try {
                extractFolderIdFromUrl(trimmed);
                return true;
            } catch {
                return "URL folder Drive tidak valid";
            }
        }
    });
}

async function initDrive() {
    if (drive) return;

    const credentials = JSON.parse(fs.readFileSync(OAUTH_PATH));
    const { client_id, client_secret, redirect_uris } = credentials.installed;
    const auth = new google.auth.OAuth2(
        client_id,
        client_secret,
        redirect_uris[0]
    );

    if (fs.existsSync(TOKEN_PATH)) {
        let token = JSON.parse(fs.readFileSync(TOKEN_PATH));
        auth.setCredentials(token);
        const stale =
            !token.access_token ||
            (token.expiry_date && token.expiry_date < Date.now() + 60_000);
        if (stale && token.refresh_token) {
            const { credentials: refreshed } = await auth.refreshAccessToken();
            token = {
                ...token,
                ...refreshed,
                refresh_token: refreshed.refresh_token || token.refresh_token
            };
            fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
            auth.setCredentials(token);
        }
    } else {
        const newAuth = await authenticate({
            keyfilePath: OAUTH_PATH,
            scopes: SCOPES
        });
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(newAuth.credentials));
        auth.setCredentials(newAuth.credentials);
    }

    drive = google.drive({ version: "v3", auth });
}

async function listFolderChildren(folderId) {
    const out = [];
    let pageToken;
    do {
        const res = await drive.files.list({
            q: `'${folderId}' in parents and trashed=false`,
            fields: "nextPageToken, files(id,name,mimeType)",
            pageSize: 1000,
            pageToken,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
        });
        out.push(...(res.data.files ?? []));
        pageToken = res.data.nextPageToken;
    } while (pageToken);
    return out;
}

async function listSpreadsheetsRecursive(folderId, out = []) {
    const children = await listFolderChildren(folderId);
    for (const file of children) {
        if (file.mimeType === "application/vnd.google-apps.folder") {
            await listSpreadsheetsRecursive(file.id, out);
        } else if (isProcessableSpreadsheet(file)) {
            out.push(file);
        }
    }
    return out;
}

/**
 * @returns {Array<{ regional: string, file: object }>}
 */
async function listRegionalSpreadsheets(rootFolderId) {
    const entries = [];
    const rootChildren = await listFolderChildren(rootFolderId);

    for (const child of rootChildren) {
        if (child.mimeType !== "application/vnd.google-apps.folder") continue;
        if (!isRegionalFolderName(child.name)) continue;

        const files = await listSpreadsheetsRecursive(child.id);
        for (const file of files) {
            entries.push({ regional: child.name.trim(), file });
        }
    }

    entries.sort(
        (a, b) =>
            a.regional.localeCompare(b.regional) ||
            a.file.name.localeCompare(b.file.name)
    );

    return entries;
}

async function main() {
    const stamp = wibTimestampForFilename();
    const gatheringRunDir = path.join(GATHERING_OUTPUT_DIR, stamp);
    const applyRunDir = path.join(APPLY_OUTPUT_DIR, stamp);

    prepareTempDownloads();

    try {
        await initDrive();

        const lppFolderUrl = await promptFolderUrl(
            "URL folder Google Drive — LPP Resource (subfolder REGIONAL*):"
        );
        const protectYFolderUrl = await promptFolderUrl(
            "URL folder Google Drive — Protect Y / Applied Last Result (subfolder REGIONAL*):"
        );

        const lppRootFolderId = extractFolderIdFromUrl(
            String(lppFolderUrl).trim()
        );
        const protectYRootFolderId = extractFolderIdFromUrl(
            String(protectYFolderUrl).trim()
        );

        console.log("Memuat indeks Protect Y (Applied)...");
        const protectYIndex = await buildAppliedFileIndex(protectYRootFolderId);
        console.log(`  ${protectYIndex.size} file`);

        console.log("Memuat file LPP Resource (subfolder REGIONAL*)...");
        const entries = await listRegionalSpreadsheets(lppRootFolderId);
        console.log(`  ${entries.length} spreadsheet`);

        if (entries.length === 0) {
            console.log("Tidak ada file untuk diproses.");
            return;
        }

        fs.mkdirSync(gatheringRunDir, { recursive: true });
        fs.mkdirSync(applyRunDir, { recursive: true });

        const allRows = [];
        let filesOk = 0;
        let filesSkipped = 0;
        let totalLppMatched = 0;
        let totalProtectYApplied = 0;
        const withinPksDuplicates = [];
        /** @type {Map<string, { equipmentNumber: string, costCenterBefore: string, files: Array<{ regional: string, pks: string }> }>} */
        const crossPksIndex = new Map();

        for (const { regional, file: lppFile } of entries) {
            const gatheringRegionalDir = path.join(gatheringRunDir, regional);
            const applyRegionalDir = path.join(applyRunDir, regional);
            fs.mkdirSync(gatheringRegionalDir, { recursive: true });
            fs.mkdirSync(applyRegionalDir, { recursive: true });
            const gatheringOutPath = path.join(
                gatheringRegionalDir,
                pksXlsxFileName(lppFile.name)
            );
            const applyOutPath = path.join(
                applyRegionalDir,
                pksXlsxFileName(lppFile.name)
            );

            console.log(`[${regional}] ${lppFile.name}`);

            const { localFile: lppLocal } = await downloadSpreadsheetToTemp(
                drive,
                lppFile,
                TEMP_DOWNLOAD_PATH,
                "lpp_"
            );

            const result = await gatherFromLocalFile(lppLocal, {
                regional,
                sourceFileName: lppFile.name
            });
            if (!result.ok) {
                if (result.missingColumns?.length > 0) {
                    stopWithError(result.reason);
                }
                filesSkipped += 1;
                console.log(`  [skipped] LPP: ${result.reason}`);
                await writeGatheringExcel(gatheringOutPath, []);
                continue;
            }

            if (result.skippedNonStas > 0) {
                console.log(
                    `  [filter] ${result.skippedNonStas} baris dilewati (COST CENTER BEFORE tanpa STAS)`
                );
            }

            const protectYFile = protectYIndex.get(fileMatchKey(lppFile.name));

            if (!protectYFile) {
                stopWithError(
                    `Protect Y tidak ditemukan untuk ${lppFile.name} — ${result.rows.length} baris LPP tanpa pasangan`
                );
            }

            const { localFile: protectYLocal } = await downloadSpreadsheetToTemp(
                drive,
                protectYFile,
                TEMP_DOWNLOAD_PATH,
                "py_"
            );
            const matched = await applyAppliedMatching(result.rows, protectYLocal);
            let outputRows = matched.rows;

            if (!matched.ok) {
                stopWithError(`Protect Y (matching): ${matched.reason}`);
            }

            if (matched.unmatchedCount > 0) {
                stopWithError(
                    `${matched.unmatchedCount} baris LPP tanpa pasangan Protect Y (${lppFile.name})`
                );
            }

            totalLppMatched += matched.matchedCount;
            console.log(`  [match LPP] ${matched.matchedCount} paired`);

            const protectYSheet = await readProtectYSheetRows(protectYLocal);
            const appliedKeyCounts = protectYSheet.ok
                ? countFunclocAfterKeys(protectYSheet.headerRow, protectYSheet.dataRows)
                : new Map();

            const filtered = filterOutputRowsByDuplicateAppliedKey(
                outputRows,
                appliedKeyCounts
            );
            outputRows = filtered.rows;

            if (filtered.removedLppDuplicates > 0) {
                console.log(
                    `  [warn] ${filtered.removedLppDuplicates} baris LPP duplikat dibuang (first wins, Applied unik)`
                );
            }
            if (filtered.duplicateKeysWarn.length > 0) {
                const preview = filtered.duplicateKeysWarn.slice(0, 5).join("; ");
                const more =
                    filtered.duplicateKeysWarn.length > 5 ? " …" : "";
                console.log(
                    `  [warn] Duplikat FUNCTLOC DESC. AFTER APPLIED: ${preview}${more}`
                );
            }

            const fileMeta = resolveFileRegionalPom(outputRows, {
                regional,
                pom: pomFromFileName(lppFile.name)
            });
            if (!fileMeta.regional || !fileMeta.pom) {
                stopWithError(
                    `REGIONAL atau POM kosong untuk ${lppFile.name} — wajib terisi di LPP atau dari nama file`
                );
            }
            normalizeOutputRowsRegionalPom(
                outputRows,
                fileMeta.regional,
                fileMeta.pom
            );

            const gathered = buildGatheringIndexFromOutputRows(outputRows);
            if (gathered.duplicateKeys.length > 0) {
                const preview = gathered.duplicateKeys.slice(0, 5).join("; ");
                const more = gathered.duplicateKeys.length > 5 ? " …" : "";
                console.log(
                    `  [warn] Index apply duplikat key (first wins): ${preview}${more}`
                );
            }

            await writeGatheringExcel(gatheringOutPath, outputRows);
            allRows.push(...outputRows);

            fs.copyFileSync(protectYLocal, applyOutPath);
            const protectYApply = await applyGatheringToProtectYFile(
                applyOutPath,
                gathered.index,
                fileMeta
            );

            if (!protectYApply.ok) {
                stopWithError(`Protect Y (apply before): ${protectYApply.reason}`);
            }

            const gatheringParsedRows =
                gatheringOutputRowsToParsedRows(outputRows);
            const failedApplyRows = collectFailedGatheringRows(
                gatheringParsedRows,
                protectYApply.appliedKeys ?? new Set()
            );

            if (failedApplyRows.length > 0) {
                stopWithError(
                    `${failedApplyRows.length} equipment gagal apply ke Protect Y (${lppFile.name}) — key tidak cocok di sheet Protect Y`
                );
            }

            const appliedCount = protectYApply.appliedKeys?.size ?? 0;
            const applyStats = protectYApply.stats ?? {};
            totalProtectYApplied += appliedCount;
            if (applyStats.skippedDuplicateProtectYRows > 0) {
                console.log(
                    `  [warn] ${applyStats.skippedDuplicateProtectYRows} baris Protect Y duplikat key dilewati (first wins)`
                );
            }
            const clearedBefore = applyStats.clearedBeforeRows ?? 0;
            console.log(
                `  [apply Protect Y] clear before: ${clearedBefore} baris, isi: ${appliedCount} equipment`
            );

            filesOk += 1;

            const duplicateMeta = { regional, pks: lppFile.name };
            trackWithinPksDuplicates(outputRows, duplicateMeta, withinPksDuplicates);
            trackCrossPksDuplicates(outputRows, duplicateMeta, crossPksIndex);

            console.log(
                `  [ok] ${outputRows.length} baris gathering → ${path.basename(gatheringOutPath)}`
            );
        }

        const summaryPath = path.join(gatheringRunDir, SUMMARY_FILE_NAME);
        await writeGatheringExcel(summaryPath, allRows);

        console.log("\nSelesai.");
        console.log(`  Berhasil: ${filesOk}, dilewati: ${filesSkipped}`);
        console.log(`  Baris LPP matched Protect Y: ${totalLppMatched}`);
        console.log(`  Equipment ter-apply ke Protect Y: ${totalProtectYApplied}`);
        console.log(`  Total baris gathering (all-pks): ${allRows.length}`);
        console.log(`  Output gathering: ${gatheringRunDir}`);
        console.log(`  Output Protect Y: ${applyRunDir}`);

        printDuplicateEquipmentReport(withinPksDuplicates, crossPksIndex);
    } finally {
        cleanupTempDownloads();
    }
}

if (require.main === module) {
    main().catch(err => {
        console.error(err.message || err);
        process.exit(1);
    });
}
