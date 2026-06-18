/**
 * apply-equipment-gathering-to-last-result.cjs
 *
 * Isi kolom before di Protect Y (Last Result) dari hasil equipment-gathering.
 * Kunci baris: FUNCTLOC DESC. AFTER APPLIED (gathering) = FUNCTLOC DESC. AFTER (Protect Y).
 *
 *   pnpm run apply-equipment-gathering-to-last-result
 *
 * Output: Output/apply-equipment-gathering-to-last-result/{timestamp WIB}/
 *   REGIONAL N/{nama PKS}.xlsx
 *   all-failed-apply-equipment.xlsx
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const { select, input } = require("@inquirer/prompts");
const {
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    extractFolderIdFromUrl,
    isRegionalFolderName
} = require("./utils/lsmw_cli.cjs");
const {
    fileMatchKey,
    resolveLocalPath,
    listLocalRegionalSpreadsheets,
    readGatheringRowsFromFile,
    buildGatheringIndexFromFile,
    collectFailedGatheringRows,
    buildLocalGatheringFileIndex,
    buildDriveGatheringFileIndex
} = require("./utils/apply_equipment_gathering_io.cjs");
const { applyGatheringToProtectYFile } = require("./utils/apply_equipment_gathering_protecty.cjs");
const { writeFailedApplyExcel } = require("./utils/apply_equipment_gathering_excel.cjs");
const { resolveFileRegionalPom } = require("./utils/equipment_gathering_columns.cjs");

const FAILED_APPLY_SUMMARY_NAME = "all-failed-apply-equipment.xlsx";

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OUTPUT_DIR = path.join(
    "Output",
    "apply-equipment-gathering-to-last-result"
);

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

function pomFromFileName(fileName) {
    const base = path
        .basename(String(fileName ?? ""))
        .replace(/\.xlsx$/i, "")
        .trim();
    const withoutPrefix = base.replace(/^\d+\.?\s*/, "").trim();
    return withoutPrefix || base;
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

function appendFailedRows(allFailedRows, regional, pks, failedGatheringRows) {
    for (const row of failedGatheringRows) {
        allFailedRows.push({
            regional,
            pks,
            values: row.values
        });
    }
}

async function main() {
    const stamp = wibTimestampForFilename();
    const runDir = path.join(OUTPUT_DIR, stamp);

    prepareTempDownloads();

    try {
        await initDrive();

        const gatheringSource = await select({
            message: "Sumber hasil equipment-gathering:",
            choices: [
                { name: "Google Drive (folder link)", value: "drive" },
                { name: "Folder lokal", value: "local" }
            ]
        });

        let gatheringRunRoot = null;
        let gatheringDriveFolderId = null;
        /** @type {Map<string, string>} */
        let localGatheringByKey = new Map();
        /** @type {Map<string, { regional: string, file: object }>} */
        let driveGatheringByKey = new Map();

        if (gatheringSource === "drive") {
            const gatheringUrl = await promptFolderUrl(
                "URL folder Google Drive — hasil equipment-gathering (run dengan subfolder REGIONAL*):"
            );
            gatheringDriveFolderId = extractFolderIdFromUrl(
                String(gatheringUrl).trim()
            );
            console.log("Memuat indeks file gathering (REGIONAL*)...");
            const gatherEntries = await listRegionalSpreadsheets(
                gatheringDriveFolderId
            );
            driveGatheringByKey = buildDriveGatheringFileIndex(gatherEntries);
            console.log(`  ${driveGatheringByKey.size} file gathering`);
        } else {
            const localPathInput = await input({
                message:
                    "Path folder hasil equipment-gathering (relatif ke project root atau absolut):",
                validate: value => {
                    const resolved = resolveLocalPath(String(value ?? "").trim());
                    if (!resolved) return "Path tidak boleh kosong";
                    if (!fs.existsSync(resolved)) {
                        return "Folder tidak ditemukan";
                    }
                    return true;
                }
            });
            gatheringRunRoot = resolveLocalPath(String(localPathInput).trim());
            const gatherEntries = listLocalRegionalSpreadsheets(
                gatheringRunRoot,
                isRegionalFolderName
            );
            localGatheringByKey = buildLocalGatheringFileIndex(gatherEntries);
            console.log(
                `Memuat indeks gathering lokal: ${localGatheringByKey.size} file`
            );
        }

        const protectYUrl = await promptFolderUrl(
            "URL folder Google Drive — Last Result PKS (Protect Y, subfolder REGIONAL*):"
        );
        const protectYRootId = extractFolderIdFromUrl(String(protectYUrl).trim());

        console.log("Memuat file Protect Y (REGIONAL*)...");
        const protectYEntries = await listRegionalSpreadsheets(protectYRootId);
        console.log(`  ${protectYEntries.length} spreadsheet`);

        if (protectYEntries.length === 0) {
            console.log("Tidak ada file Protect Y untuk diproses.");
            return;
        }

        fs.mkdirSync(runDir, { recursive: true });

        let filesOk = 0;
        let filesSkipped = 0;
        let totalApplied = 0;
        let totalFailedApply = 0;
        let totalUnmatched = 0;
        let totalEmptyKey = 0;
        const allDuplicateKeys = [];
        /** @type {Array<{ regional: string, pks: string, values: unknown[] }>} */
        const allFailedRows = [];

        for (const { regional, file: pyFile } of protectYEntries) {
            const key = fileMatchKey(pyFile.name);
            const pksLabel = pyFile.name;
            console.log(`[${regional}] ${pyFile.name}`);

            let gatheringLocalPath = null;

            if (gatheringSource === "local") {
                gatheringLocalPath = localGatheringByKey.get(key);
                if (!gatheringLocalPath) {
                    filesSkipped += 1;
                    console.log("  [skipped] Gathering tidak ditemukan");
                    continue;
                }
            } else {
                const gatherEntry = driveGatheringByKey.get(key);
                if (!gatherEntry) {
                    filesSkipped += 1;
                    console.log("  [skipped] Gathering tidak ditemukan");
                    continue;
                }
                const { localFile } = await downloadSpreadsheetToTemp(
                    drive,
                    gatherEntry.file,
                    TEMP_DOWNLOAD_PATH,
                    "gather_"
                );
                gatheringLocalPath = localFile;
            }

            const gatheringRowsResult =
                await readGatheringRowsFromFile(gatheringLocalPath);
            if (!gatheringRowsResult.ok) {
                filesSkipped += 1;
                console.log(`  [skipped] Gathering: ${gatheringRowsResult.reason}`);
                continue;
            }

            const gathered = await buildGatheringIndexFromFile(gatheringLocalPath);
            if (!gathered.ok) {
                filesSkipped += 1;
                console.log(`  [skipped] Gathering: ${gathered.reason}`);
                continue;
            }

            const outputValues = gatheringRowsResult.rows.map(r => r.values);
            const fileMeta = resolveFileRegionalPom(outputValues, {
                regional,
                pom: pomFromFileName(pyFile.name)
            });
            if (!fileMeta.regional || !fileMeta.pom) {
                filesSkipped += 1;
                console.log(
                    "  [skipped] REGIONAL atau POM kosong — wajib diisi di gathering atau dari nama file"
                );
                continue;
            }

            if (gathered.duplicateKeys.length > 0) {
                allDuplicateKeys.push(
                    ...gathered.duplicateKeys.map(k => `${pyFile.name}: ${k}`)
                );
            }

            const { localFile: pyLocal } = await downloadSpreadsheetToTemp(
                drive,
                pyFile,
                TEMP_DOWNLOAD_PATH,
                "py_"
            );

            const regionalDir = path.join(runDir, regional);
            fs.mkdirSync(regionalDir, { recursive: true });
            const outPath = path.join(regionalDir, pksXlsxFileName(pyFile.name));
            fs.copyFileSync(pyLocal, outPath);

            const applied = await applyGatheringToProtectYFile(
                outPath,
                gathered.index,
                fileMeta
            );

            if (!applied.ok) {
                filesSkipped += 1;
                const failedRows = collectFailedGatheringRows(
                    gatheringRowsResult.rows,
                    new Set(),
                    { allFailed: true }
                );
                appendFailedRows(allFailedRows, regional, pksLabel, failedRows);
                totalFailedApply += failedRows.length;
                console.log(
                    `  [skipped] Protect Y: ${applied.reason} — ${failedRows.length} equipment → failed apply`
                );
                continue;
            }

            filesOk += 1;
            const appliedKeys = applied.appliedKeys ?? new Set();
            totalApplied += appliedKeys.size;
            totalUnmatched += applied.stats.unmatched;
            totalEmptyKey += applied.stats.emptyKey;

            const failedRows = collectFailedGatheringRows(
                gatheringRowsResult.rows,
                appliedKeys
            );
            appendFailedRows(allFailedRows, regional, pksLabel, failedRows);
            totalFailedApply += failedRows.length;

            console.log(
                `  [ok] applied ${appliedKeys.size}, failed apply ${failedRows.length} equipment`
            );
        }

        console.log("\nSelesai.");
        console.log(`  Berhasil: ${filesOk}, dilewati: ${filesSkipped}`);
        console.log(`  Equipment ter-apply: ${totalApplied}`);
        console.log(`  Equipment gagal apply: ${totalFailedApply}`);
        console.log(`  Baris Protect Y unmatched (key tidak di gathering): ${totalUnmatched}`);
        console.log(`  Baris Protect Y key kosong: ${totalEmptyKey}`);
        console.log(`  Output: ${runDir}`);

        if (totalFailedApply > 0) {
            const failedPath = path.join(runDir, FAILED_APPLY_SUMMARY_NAME);
            await writeFailedApplyExcel(failedPath, allFailedRows);
            console.log(`  Failed report: ${failedPath}`);
        }

        if (allDuplicateKeys.length > 0) {
            console.log(
                `\n--- Duplikat FUNCTLOC DESC. AFTER APPLIED di gathering (${allDuplicateKeys.length}) ---`
            );
            const show = allDuplicateKeys.slice(0, 20);
            for (const line of show) {
                console.log(`  ${line}`);
            }
            if (allDuplicateKeys.length > 20) {
                console.log(`  … dan ${allDuplicateKeys.length - 20} lainnya`);
            }
        }
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
