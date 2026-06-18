/**
 * check_lsmw_measuring_document.cjs
 *
 * Bandingkan measuring document (Drive utama + missing) vs CDS.
 *
 *   pnpm run check-lsmw-measuring-document
 *   node check_lsmw_measuring_document.cjs \
 *     --measuring-doc-link <url> --missing-link <url> [--cds-path ./cds_zpp_jamjalan_eqv.xlsx]
 *
 * Output: Output/check-lsmw-measuring-document/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx
 *         + {REGIONAL N}.xlsx + all-pks.xlsx
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
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
const { readSpreadsheetRows } = require("./utils/equipment_excel_io.cjs");
const {
    writeRegionalAndAllPksAggregates
} = require("./utils/lsmw_pks_aggregates.cjs");
const {
    listMeasuringDocumentPksEntries,
    pksXlsxFileName
} = require("./utils/measuring_document_drive_entries.cjs");
const {
    loadCdsMeasuringPointEquipmentIndex,
    extractMeasuringDocRows,
    dedupePairs,
    buildCheckRows,
    writeCheckMeasuringDocumentExcel
} = require("./utils/check_lsmw_measuring_document.cjs");

const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const OUTPUT_ROOT = path.join("Output", "check-lsmw-measuring-document");
const DEFAULT_CDS_PATH = "./cds_zpp_jamjalan_eqv.xlsx";

let drive;

function wibTimestampReadable() {
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

/**
 * @param {string[]} argv
 */
function parseCli(argv) {
    const out = {
        measuringDocFolderId: null,
        missingFolderId: null,
        cdsPath: DEFAULT_CDS_PATH
    };

    const measuringIdx = argv.indexOf("--measuring-doc-link");
    if (measuringIdx !== -1 && argv[measuringIdx + 1]) {
        out.measuringDocFolderId = extractFolderIdFromUrl(argv[measuringIdx + 1]);
    }

    const missingIdx = argv.indexOf("--missing-link");
    if (missingIdx !== -1 && argv[missingIdx + 1]) {
        out.missingFolderId = extractFolderIdFromUrl(argv[missingIdx + 1]);
    }

    const cdsIdx = argv.indexOf("--cds-path");
    if (cdsIdx !== -1 && argv[cdsIdx + 1]) {
        out.cdsPath = argv[cdsIdx + 1];
    }

    return out;
}

/**
 * @param {ReturnType<typeof parseCli>} cli
 */
async function resolveFolderIds(cli) {
    let measuringDocFolderId = cli.measuringDocFolderId;
    let missingFolderId = cli.missingFolderId;

    if (!measuringDocFolderId) {
        const url = await input({
            message:
                "URL folder Google Drive — output measuring document (REGIONAL N/):"
        });
        measuringDocFolderId = extractFolderIdFromUrl(String(url).trim());
    }

    if (!missingFolderId) {
        const url = await input({
            message:
                "URL folder Google Drive — Missing Measuring Document (REGIONAL N/):"
        });
        missingFolderId = extractFolderIdFromUrl(String(url).trim());
    }

    return { measuringDocFolderId, missingFolderId };
}

async function initDrive() {
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
            console.log("Refreshing token...");
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
            scopes: SCOPES,
            keyfilePath: OAUTH_PATH
        });
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(newAuth.credentials));
        auth.setCredentials(newAuth.credentials);
    }

    drive = google.drive({ version: "v3", auth });
}

/**
 * @param {string} rootFolderId
 * @param {boolean} skipMissingFolder
 */
async function scanFolder(rootFolderId, skipMissingFolder) {
    return listMeasuringDocumentPksEntries({
        drive,
        rootFolderId,
        isRegionalFolderName,
        isProcessableSpreadsheet,
        skipMissingFolder
    });
}

/**
 * @param {Array<{ regional: string, pksName: string, file: object }>} entries
 */
function mergeEntriesByRegionalPks(entries) {
    /** @type {Map<string, { regional: string, pksName: string, files: object[] }>} */
    const byKey = new Map();

    for (const entry of entries) {
        const key = `${entry.regional}|${entry.pksName}`;
        if (!byKey.has(key)) {
            byKey.set(key, {
                regional: entry.regional,
                pksName: entry.pksName,
                files: []
            });
        }
        byKey.get(key).files.push(entry.file);
    }

    return [...byKey.values()].sort((a, b) => {
        const ra = a.regional.localeCompare(b.regional, "id");
        if (ra !== 0) return ra;
        return a.pksName.localeCompare(b.pksName, "id");
    });
}

async function main() {
    const cli = parseCli(process.argv.slice(2));
    const stamp = wibTimestampReadable();
    const outputDir = path.join(OUTPUT_ROOT, stamp);
    let exitCode = 0;

    let totalRows = 0;
    let totalFinded = 0;
    let totalMissing = 0;
    let filesWritten = 0;

    prepareTempDownloads();

    try {
        const { measuringDocFolderId, missingFolderId } =
            await resolveFolderIds(cli);

        await initDrive();

        console.log("Folder measuring document →", measuringDocFolderId);
        console.log("Folder missing →", missingFolderId);
        console.log("CDS →", path.resolve(cli.cdsPath));
        console.log("Output lokal →", path.resolve(outputDir));

        const cdsIndex = await loadCdsMeasuringPointEquipmentIndex(cli.cdsPath);
        console.log(`Index CDS: ${cdsIndex.size} pasangan (Measuring point + Equipment)`);

        const mainEntries = await scanFolder(measuringDocFolderId, true);
        const missingEntries = await scanFolder(missingFolderId, false);
        const merged = mergeEntriesByRegionalPks([
            ...mainEntries,
            ...missingEntries
        ]);

        if (merged.length === 0) {
            console.error("Tidak ada file PKS yang diproses.");
            exitCode = 1;
            return;
        }

        console.log(
            `Scan: ${mainEntries.length} file utama + ${missingEntries.length} file missing → ${merged.length} PKS unik`
        );

        /** @type {Map<string, Array<object>>} */
        const regionalRowsByName = new Map();
        /** @type {Array<object>} */
        const allRows = [];

        for (const item of merged) {
            const label = `${item.regional}/${pksXlsxFileName(item.pksName)}`;
            console.log("Processing:", label);

            /** @type {Array<{ measuringPoint: string, equipment: string }>} */
            let allPairs = [];

            for (const file of item.files) {
                const { localFile } = await downloadSpreadsheetToTemp(
                    drive,
                    file,
                    TEMP_DOWNLOAD_PATH,
                    `${file.id}_`
                );
                const sheet = await readSpreadsheetRows(localFile);
                if (!sheet.ok) {
                    console.warn(`  [skip] ${sheet.reason}`);
                    continue;
                }
                allPairs.push(...extractMeasuringDocRows(sheet.rawRows));
            }

            const pairs = dedupePairs(allPairs);
            const checkRows = buildCheckRows(
                item.regional,
                item.pksName,
                pairs,
                cdsIndex
            );

            if (checkRows.length === 0) {
                console.log("  → 0 baris — dilewati");
                continue;
            }

            const outPath = path.join(
                outputDir,
                item.regional,
                pksXlsxFileName(item.pksName)
            );
            await writeCheckMeasuringDocumentExcel(outPath, checkRows);

            const finded = checkRows.filter(r => r.status === "finded").length;
            const missing = checkRows.filter(r => r.status === "missing").length;

            if (!regionalRowsByName.has(item.regional)) {
                regionalRowsByName.set(item.regional, []);
            }
            regionalRowsByName.get(item.regional).push(...checkRows);
            allRows.push(...checkRows);

            totalRows += checkRows.length;
            totalFinded += finded;
            totalMissing += missing;
            filesWritten += 1;

            console.log(
                `  → ${checkRows.length} baris (finded: ${finded}, missing: ${missing})`
            );
        }

        await writeRegionalAndAllPksAggregates({
            runDir: outputDir,
            regionalRowsByName,
            allRows,
            label: "baris check",
            writeRows: (outPath, rows) =>
                writeCheckMeasuringDocumentExcel(outPath, rows)
        });

        console.log("\nRingkasan:");
        console.log(`  Output: ${path.resolve(outputDir)}`);
        console.log(`  File PKS: ${filesWritten}`);
        console.log(`  Total baris: ${totalRows}`);
        console.log(`  Finded: ${totalFinded}, missing: ${totalMissing}`);
    } catch (err) {
        console.error(err.message || err);
        exitCode = 1;
    } finally {
        cleanupTempDownloads();
        if (exitCode !== 0) process.exit(exitCode);
    }
}

if (require.main === module) {
    main().catch(err => {
        console.error(err);
        process.exit(1);
    });
}

module.exports = {
    parseCli,
    mergeEntriesByRegionalPks
};
