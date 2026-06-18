/**
 * check_lsmw_maintenance_plan.cjs
 *
 * Bandingkan SZAEH maintenance plan (Drive) vs CountReadng CDS (ZPP jam jalan).
 *
 *   pnpm run check-lsmw-maintenance-plan
 *   node check_lsmw_maintenance_plan.cjs \
 *     --maintenance-plan-link <url> [--cds-path ./cds_zpp_jamjalan_eqv.xlsx] [--ik07-path ./measuring-item-ik07.xlsx]
 *
 * Output: Output/check-lsmw-maintenance-plan/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx
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
const {
    DEFAULT_IK07_PATH,
    loadMeasuringPointByEquipment
} = require("./utils/measuring_point_ik07_mapping.cjs");
const {
    writeRegionalAndAllPksAggregates
} = require("./utils/lsmw_pks_aggregates.cjs");
const {
    listMeasuringDocumentPksEntries,
    pksXlsxFileName
} = require("./utils/measuring_document_drive_entries.cjs");
const {
    loadEquipmentFuncLocIndex,
    DEFAULT_MAPPING_PATH
} = require("./utils/equipment_number_mapping.cjs");
const {
    buildPointToEquipmentsMap,
    loadCdsEntriesByEquipment,
    extractMaintenancePlanRowsFromPath,
    buildMaintenancePlanCheckRows,
    writeCheckMaintenancePlanExcel
} = require("./utils/check_lsmw_maintenance_plan.cjs");

const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const OUTPUT_ROOT = path.join("Output", "check-lsmw-maintenance-plan");
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
        maintenancePlanFolderId: null,
        cdsPath: DEFAULT_CDS_PATH,
        ik07Path: DEFAULT_IK07_PATH
    };

    const planIdx = argv.indexOf("--maintenance-plan-link");
    if (planIdx !== -1 && argv[planIdx + 1]) {
        out.maintenancePlanFolderId = extractFolderIdFromUrl(argv[planIdx + 1]);
    }

    const cdsIdx = argv.indexOf("--cds-path");
    if (cdsIdx !== -1 && argv[cdsIdx + 1]) {
        out.cdsPath = argv[cdsIdx + 1];
    }

    const ik07Idx = argv.indexOf("--ik07-path");
    if (ik07Idx !== -1 && argv[ik07Idx + 1]) {
        out.ik07Path = argv[ik07Idx + 1];
    }

    return out;
}

/**
 * @param {ReturnType<typeof parseCli>} cli
 */
async function resolveMaintenancePlanFolderId(cli) {
    if (cli.maintenancePlanFolderId) return cli.maintenancePlanFolderId;

    const url = await input({
        message:
            "URL folder Google Drive — output lsmw-maintenance-plan (REGIONAL N/):"
    });
    return extractFolderIdFromUrl(String(url).trim());
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

async function main() {
    const cli = parseCli(process.argv.slice(2));
    const stamp = wibTimestampReadable();
    const outputDir = path.join(OUTPUT_ROOT, stamp);
    let exitCode = 0;

    const statusCounts = {
        match: 0,
        mismatch: 0,
        cds_missing: 0,
        no_equipment: 0
    };
    let totalRows = 0;
    let filesWritten = 0;

    prepareTempDownloads();

    try {
        const maintenancePlanFolderId =
            await resolveMaintenancePlanFolderId(cli);

        await initDrive();

        console.log("Folder maintenance plan →", maintenancePlanFolderId);
        console.log("CDS →", path.resolve(cli.cdsPath));
        console.log("IK07 →", path.resolve(cli.ik07Path));
        console.log("Output lokal →", path.resolve(outputDir));

        const ik07Mapping = loadMeasuringPointByEquipment(cli.ik07Path);
        const pointToEquipments = buildPointToEquipmentsMap(
            ik07Mapping.byEquipment
        );
        console.log(
            `Index IK07: ${ik07Mapping.byEquipment.size} equipment, ${pointToEquipments.size} measuring point`
        );

        const equipmentFuncLocIndex = loadEquipmentFuncLocIndex(
            DEFAULT_MAPPING_PATH
        );
        console.log(
            `Index new-equipment-number: ${equipmentFuncLocIndex.byEquipment.size} equipment, ${equipmentFuncLocIndex.totalRows} baris`
        );

        const cdsByEquipment = await loadCdsEntriesByEquipment(cli.cdsPath);
        let cdsTotalRows = 0;
        let cdsDuplicateEquipment = 0;
        for (const entries of cdsByEquipment.values()) {
            cdsTotalRows += entries.length;
            if (entries.length > 1) cdsDuplicateEquipment += 1;
        }
        console.log(
            `Index CDS: ${cdsByEquipment.size} equipment, ${cdsTotalRows} baris CDS total, ${cdsDuplicateEquipment} equipment duplikat`
        );

        const entries = await listMeasuringDocumentPksEntries({
            drive,
            rootFolderId: maintenancePlanFolderId,
            isRegionalFolderName,
            isProcessableSpreadsheet,
            skipMissingFolder: false
        });

        if (entries.length === 0) {
            console.error("Tidak ada file PKS yang diproses.");
            exitCode = 1;
            return;
        }

        console.log(`Scan: ${entries.length} file PKS`);

        /** @type {Map<string, Array<object>>} */
        const regionalRowsByName = new Map();
        /** @type {Array<object>} */
        const allRows = [];

        for (const entry of entries) {
            const label = `${entry.regional}/${pksXlsxFileName(entry.pksName)}`;
            console.log("Processing:", label);

            const { localFile } = await downloadSpreadsheetToTemp(
                drive,
                entry.file,
                TEMP_DOWNLOAD_PATH,
                `${entry.file.id}_`
            );
            const planRows = await extractMaintenancePlanRowsFromPath(localFile);
            const checkRows = buildMaintenancePlanCheckRows(
                entry.regional,
                entry.pksName,
                planRows,
                pointToEquipments,
                cdsByEquipment,
                equipmentFuncLocIndex
            );

            if (checkRows.length === 0) {
                console.log("  → 0 baris — dilewati");
                continue;
            }

            const outPath = path.join(
                outputDir,
                entry.regional,
                pksXlsxFileName(entry.pksName)
            );
            await writeCheckMaintenancePlanExcel(outPath, checkRows);

            if (!regionalRowsByName.has(entry.regional)) {
                regionalRowsByName.set(entry.regional, []);
            }
            regionalRowsByName.get(entry.regional).push(...checkRows);
            allRows.push(...checkRows);

            for (const row of checkRows) {
                statusCounts[row.status] = (statusCounts[row.status] || 0) + 1;
            }
            totalRows += checkRows.length;
            filesWritten += 1;

            const mismatch = checkRows.filter(r => r.status === "mismatch").length;
            console.log(
                `  → ${checkRows.length} baris (mismatch: ${mismatch})`
            );
        }

        await writeRegionalAndAllPksAggregates({
            runDir: outputDir,
            regionalRowsByName,
            allRows,
            label: "baris check",
            writeRows: (outPath, rows) =>
                writeCheckMaintenancePlanExcel(outPath, rows)
        });

        console.log("\nRingkasan:");
        console.log(`  Output: ${path.resolve(outputDir)}`);
        console.log(`  File PKS: ${filesWritten}`);
        console.log(`  Total baris: ${totalRows}`);
        console.log(`  match: ${statusCounts.match}`);
        console.log(`  mismatch: ${statusCounts.mismatch}`);
        console.log(`  cds_missing: ${statusCounts.cds_missing}`);
        console.log(`  no_equipment: ${statusCounts.no_equipment}`);
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
    parseCli
};
