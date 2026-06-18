/**
 * check_lsmw_update_change_equipment.cjs
 *
 * Baca output LSMW change equipment di folder Drive (subfolder REGIONAL*),
 * tulis ulang ke lokal dengan baris EQUNR duplikat (per file) di-fill merah.
 *
 *   pnpm run check-lsmw-update-change-equipment
 *   node check_lsmw_update_change_equipment.cjs --folder-link <url>
 *
 * Output: Output/check-lsmw-update-change-equipment/{timestamp WIB}/REGIONAL N/*.xlsx
 *         + summary.xlsx
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const ExcelJS = require("exceljs");
const {
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    parseCheckRowsCli,
    resolveCheckRowsFolderId,
    isRegionalFolderName
} = require("./utils/lsmw_cli.cjs");
const {
    assertEqunrHeaderRow1,
    readChangesEquipmentRows,
    findDuplicateRowIndexes,
    writeChangesEquipmentCheckWorkbook
} = require("./utils/check_lsmw_changes_equipment_dup.cjs");

const TEMPLATE_PATH = "./template_changes_equipment.xlsx";

const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const OUTPUT_ROOT = path.join("Output", "check-lsmw-update-change-equipment");

const SKIP_FILE_NAMES = new Set([
    "validation.xlsx",
    "column.json",
    "all-pks.xlsx",
    "summary.xlsx"
]);

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

function shouldSkipFileName(name) {
    const base = path.basename(String(name ?? ""));
    if (base.startsWith("~$")) return true;
    if (SKIP_FILE_NAMES.has(base.toLowerCase())) return true;
    return false;
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
 * @param {string} folderId
 * @param {string} regionalName
 * @param {string} [pathPrefix]
 * @returns {Promise<Array<{ id: string, name: string, mimeType: string, relativePath: string, regionalFolder: string, outputSubPath: string }>>}
 */
async function collectSpreadsheets(folderId, regionalName, pathPrefix = "") {
    const out = [];
    const res = await drive.files.list({
        q: `'${folderId}' in parents and trashed=false`,
        fields: "files(id,name,mimeType)",
        pageSize: 1000,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
    });

    for (const item of res.data.files || []) {
        if (item.mimeType === "application/vnd.google-apps.folder") {
            const childPrefix = pathPrefix
                ? `${pathPrefix}/${item.name}`
                : item.name;
            const nested = await collectSpreadsheets(
                item.id,
                regionalName,
                childPrefix
            );
            out.push(...nested);
            continue;
        }

        if (!isProcessableSpreadsheet(item)) continue;
        if (shouldSkipFileName(item.name)) continue;

        const relativePath = pathPrefix
            ? `${regionalName}/${pathPrefix}/${item.name}`
            : `${regionalName}/${item.name}`;

        const outputSubPath = pathPrefix
            ? path.join(regionalName, pathPrefix, item.name)
            : path.join(regionalName, item.name);

        out.push({
            id: item.id,
            name: item.name,
            mimeType: item.mimeType,
            relativePath,
            regionalFolder: regionalName,
            outputSubPath
        });
    }

    return out;
}

/**
 * @param {string} localPath
 * @param {string} outPath
 * @param {string} fileLabel
 * @returns {Promise<{ dataRowCount: number, duplicateEqunrCount: number, duplicateRowCount: number }>}
 */
async function processChangesEquipmentFile(localPath, outPath, fileLabel) {
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(localPath);

    const sourceSheet = sourceWorkbook.worksheets[0];
    if (!sourceSheet) {
        throw new Error(`Sheet not found: ${fileLabel}`);
    }

    const equnrColumn = assertEqunrHeaderRow1(sourceSheet, fileLabel);
    const rows = readChangesEquipmentRows(sourceSheet, equnrColumn);
    const { duplicateEqunrCount, duplicateRowIndexes } =
        findDuplicateRowIndexes(rows);

    await writeChangesEquipmentCheckWorkbook(
        TEMPLATE_PATH,
        rows,
        duplicateRowIndexes,
        outPath
    );

    return {
        dataRowCount: rows.length,
        duplicateEqunrCount,
        duplicateRowCount: duplicateRowIndexes.size
    };
}

/**
 * @param {Array<{ regional: string, file: string, dataRowCount: number, duplicateEqunrCount: number, duplicateRowCount: number }>} rows
 * @param {string} outPath
 */
async function writeSummaryExcel(rows, outPath) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Ringkasan", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    const headers = [
        "Regional",
        "File",
        "BarisData",
        "EqunrDuplikat",
        "BarisDuplikatMerah"
    ];
    ws.addRow(headers);
    ws.getRow(1).font = { bold: true };

    for (const row of rows) {
        ws.addRow([
            row.regional,
            row.file,
            row.dataRowCount,
            row.duplicateEqunrCount,
            row.duplicateRowCount
        ]);
    }

    ws.columns = [
        { width: 18 },
        { width: 48 },
        { width: 12 },
        { width: 16 },
        { width: 20 }
    ];

    if (rows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + rows.length, column: headers.length }
        };
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await wb.xlsx.writeFile(outPath);
    console.log("Ringkasan:", outPath);
}

async function main() {
    const cli = parseCheckRowsCli(process.argv.slice(2));
    const stamp = wibTimestampReadable();
    const outputDir = path.join(OUTPUT_ROOT, stamp);
    const summaryRows = [];
    let totalDuplicateRows = 0;
    let exitCode = 0;

    prepareTempDownloads();

    try {
        const rootFolderId = await resolveCheckRowsFolderId(cli);
        await initDrive();

        console.log("Folder Drive →", rootFolderId);
        console.log("Output lokal →", path.resolve(outputDir));

        const res = await drive.files.list({
            q: `'${rootFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
            fields: "files(id,name)",
            pageSize: 1000,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
        });

        const regionalFolders = (res.data.files || []).filter(f =>
            isRegionalFolderName(f.name)
        );

        if (regionalFolders.length === 0) {
            console.error("Tidak ada subfolder REGIONAL* di folder root.");
            exitCode = 1;
            return;
        }

        /** @type {Array<{ id: string, name: string, mimeType: string, relativePath: string, regionalFolder: string, outputSubPath: string }>} */
        let allFiles = [];
        for (const folder of regionalFolders) {
            console.log("Scanning:", folder.name);
            const files = await collectSpreadsheets(folder.id, folder.name);
            allFiles.push(...files);
        }

        allFiles.sort((a, b) =>
            a.relativePath.localeCompare(b.relativePath, "id")
        );

        if (allFiles.length === 0) {
            console.error("Tidak ada file spreadsheet yang diproses.");
            exitCode = 1;
            return;
        }

        for (const file of allFiles) {
            console.log("Processing:", file.relativePath);

            const { localFile } = await downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                `${file.id}_`
            );

            const outPath = path.join(outputDir, file.outputSubPath);
            const stats = await processChangesEquipmentFile(
                localFile,
                outPath,
                file.relativePath
            );

            totalDuplicateRows += stats.duplicateRowCount;

            summaryRows.push({
                regional: file.regionalFolder,
                file: file.relativePath,
                dataRowCount: stats.dataRowCount,
                duplicateEqunrCount: stats.duplicateEqunrCount,
                duplicateRowCount: stats.duplicateRowCount
            });

            console.log(
                `  ${stats.dataRowCount} baris data, ${stats.duplicateRowCount} baris duplikat merah (${stats.duplicateEqunrCount} EQUNR duplikat)`
            );
        }

        await writeSummaryExcel(
            summaryRows,
            path.join(outputDir, "summary.xlsx")
        );

        console.log(
            `\nSelesai: ${allFiles.length} file, ${totalDuplicateRows} baris duplikat ditandai merah.`
        );
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
    shouldSkipFileName,
    processChangesEquipmentFile,
    writeSummaryExcel
};
