/**
 * check_number_of_line_in_same_template.cjs
 *
 * Bandingkan jumlah baris data nyata semua spreadsheet di folder Drive (subfolder REGIONAL*).
 * Mode terminate (exit 1) atau warning saja jika jumlah baris berbeda dari file pertama.
 *
 *   pnpm run check-number-of-line-in-same-template
 *   node check_number_of_line_in_same_template.cjs --folder-link <url>
 *   node check_number_of_line_in_same_template.cjs --on-mismatch warn
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const { select } = require("@inquirer/prompts");
const XLSX = require("xlsx");
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
const { countRealDataRows } = require("./utils/count_real_data_rows.cjs");

const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const OUTPUT_DIR = path.join(
    "Output",
    "check-number-of-line-in-same-tenplate"
);
const AGGREGATE_SKIP = "all-pks.xlsx";

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
    if (base.toLowerCase() === AGGREGATE_SKIP) return true;
    return false;
}

async function promptMismatchMode() {
    return select({
        message: "Jika jumlah baris berbeda dari file pertama (baseline):",
        choices: [
            {
                name: "Terminate — hentikan proses (exit 1) setelah scan selesai",
                value: "terminate"
            },
            {
                name: "Warning saja — tampilkan peringatan, Excel lengkap, exit 0",
                value: "warn"
            }
        ]
    });
}

/**
 * @param {{ onMismatch?: string | null }} cli
 */
async function resolveMismatchMode(cli) {
    if (cli?.onMismatch === "terminate" || cli?.onMismatch === "warn") {
        return cli.onMismatch;
    }
    return promptMismatchMode();
}

/**
 * @param {number} mismatchCount
 * @param {"terminate"|"warn"} mismatchMode
 */
function exitCodeForMismatch(mismatchCount, mismatchMode) {
    if (mismatchCount > 0 && mismatchMode === "terminate") return 1;
    return 0;
}

/**
 * @param {Array<{ rowCount: number, relativePath: string, fileName: string, subFolder: string }>} reportRows
 * @param {number|null} expectedCount
 * @param {string|null} baselinePath
 */
function annotateReportRows(reportRows, expectedCount, baselinePath) {
    if (expectedCount === null) return reportRows;

    return reportRows.map((row, index) => {
        const isBaseline = row.relativePath === baselinePath || index === 0;
        let status = "beda";
        if (isBaseline) status = "baseline";
        else if (row.rowCount === expectedCount) status = "cocok";
        return {
            ...row,
            baseline: expectedCount,
            status
        };
    });
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
 * @returns {Promise<Array<{ id: string, name: string, mimeType: string, relativePath: string, regionalFolder: string }>>}
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

        out.push({
            id: item.id,
            name: item.name,
            mimeType: item.mimeType,
            relativePath,
            regionalFolder: regionalName
        });
    }

    return out;
}

function readSheetRows(localPath) {
    const wb = XLSX.readFile(localPath, { cellDates: true });
    const sheetName = wb.SheetNames[0];
    if (!sheetName) return [];
    const sheet = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
}

/**
 * @param {Array<{ fileName: string, subFolder: string, rowCount: number, baseline?: number, status?: string }>} rows
 * @param {string} outPath
 */
async function writeReportExcel(rows, outPath) {
    fs.mkdirSync(path.dirname(outPath), { recursive: true });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Row counts");

    ws.columns = [
        { header: "nama file", key: "fileName", width: 40 },
        { header: "nama sub folder", key: "subFolder", width: 24 },
        { header: "jumlah baris", key: "rowCount", width: 14 },
        { header: "baseline", key: "baseline", width: 12 },
        { header: "status", key: "status", width: 12 }
    ];

    for (const row of rows) {
        ws.addRow({
            fileName: row.fileName,
            subFolder: row.subFolder,
            rowCount: row.rowCount,
            baseline: row.baseline ?? "",
            status: row.status ?? ""
        });
    }

    await wb.xlsx.writeFile(outPath);
    console.log("Laporan:", outPath);
}

async function main() {
    const cli = parseCheckRowsCli(process.argv.slice(2));
    const mismatchMode = await resolveMismatchMode(cli);
    const reportRows = [];
    let exitCode = 0;
    let reportPath = "";
    let mismatchCount = 0;

    prepareTempDownloads();

    try {
        console.log(
            `Mode mismatch: ${mismatchMode === "terminate" ? "terminate (exit 1)" : "warning saja (exit 0)"}`
        );

        const rootFolderId = await resolveCheckRowsFolderId(cli);
        await initDrive();

        console.log("Folder Drive →", rootFolderId);

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

        /** @type {Array<{ id: string, name: string, mimeType: string, relativePath: string, regionalFolder: string }>} */
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

        let expectedCount = null;
        let baselinePath = null;

        for (const file of allFiles) {
            const { localFile } = await downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                `${file.id}_`
            );

            const sheetRows = readSheetRows(localFile);
            const { count, keyColumn } = countRealDataRows(sheetRows);

            const subFolder = file.regionalFolder;
            const fileName = file.name;

            reportRows.push({
                fileName,
                subFolder,
                rowCount: count,
                relativePath: file.relativePath
            });

            const keyHint = keyColumn ? ` [${keyColumn}]` : "";
            console.log(
                `[ok] ${file.relativePath} → ${count} baris${keyHint}`
            );

            if (expectedCount === null) {
                expectedCount = count;
                baselinePath = file.relativePath;
                continue;
            }

            if (count !== expectedCount) {
                mismatchCount += 1;
                const logFn =
                    mismatchMode === "terminate"
                        ? console.error
                        : console.warn;
                logFn(
                    `[count-mismatch] Baseline: ${expectedCount} baris (${baselinePath})`
                );
                logFn(
                    `[count-mismatch]         ${count} baris (${file.relativePath})`
                );
            }
        }

        if (mismatchCount === 0) {
            console.log(
                `\nSemua ${allFiles.length} file cocok: ${expectedCount} baris data.`
            );
        } else {
            console.log(
                `\n${mismatchCount} file beda dari baseline (${expectedCount} baris).`
            );
            exitCode = exitCodeForMismatch(mismatchCount, mismatchMode);
        }
    } catch (err) {
        console.error(err);
        exitCode = 1;
    } finally {
        if (reportRows.length > 0) {
            const annotated = annotateReportRows(
                reportRows,
                reportRows.length > 0 ? reportRows[0].rowCount : null,
                reportRows[0]?.relativePath ?? null
            );
            const stamp = wibTimestampReadable();
            reportPath = path.join(OUTPUT_DIR, `${stamp}.xlsx`);
            await writeReportExcel(annotated, reportPath);
        }
        cleanupTempDownloads();
        if (exitCode !== 0) process.exit(exitCode);
    }
}

module.exports = {
    parseCheckRowsCli,
    resolveMismatchMode,
    exitCodeForMismatch,
    annotateReportRows,
    writeReportExcel,
    shouldSkipFileName
};

if (require.main === module) {
    main().catch(err => {
        console.error(err);
        process.exit(1);
    });
}
