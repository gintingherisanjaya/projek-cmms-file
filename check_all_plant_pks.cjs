/**
 * check_all_plant_pks.cjs
 *
 * Daftar stasiun per file PKS (funcloc 4 segmen) + maintenance plant, QTY func loc,
 * work center, planner group, cost center.
 * Lookup WC/PG mengikuti lsmw-func-loc (tanpa mengubah skrip tersebut).
 *
 *   pnpm run check-all-plant-pks
 *   node check_all_plant_pks.cjs --protect-y-folder-link <url>
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const {
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    parseLsmwCli,
    resolveProtectYSourceFolderId,
    isRegionalFolderName
} = require("./utils/lsmw_cli.cjs");
const {
    buildColumnIndexFirstWins,
    ensureCwcPlannerLookups,
    getEffectiveCostCenter,
    logRegionalMappingForFile,
    resolvePlannerGroup,
    resolveWorkCenter
} = require("./utils/lsmw_lookups.cjs");
const { findDataLayout, normalizeHeader } = require("./utils/lsmw_cell_fill.cjs");
const {
    findCostCenterColumnIndex,
    findFunclocDescAfterColumnIndex
} = require("./utils/equipment_gathering_columns.cjs");

const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const OUTPUT_DIR = path.join("Output", "check-all-plant-pks");
const AGGREGATE_SKIP = "all-pks.xlsx";
const STATION_FUNCLOC_SEGMENT_COUNT = 4;

let drive;

function isEmpty(v) {
    return v === null || v === undefined || String(v).trim() === "";
}

function isValidPlantCode(code) {
    const s = String(code ?? "").trim().toUpperCase();
    return /^[A-Z0-9]{4}$/.test(s);
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

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

/** Nama PKS: "1.PKS AUR GADING.xlsx" → "PKS AUR GADING". */
function extractPksNameFromSourceName(sourceName) {
    let base = path.basename(String(sourceName ?? ""));
    base = base.replace(/\.xlsx$/i, "").trim();
    base = base.replace(/^\d+\.\s*/, "").trim();
    return base || "PKS";
}

function getFuncLocSegmentCount(code) {
    const key = normalizeFuncLocKey(code);
    if (!key) return 0;
    return key.split("-").filter(Boolean).length;
}

function isStationLevelFuncLoc(funcLocKey) {
    return getFuncLocSegmentCount(funcLocKey) === STATION_FUNCLOC_SEGMENT_COUNT;
}

/** Cost center stasiun dari STRNO (sama lsmw-func-loc deriveKostlFromFuncLoc). */
function deriveKostlFromFuncLoc(funcLocKey) {
    const s = normalizeFuncLocKey(funcLocKey).toUpperCase();
    if (!s) return null;
    const parts = s.split("-").filter(Boolean);
    if (parts.length < 4) return null;
    const seg2 = parts[1] ?? "";
    const seg4 = parts[3] ?? "";
    const left = String(seg2)
        .replace(/[^A-Z0-9]/g, "")
        .slice(0, 4)
        .padEnd(4, "0");
    const digits = String(seg4).replace(/\D/g, "");
    const tail = (digits.slice(-2) || "00").padStart(2, "0").slice(-2);
    const kostl = `${left}STAS${tail}`.slice(0, 10);
    return kostl || null;
}

/**
 * Jumlah FUNCTIONAL LOCATION AFTER unik di subtree stasiun (stasiun + turunan).
 */
function countFuncLocUnderStation(dataRows, idxFuncLoc, stationKey) {
    const station = normalizeFuncLocKey(stationKey);
    if (!station) return 0;
    const prefix = `${station}-`;
    const seen = new Set();
    for (const r of dataRows) {
        const key = normalizeFuncLocKey(r[idxFuncLoc]);
        if (!key) continue;
        if (key !== station && !key.startsWith(prefix)) continue;
        seen.add(key);
    }
    return seen.size;
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
 * Maintenance plant per file (prioritas sama skrip LSMW).
 */
function resolveMaintenancePlantPerFile(dataRows, colIndex, idxFuncLoc) {
    function val(r, name) {
        const idx = colIndex[normalizeHeader(name)];
        return idx !== undefined ? r[idx] ?? null : null;
    }

    for (const r of dataRows) {
        const mp = val(r, "MAINTENANCE PLAN AFTER");
        if (!isEmpty(mp)) return String(mp).trim().toUpperCase();
    }

    for (const r of dataRows) {
        const p = val(r, "MAINTENANCE PLANT");
        if (!isEmpty(p) && isValidPlantCode(p)) {
            return String(p).trim().toUpperCase();
        }
    }

    for (const r of dataRows) {
        const p = val(r, "POM");
        if (!isEmpty(p) && isValidPlantCode(p)) {
            return String(p).trim().toUpperCase();
        }
    }

    if (idxFuncLoc !== undefined) {
        for (const r of dataRows) {
            const raw = r[idxFuncLoc];
            if (!raw) continue;
            const segs = String(raw).trim().split("-").filter(Boolean);
            if (segs.length >= 2) {
                const plant = segs[1].trim().toUpperCase();
                if (isValidPlantCode(plant)) return plant;
            }
        }
    }

    return "";
}

/**
 * @param {{ name: string, regionalFolder: string }} fileMeta
 * @param {Array<Array<unknown>>} rawRows
 * @returns {Promise<Array<{ namaPks: string, namaRegional: string, maintenancePlant: string, stasiun: string, workCenter: string, plannerGroup: string, costCenter: string, qtyFuncLoc: number }>>}
 */
async function collectStationRowsForFile(fileMeta, rawRows) {
    const layout = findDataLayout(rawRows);
    const headerRow = rawRows[layout.headerRowIndex];
    if (!headerRow) return [];

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    if (idxFuncLoc === undefined) return [];

    const idxCostCenter = findCostCenterColumnIndex(headerRow);
    const idxFunclocDesc = findFunclocDescAfterColumnIndex(headerRow);
    const dataRows = rawRows.slice(layout.headerRowIndex + 1);

    const funcLocMap = {};
    for (const r of dataRows) {
        const code = normalizeFuncLocKey(r[idxFuncLoc]);
        if (code) funcLocMap[code] = r;
    }

    function val(r, name) {
        const idx = colIndex[normalizeHeader(name)];
        return idx !== undefined ? r[idx] ?? null : null;
    }

    const fileSwerk = resolveMaintenancePlantPerFile(
        dataRows,
        colIndex,
        idxFuncLoc
    );
    if (!isEmpty(fileSwerk)) {
        await logRegionalMappingForFile(fileSwerk, fileMeta.name);
    }

    const namaPks = extractPksNameFromSourceName(fileMeta.name);
    const namaRegional = fileMeta.regionalFolder;
    const out = [];

    for (const r of dataRows) {
        const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);
        if (!funcLocKey || !isStationLevelFuncLoc(funcLocKey)) continue;

        const stasiun =
            idxFunclocDesc !== undefined
                ? String(r[idxFunclocDesc] ?? "").trim()
                : "";

        const effectiveCostCenter = getEffectiveCostCenter(
            r,
            funcLocKey,
            funcLocMap,
            idxCostCenter
        );
        const workCenter = resolveWorkCenter(
            r,
            funcLocKey,
            effectiveCostCenter,
            val,
            fileSwerk
        );
        const plannerGroup = resolvePlannerGroup(
            r,
            funcLocKey,
            effectiveCostCenter,
            val,
            fileSwerk
        );
        const costCenter = deriveKostlFromFuncLoc(funcLocKey) ?? "";
        const qtyFuncLoc = countFuncLocUnderStation(
            dataRows,
            idxFuncLoc,
            funcLocKey
        );

        out.push({
            namaPks,
            namaRegional,
            maintenancePlant: fileSwerk ?? "",
            stasiun,
            workCenter: workCenter ?? "",
            plannerGroup: plannerGroup ?? "",
            costCenter,
            qtyFuncLoc
        });
    }

    return out;
}

function sortReportRows(rows) {
    rows.sort((a, b) => {
        const byRegional = a.namaRegional.localeCompare(b.namaRegional, "id");
        if (byRegional !== 0) return byRegional;
        const byPks = a.namaPks.localeCompare(b.namaPks, "id");
        if (byPks !== 0) return byPks;
        return a.stasiun.localeCompare(b.stasiun, "id");
    });
}

/**
 * @param {Array<{ namaPks: string, namaRegional: string, maintenancePlant: string, stasiun: string, workCenter: string, plannerGroup: string, costCenter: string, qtyFuncLoc: number }>} rows
 * @param {string} outPath
 */
async function writeReportExcel(rows, outPath) {
    fs.mkdirSync(path.dirname(outPath), { recursive: true });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Stasiun per PKS");

    ws.columns = [
        { header: "nama pks", key: "namaPks", width: 32 },
        { header: "nama regional", key: "namaRegional", width: 18 },
        { header: "maintenance plant", key: "maintenancePlant", width: 18 },
        { header: "stasiun", key: "stasiun", width: 36 },
        { header: "work center", key: "workCenter", width: 14 },
        { header: "planner group", key: "plannerGroup", width: 14 },
        { header: "cost center", key: "costCenter", width: 14 },
        { header: "QTY func loc", key: "qtyFuncLoc", width: 12 }
    ];

    for (const row of rows) {
        ws.addRow(row);
    }

    await wb.xlsx.writeFile(outPath);
    console.log("Laporan:", outPath);
}

async function main() {
    const cli = parseLsmwCli(process.argv.slice(2));
    const reportRows = [];
    let exitCode = 0;
    let reportPath = "";

    prepareTempDownloads();

    try {
        const rootFolderId = await resolveProtectYSourceFolderId(cli);
        await initDrive();
        await ensureCwcPlannerLookups();

        console.log("Folder Protect Y →", rootFolderId);

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

        let emptyWc = 0;
        let emptyPg = 0;
        let emptyCc = 0;

        for (const file of allFiles) {
            const { localFile } = await downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                `${file.id}_`
            );

            const rawRows = readSheetRows(localFile);
            const stationRows = await collectStationRowsForFile(file, rawRows);
            reportRows.push(...stationRows);

            for (const row of stationRows) {
                if (!row.workCenter) emptyWc += 1;
                if (!row.plannerGroup) emptyPg += 1;
                if (!row.costCenter) emptyCc += 1;
            }

            console.log(
                `[ok] ${file.relativePath} → ${stationRows.length} stasiun`
            );
        }

        sortReportRows(reportRows);

        console.log(
            `\n${allFiles.length} file; ${reportRows.length} baris stasiun`
        );
        console.log(
            `  kosong — work center: ${emptyWc}, planner group: ${emptyPg}, cost center: ${emptyCc}`
        );
    } catch (err) {
        console.error(err);
        exitCode = 1;
    } finally {
        if (reportRows.length > 0) {
            const stamp = wibTimestampReadable();
            reportPath = path.join(OUTPUT_DIR, `${stamp}.xlsx`);
            await writeReportExcel(reportRows, reportPath);
        } else if (exitCode === 0) {
            console.log("Tidak ada baris stasiun untuk dilaporkan.");
        }
        cleanupTempDownloads();
        if (exitCode !== 0) process.exit(exitCode);
    }
}

module.exports = {
    wibTimestampReadable,
    resolveMaintenancePlantPerFile,
    extractPksNameFromSourceName,
    isStationLevelFuncLoc,
    deriveKostlFromFuncLoc,
    countFuncLocUnderStation,
    collectStationRowsForFile,
    writeReportExcel,
    shouldSkipFileName,
    sortReportRows
};

if (require.main === module) {
    main().catch(err => {
        console.error(err);
        process.exit(1);
    });
}
