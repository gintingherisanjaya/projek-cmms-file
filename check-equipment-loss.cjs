/**
 * check-equipment-loss.cjs
 * Bandingkan EQUIPMENT NUMBER + EQKTU BEFORE: LPP Standart vs Protect Y.
 *
 *   pnpm run check-equipment-loss
 *
 * Sumber LPP: folder Drive LPP Standart (REGIONAL*)
 * Referensi: folder Protect Y (prompt URL, subfolder REGIONAL*), template sama
 *
 * Status:
 *   EXIST   — ada di LPP dan Protect Y (COST CENTER BEFORE mengandung STAS)
 *   MISSING — ada di LPP, tidak di Protect Y
 *   ANOMALY — tidak di LPP, ada di Protect Y
 *   UNUSED  — COST CENTER BEFORE tidak mengandung STAS
 *
 * Output: Output/check-equipment-loss/{timestamp WIB}/
 *   REGIONAL N/{nama PKS}.xlsx
 *   semua-pks.xlsx — gabungan semua baris
 */

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const { input } = require("@inquirer/prompts");
const {
    LPP_STANDARD_ROOT_FOLDER_ID,
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    extractFolderIdFromUrl,
    isRegionalFolderName
} = require("./utils/lsmw_cli.cjs");
const { buildColumnIndexFirstWins } = require("./utils/lsmw_lookups.cjs");
const { findDataLayout, normalizeHeader } = require("./utils/lsmw_cell_fill.cjs");
const {
    writeEquipmentLossExcel
} = require("./utils/check_equipment_loss_excel.cjs");
const {
    STATUS,
    compareEquipmentMaps
} = require("./utils/check_equipment_loss.cjs");
const { normalizeCostCenter } = require("./utils/equipment_gathering_columns.cjs");

const OUTPUT_DIR = path.join("Output", "check-equipment-loss");
const SUMMARY_FILE_NAME = "semua-pks.xlsx";
const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = ["https://www.googleapis.com/auth/drive"];

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

function normalizeEquipmentNumber(value) {
    if (value === null || value === undefined || value === "") return "";
    return String(value).trim();
}

function fileMatchKey(name) {
    return path
        .basename(String(name ?? ""))
        .trim()
        .toLowerCase()
        .replace(/\.xlsx$/i, "")
        .replace(/\s+/g, " ");
}

function pksLabelFromFileName(fileName) {
    return path.basename(String(fileName ?? "")).replace(/\.xlsx$/i, "").trim();
}

function pksXlsxFileName(sourceFileName) {
    const base = pksLabelFromFileName(sourceFileName);
    return `${base.replace(/[/\\?%*:|"<>]/g, "_")}.xlsx`;
}

function readEquipmentRows(localPath) {
    const workbook = XLSX.readFile(localPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) {
        return { ok: false, reason: "Sheet kosong" };
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const layout = findDataLayout(rows);
    const headerRow = rows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header tidak ditemukan" };
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxEquip = colIndex["EQUIPMENT NUMBER"];
    const idxEqktu = colIndex["EQKTU BEFORE"];
    const idxCcBefore = colIndex["COST CENTER BEFORE"];

    if (idxEquip === undefined) {
        return { ok: false, reason: "Kolom EQUIPMENT NUMBER tidak ada" };
    }
    if (idxCcBefore === undefined) {
        return { ok: false, reason: "Kolom COST CENTER BEFORE tidak ada" };
    }

    const byEquipNum = new Map();
    const dataRows = rows.slice(layout.headerRowIndex + 1);

    for (const r of dataRows) {
        const equipNum = normalizeEquipmentNumber(r[idxEquip]);
        if (!equipNum) continue;

        const eqktu =
            idxEqktu !== undefined ? String(r[idxEqktu] ?? "").trim() : "";
        const costCenterBefore = normalizeCostCenter(r[idxCcBefore]);

        if (!byEquipNum.has(equipNum)) {
            byEquipNum.set(equipNum, { eqktuBefore: eqktu, costCenterBefore });
        }
    }

    return { ok: true, byEquipNum };
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

async function buildProtectYFileIndex(rootFolderId) {
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

async function promptProtectYFolderUrl() {
    return input({
        message:
            "URL folder Google Drive — Protect Y yang akan dicek (subfolder REGIONAL*):",
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

/**
 * @returns {Array<{ regional: string, file: object }>}
 */
async function listLppFilesByRegional() {
    const entries = [];
    const rootChildren = await listFolderChildren(LPP_STANDARD_ROOT_FOLDER_ID);

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
    const runDir = path.join(OUTPUT_DIR, stamp);

    prepareTempDownloads();

    try {
        await initDrive();

        const protectYFolderUrl = await promptProtectYFolderUrl();
        const protectYRootFolderId = extractFolderIdFromUrl(
            String(protectYFolderUrl).trim()
        );

        console.log("Memuat indeks Protect Y...");
        const protectYIndex = await buildProtectYFileIndex(protectYRootFolderId);
        console.log(`  ${protectYIndex.size} file`);

        console.log("Memuat file LPP Standart (folder REGIONAL*)...");
        const lppEntries = await listLppFilesByRegional();
        console.log(`  ${lppEntries.length} spreadsheet`);

        fs.mkdirSync(runDir, { recursive: true });

        const allRows = [];
        let filesCompared = 0;
        let filesSkipped = 0;
        const counts = { EXIST: 0, MISSING: 0, ANOMALY: 0, UNUSED: 0 };

        for (const { regional, file: lppFile } of lppEntries) {
            const namaPks = pksLabelFromFileName(lppFile.name);
            const regionalDir = path.join(runDir, regional);
            fs.mkdirSync(regionalDir, { recursive: true });

            const outPath = path.join(regionalDir, pksXlsxFileName(lppFile.name));
            console.log(`[${regional}] ${lppFile.name}`);

            const protectYFile = protectYIndex.get(fileMatchKey(lppFile.name));

            if (!protectYFile) {
                filesSkipped += 1;
                console.log("  [skipped] Protect Y tidak ditemukan");
                await writeEquipmentLossExcel(outPath, [], {
                    reason: "Protect Y tidak ditemukan"
                });
                continue;
            }

            const { localFile: lppLocal } = await downloadSpreadsheetToTemp(
                drive,
                lppFile,
                TEMP_DOWNLOAD_PATH,
                "lpp_"
            );
            const { localFile: protectYLocal } = await downloadSpreadsheetToTemp(
                drive,
                protectYFile,
                TEMP_DOWNLOAD_PATH,
                "py_"
            );

            const lpp = readEquipmentRows(lppLocal);
            if (!lpp.ok) {
                filesSkipped += 1;
                console.log(`  [skipped] LPP: ${lpp.reason}`);
                await writeEquipmentLossExcel(outPath, [], { reason: lpp.reason });
                continue;
            }

            const protectY = readEquipmentRows(protectYLocal);
            if (!protectY.ok) {
                filesSkipped += 1;
                console.log(`  [skipped] Protect Y: ${protectY.reason}`);
                await writeEquipmentLossExcel(outPath, [], {
                    reason: protectY.reason
                });
                continue;
            }

            const items = compareEquipmentMaps(
                lpp.byEquipNum,
                protectY.byEquipNum,
                namaPks
            );

            for (const row of items) {
                counts[row.status] = (counts[row.status] ?? 0) + 1;
            }

            allRows.push(...items);
            filesCompared += 1;

            await writeEquipmentLossExcel(outPath, items);

            console.log(
                `  [ok] EXIST ${items.filter(r => r.status === STATUS.EXIST).length}, ` +
                    `MISSING ${items.filter(r => r.status === STATUS.MISSING).length}, ` +
                    `ANOMALY ${items.filter(r => r.status === STATUS.ANOMALY).length}, ` +
                    `UNUSED ${items.filter(r => r.status === STATUS.UNUSED).length}`
            );
        }

        const summaryPath = path.join(runDir, SUMMARY_FILE_NAME);
        await writeEquipmentLossExcel(summaryPath, allRows);

        console.log("\nSelesai.");
        console.log(`  Dibandingkan: ${filesCompared}, dilewati: ${filesSkipped}`);
        console.log(`  EXIST: ${counts.EXIST ?? 0}`);
        console.log(`  MISSING: ${counts.MISSING ?? 0}`);
        console.log(`  ANOMALY: ${counts.ANOMALY ?? 0}`);
        console.log(`  UNUSED: ${counts.UNUSED ?? 0}`);
        console.log(`  Total baris: ${allRows.length}`);
        console.log(`  Output: ${runDir}`);
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
