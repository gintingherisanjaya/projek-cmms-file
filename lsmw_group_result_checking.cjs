/**
 * lsmw_group_result_checking.cjs
 * Bandingkan EQUIPMENT GROUP AFTER (Protect Y) vs EQART (output LSMW equipment).
 *
 *   pnpm run lsmw-group-result-checking
 *
 * 1. Prompt URL folder Drive berisi output LSMW equipment (*.xlsx)
 * 2. Unduh file LSMW + pasangan Protect Y (SOURCE_ROOT) berdasarkan nama file
 * 3. Per TPLNR / FUNCTIONAL LOCATION AFTER yang sama → bandingkan nilai group
 * 4. Output: Output/lsmw-group-result-checking/{timestamp WIB}/{nama PKS}.xlsx
 */

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { writeGroupCheckExcel } = require("./utils/lsmw_group_check_excel.cjs");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const { input } = require("@inquirer/prompts");
const {
    SOURCE_ROOT_FOLDER_ID,
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    extractFolderIdFromUrl
} = require("./utils/lsmw_cli.cjs");
const { buildColumnIndexFirstWins } = require("./utils/lsmw_lookups.cjs");
const { findDataLayout } = require("./utils/lsmw_cell_fill.cjs");

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OUTPUT_DIR = path.join("Output", "lsmw-group-result-checking");

/** Layout template_equipment.xlsx — sama dengan lsmw_equipment_v0.cjs */
const LSMW_HEADER_ROW = 3;
const LSMW_DATA_START_ROW = 7;

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

function normalizeHeader(name) {
    return String(name ?? "").trim().toUpperCase();
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

function normalizeEquipmentGroup(value) {
    return String(value ?? "").trim().toUpperCase();
}

function funcLocLevel(code) {
    const s = normalizeFuncLocKey(code);
    if (!s) return 0;
    return s.split("-").filter(Boolean).length;
}

function cellText(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) return value.toISOString();
    if (typeof value === "object" && value !== null) {
        if (value.text !== undefined) return String(value.text);
        if (value.result !== undefined) return String(value.result);
        if (value.richText) {
            return value.richText.map(t => t.text ?? "").join("");
        }
    }
    return String(value);
}

function sheetCellText(sheet, row1, col1) {
    const addr = XLSX.utils.encode_cell({ r: row1 - 1, c: col1 - 1 });
    const cell = sheet[addr];
    if (!cell) return "";
    if (cell.w != null && cell.w !== "") return String(cell.w).trim();
    return cellText(cell.v).trim();
}

function fileMatchKey(name) {
    return path
        .basename(String(name ?? ""))
        .trim()
        .toLowerCase()
        .replace(/\.xlsx$/i, "")
        .replace(/\s+/g, " ");
}

function isFunclocDescAfterHeader(header) {
    const h = normalizeHeader(header);
    if (!h) return false;
    if (/\bBEFORE\b/.test(h)) return false;
    if (!h.includes("FUNCTLOC")) return false;
    if (!/\bDESC\b/.test(h)) return false;
    if (!/\b(AFTER|AFTRER)\b/.test(h)) return false;
    return true;
}

function findFunclocDescAfterColumnIndex(headerRow) {
    const indices = [];
    for (let i = 0; i < headerRow.length; i++) {
        if (isFunclocDescAfterHeader(headerRow[i])) indices.push(i);
    }
    if (indices.length === 0) return undefined;
    if (indices.length === 1) return indices[0];
    const withLevel = indices.filter(i =>
        /\bLEVEL\b/.test(normalizeHeader(headerRow[i]))
    );
    return withLevel.length >= 1 ? withLevel[0] : indices[0];
}

/** Sama dengan filter baris equipment di lsmw_equipment_v0.cjs */
function isEquipmentRow(funcLoc, desc) {
    if (funcLocLevel(funcLoc) < 5) return false;
    const d = String(desc ?? "");
    if (/drive\s*unit/i.test(d)) return false;
    if (/accessories\s+weigh\s*bridge/i.test(d)) return false;
    return true;
}

/** Baris 3 template: nama kolom → nomor kolom Excel (1-based). */
function buildLsmwColumnMap(sheet) {
    const range = sheet["!ref"]
        ? XLSX.utils.decode_range(sheet["!ref"])
        : { e: { c: 25, r: LSMW_HEADER_ROW } };

    const map = {};
    for (let c = 1; c <= range.e.c + 1; c += 1) {
        const name = sheetCellText(sheet, LSMW_HEADER_ROW, c);
        if (!name || map[name] !== undefined) continue;
        map[name] = c;
    }
    return map;
}

function readProtectYRows(localPath) {
    const workbook = XLSX.readFile(localPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const layout = findDataLayout(rows);
    const headerRow = rows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header Protect Y tidak ditemukan" };
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxEquipGroup = colIndex["EQUIPMENT GROUP AFTER"];
    const idxDesc = findFunclocDescAfterColumnIndex(headerRow);

    if (idxFuncLoc === undefined) {
        return { ok: false, reason: "Kolom FUNCTIONAL LOCATION AFTER tidak ada" };
    }
    if (idxEquipGroup === undefined) {
        return { ok: false, reason: "Kolom EQUIPMENT GROUP AFTER tidak ada" };
    }

    const byFuncLoc = new Map();
    const dataRows = rows.slice(layout.headerRowIndex + 1);

    for (const r of dataRows) {
        const funcLoc = normalizeFuncLocKey(r[idxFuncLoc]);
        if (!funcLoc) continue;

        const desc =
            idxDesc !== undefined ? String(r[idxDesc] ?? "").trim() : "";
        if (!isEquipmentRow(funcLoc, desc)) continue;

        const eg = String(r[idxEquipGroup] ?? "").trim();
        if (!byFuncLoc.has(funcLoc)) {
            byFuncLoc.set(funcLoc, { funcDesc: desc, equipmentGroupSource: eg });
        }
    }

    return { ok: true, byFuncLoc };
}

/**
 * Baca output LSMW: header baris 3 (EQART kol 5, TPLNR kol 21), data dari baris 7.
 */
function readLsmwRows(localPath) {
    const workbook = XLSX.readFile(localPath, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) return { ok: false, reason: "Sheet LSMW kosong" };

    const colMap = buildLsmwColumnMap(sheet);
    const colEqart = colMap.EQART;
    const colTplnr = colMap.TPLNR;
    const colShtxt = colMap["SHTXT*"];

    if (!colEqart || !colTplnr) {
        const found = Object.keys(colMap).sort().join(", ");
        return {
            ok: false,
            reason:
                `Bukan file output LSMW equipment (baris ${LSMW_HEADER_ROW}): ` +
                `EQART=${colEqart ?? "-"}, TPLNR=${colTplnr ?? "-"}. ` +
                `Kolom: ${found || "(kosong)"}`
        };
    }

    const range = sheet["!ref"]
        ? XLSX.utils.decode_range(sheet["!ref"])
        : { e: { r: LSMW_DATA_START_ROW, c: colTplnr - 1 } };

    const byFuncLoc = new Map();

    for (let r = LSMW_DATA_START_ROW; r <= range.e.r + 1; r += 1) {
        const funcLoc = normalizeFuncLocKey(
            sheetCellText(sheet, r, colTplnr)
        );
        const eqart = sheetCellText(sheet, r, colEqart);
        const funcDesc = colShtxt ? sheetCellText(sheet, r, colShtxt) : "";

        if (!funcLoc && !eqart && !funcDesc) {
            if (byFuncLoc.size > 0) break;
            continue;
        }

        if (!funcLoc) continue;

        if (!byFuncLoc.has(funcLoc)) {
            byFuncLoc.set(funcLoc, {
                funcDesc,
                equipmentGroupLsmw: eqart
            });
        }
    }

    if (byFuncLoc.size === 0) {
        return {
            ok: false,
            reason: `Tidak ada data LSMW (baris ${LSMW_DATA_START_ROW}+, EQART kol ${colEqart}, TPLNR kol ${colTplnr})`
        };
    }

    return {
        ok: true,
        byFuncLoc,
        layoutSource: `baris ${LSMW_HEADER_ROW}: EQART→${colEqart}, TPLNR→${colTplnr}`
    };
}

/** Semua baris LSMW + hitung yang berbeda (untuk log terminal). */
function compareMaps(protectYMap, lsmwMap) {
    const items = [];
    let differenceCount = 0;

    for (const [funcLoc, lsmwRow] of lsmwMap) {
        const src = protectYMap.get(funcLoc);
        const egSource = src?.equipmentGroupSource ?? "";
        const egLsmw = lsmwRow.equipmentGroupLsmw ?? "";

        if (
            normalizeEquipmentGroup(egSource) !==
            normalizeEquipmentGroup(egLsmw)
        ) {
            differenceCount += 1;
        }

        items.push({
            funcLoc,
            funcDesc: lsmwRow.funcDesc || src?.funcDesc || "",
            equipmentGroupSource: egSource,
            equipmentGroupLsmw: egLsmw
        });
    }

    return { items, differenceCount };
}

function pksLabelFromFileName(fileName) {
    return path.basename(String(fileName ?? "")).replace(/\.xlsx$/i, "").trim();
}

function pksOutputBaseName(sourceFileName) {
    const base = pksLabelFromFileName(sourceFileName);
    return base.replace(/[/\\?%*:|"<>]/g, "_");
}

function pksXlsxFileName(sourceFileName) {
    return `${pksOutputBaseName(sourceFileName)}.xlsx`;
}

/** Tulis Excel hasil satu PKS segera setelah selesai diproses. */
async function writePksExcel(runDir, sourceFileName, meta) {
    const xlsxName = pksXlsxFileName(sourceFileName);
    const xlsxPath = path.join(runDir, xlsxName);

    await writeGroupCheckExcel(xlsxPath, meta.items ?? [], {
        status: meta.status,
        reason: meta.reason,
        pks: meta.pks
    });

    return { xlsxPath, xlsxName };
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

async function listSpreadsheetsRecursive(folderId, out = []) {
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

        for (const file of res.data.files ?? []) {
            if (file.mimeType === "application/vnd.google-apps.folder") {
                await listSpreadsheetsRecursive(file.id, out);
            } else if (isProcessableSpreadsheet(file)) {
                out.push(file);
            }
        }
        pageToken = res.data.nextPageToken;
    } while (pageToken);

    return out;
}

async function buildProtectYFileIndex() {
    const index = new Map();

    async function walk(folderId) {
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

            for (const file of res.data.files ?? []) {
                if (file.mimeType === "application/vnd.google-apps.folder") {
                    await walk(file.id);
                } else if (isProcessableSpreadsheet(file)) {
                    const key = fileMatchKey(file.name);
                    if (!index.has(key)) index.set(key, file);
                }
            }
            pageToken = res.data.nextPageToken;
        } while (pageToken);
    }

    await walk(SOURCE_ROOT_FOLDER_ID);
    return index;
}

async function main() {
    const stamp = wibTimestampForFilename();
    const runDir = path.join(OUTPUT_DIR, stamp);

    prepareTempDownloads();

    try {
        await initDrive();

        const lsmwFolderUrl = await input({
            message:
                "URL folder Google Drive output LSMW Equipment (REGIONAL/…/*.xlsx):",
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

        const lsmwFolderId = extractFolderIdFromUrl(String(lsmwFolderUrl).trim());

        console.log("Memuat indeks Protect Y...");
        const protectYIndex = await buildProtectYFileIndex();
        console.log(`  ${protectYIndex.size} file`);

        console.log("Memuat daftar file LSMW...");
        const lsmwFiles = await listSpreadsheetsRecursive(lsmwFolderId);
        console.log(`  ${lsmwFiles.length} spreadsheet`);

        fs.mkdirSync(runDir, { recursive: true });

        let filesCompared = 0;
        let filesSkipped = 0;
        let totalRows = 0;
        let totalDifferences = 0;

        for (const lsmwFile of lsmwFiles) {
            const pks = pksLabelFromFileName(lsmwFile.name);
            console.log("Memproses:", lsmwFile.name);

            const finishPks = async (status, extra = {}) => {
                const { xlsxName } = await writePksExcel(runDir, lsmwFile.name, {
                    pks,
                    status,
                    reason: extra.reason,
                    items: extra.items ?? []
                });

                if (status === "skipped") {
                    console.log(
                        `  [skipped] ${extra.reason ?? "unknown"} → ${xlsxName}`
                    );
                } else {
                    console.log(
                        `  [compared] ${extra.rowCount ?? 0} baris (${extra.differenceCount ?? 0} berbeda) → ${xlsxName}`
                    );
                }
            };

            const protectYFile = protectYIndex.get(fileMatchKey(lsmwFile.name));

            if (!protectYFile) {
                filesSkipped += 1;
                await finishPks("skipped", {
                    reason: "Protect Y tidak ditemukan",
                    protectYFile: null
                });
                continue;
            }

            const { localFile: lsmwLocal } = await downloadSpreadsheetToTemp(
                drive,
                lsmwFile,
                TEMP_DOWNLOAD_PATH,
                "lsmw_"
            );
            const { localFile: protectYLocal } = await downloadSpreadsheetToTemp(
                drive,
                protectYFile,
                TEMP_DOWNLOAD_PATH,
                "py_"
            );

            const protectY = readProtectYRows(protectYLocal);
            if (!protectY.ok) {
                filesSkipped += 1;
                await finishPks("skipped", {
                    reason: protectY.reason,
                    protectYFile: protectYFile.name
                });
                continue;
            }

            const lsmw = readLsmwRows(lsmwLocal);
            if (!lsmw.ok) {
                filesSkipped += 1;
                await finishPks("skipped", {
                    reason: lsmw.reason,
                    protectYFile: protectYFile.name
                });
                continue;
            }

            const { items, differenceCount } = compareMaps(
                protectY.byFuncLoc,
                lsmw.byFuncLoc
            );
            filesCompared += 1;
            totalRows += items.length;
            totalDifferences += differenceCount;

            await finishPks("compared", {
                rowCount: items.length,
                differenceCount,
                items
            });
        }

        console.log("\nSelesai.");
        console.log(`  Dibandingkan: ${filesCompared}, dilewati: ${filesSkipped}`);
        console.log(`  Total baris: ${totalRows}, berbeda: ${totalDifferences}`);
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
} else {
    module.exports = {
        readProtectYRows,
        readLsmwRows,
        compareMaps,
        buildLsmwColumnMap,
        fileMatchKey,
        LSMW_HEADER_ROW,
        LSMW_DATA_START_ROW
    };
}
