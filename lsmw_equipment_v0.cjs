/**
 * lsmw_equipment_v0.cjs
 * ─────────────────────────────────────────────────────────────────────────────
 * Pipeline konversi data equipment PKS → format upload LSMW/SAP (template Excel).
 *
 * Alur utama
 *   1. Autentikasi Google Drive (oauth.json + token.json, auto-refresh).
 *   2. Baca folder REGIONAL 1–7 di folder Protect Y (prompt / --protect-y-folder-link).
 *   3. Unduh setiap .xlsx / Google Sheets sumber (Output/temp_downloads).
 *   4. Transformasi → isi template_equipment.xlsx (baris mulai 7).
 *   5. Simpan ke Output/1. LSMW Create Equipment V.O/… (struktur folder mengikuti sumber).
 *   6. Unggah ke Drive hanya jika ada flag --output-drive-folder-link <url>.
 *   7. Tulis column.json (audit kolom).
 *   8. Jika ada baris tanpa mapping EQART: hentikan run, laporkan FUNCTLOC DESC. AFTER, tulis invalid-group-mapping JSON.
 *
 * Sumber: folder Protect Y (prompt URL saat dijalankan, atau --protect-y-folder-link).
 *
 *   node lsmw_equipment_v0.cjs
 *   node lsmw_equipment_v0.cjs --protect-y-folder-link https://drive.google.com/drive/folders/...
 *   node lsmw_equipment_v0.cjs --output-drive-folder-link https://drive.google.com/drive/folders/...
 *
 * Struktur file sumber
 *   - Sheet pertama; header di baris ke-2; data dari baris ke-3.
 *   - Kolom wajib: FUNCTIONAL LOCATION AFTER.
 *
 * Deteksi kolom FUNCTLOC DESC (dinamis antar PKS)
 *   Header dipilih jika mengandung ketiga komponen: FUNCTLOC, DESC, AFTER.
 *   - FUNCTLOC: substring (termasuk typo mis. EFUNCTLOC).
 *   - DESC / AFTER: kata utuh (regex word boundary); AFTER juga menerima typo AFTRER.
 *   - Varian BEFORE (mis. FUNCTLOC DESC. BEFORE) diabaikan agar tidak salah ambil.
 *   - Jika beberapa kolom cocok, prioritas yang mengandung kata LEVEL.
 *
 * Filter baris equipment
 *   - Functional location level ≥ 5 (jumlah segmen dipisah "-").
 *   - Dikecualikan jika deskripsi mengandung:
 *       • drive unit
 *       • accessories weigh bridge / accessories weighbridge
 *   - FUNCTIONAL LOCATION AFTER = primary key lintas file sumber; duplikat tidak
 *     masuk output equipment (tanpa log/arsip baris yang dilewati).
 *   - Baris dengan background merah pada kolom EQUIPMENT GROUP AFTER di-skip
 *     (Google Sheets API + deteksi fill di .xlsx).
 *
 * Isian nilai
 *   - CAPACITY: dari baris sendiri; jika kosong, naik ke parent.
 *     CAPACITY: teks dalam kurung dihapus (mis. "19 Pintu (@ 15 Ton. TBS)" → "19 Pintu").
 *     funcloc (potong suffix kode per "-" sampai ketemu nilai).
 *   - BAUJJ: STANDART UMUR TEKNIS (TAHUN) jika 4 digit (baris/parent), else CONSTRUCTION YEAR
 *     (baris/parent); kosong boleh.
 *   - CAPACITY & BAUJJ di output: hanya ditulis jika berubah dari baris
 *     sebelumnya (selain itu null).
 *   - Konstanta output: TEK, PALM, ZPM_PTPN.
 *   - EQART (EQUIPMENT GROUP AFTER): hanya dari equipment-group-mapping.json
 *     (equipmentGroupFromDesc.js): childrens → child_follow (jika parent punya EG) → parents
 *     → child_follow lagi. Frasa child_follow_parent_group tidak dipakai di childrens/parents.
 *     Tanpa mapping → proses dihentikan segera (fail-fast); laporan terminal + JSON invalid-group-mapping.
 *   - ABC INDICATORS: sumber → utils/abc.csv (pola funcloc) → default 2.
 *   - PLANNER GROUP / WORK CENTER: utils/planner-group.csv & utils/cwc.csv (menggantikan sumber);
 *     kolom regional CSV dari maintenance plant per file (SWERK/BEGRU).
 *   - BEGRU*, SWERK*, IWERK* & WERGW: nilai per file (MAINTENANCE PLAN AFTER pertama;
 *     fallback MAINTENANCE PLANT / POM / segmen funcloc) — sama di semua baris output.
 *
 * FUNCTLOC DESC AFTER → SHTXT*: alias dari mapping-alias-desc-func-loc.xlsx jika cocok.
 *
 * File: oauth.json, token.json, template_equipment.xlsx, mapping-alias-desc-func-loc.xlsx
 *
 * Output lokal: Output/1. LSMW Create Equipment V.O/***.xlsx (nama sama dengan file sumber)
 * Output Drive:  hanya dengan --output-drive-folder-link (folder tujuan + struktur REGIONAL)
 * Saat gagal mapping: Output/invalid-group-mapping/{timestamp WIB}.json (satu baris) + exit 1
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
    isRegionalFolderName,
    lsmwOutputFileName
} = require("./utils/lsmw_cli.cjs");
const {
    isEmpty,
    ensureCwcPlannerLookups,
    shouldSkipLevel5Stas13Row,
    getEffectiveCostCenter,
    logRegionalMappingForFile,
    buildColumnIndexFirstWins,
    resolvePlannerGroup,
    resolveWorkCenter
} = require("./utils/lsmw_lookups.cjs");
const { ensureAbcRules, resolveAbcIndicator } = require("./utils/lsmw_abc.cjs");
const {
    collectRedEquipGroupRowIndices
} = require("./utils/lsmw_equip_group_red.cjs");
const { findDataLayout } = require("./utils/lsmw_cell_fill.cjs");
const {
    ensureFunclocDescAliasMap,
    fixFunclocDescSourceTypos,
    applyFunclocDescAlias
} = require("./utils/funcloc_desc_alias.cjs");
const { normalizeCapacityValue } = require("./utils/lsmw_capacity.cjs");
const { findCostCenterColumnIndex } = require("./utils/equipment_gathering_columns.cjs");
const {
    auditProtectYColumns,
    resolutionsToValidationMap,
    resolveEquipmentGroupAfterColumn
} = require("./utils/lsmw_equipment_column_audit.cjs");
const {
    writeEquipmentColumnValidationExcel
} = require("./utils/lsmw_equipment_validation_excel.cjs");
const { resolveBaujjYear } = require("./utils/lsmw_baujj.cjs");
const {
    ALL_PKS_FILE_NAME,
    regionalAggregateFileName,
    writeRegionalAndAllPksAggregates
} = require("./utils/lsmw_pks_aggregates.cjs");
const {
    withRetry,
    isRetryableDriveError,
    sleep
} = require("./utils/lsmw_retry.cjs");
const { loadEquipmentNumberMapping } = require("./utils/equipment_number_mapping.cjs");
const {
    lookupEquipmentNumber,
    buildMaintenanceItemFallbackEqunr
} = require("./utils/lsmw_maintenance_item_transform.cjs");
const { getFuncLocSecondSegment } = require("./utils/lsmw_tasklist_transform.cjs");
const { writeDeleteWorkbook } = require("./utils/lsmw_delete_runner.cjs");

const DRIVE_RETRY = {
    maxAttempts: 6,
    delayMs: 5000,
    retryIf: isRetryableDriveError
};

const SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly"
];

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";

const LOCAL_OUTPUT_ROOT = path.join("Output", "1. LSMW Create Equipment V.O");
const INVALID_GROUP_MAPPING_DIR = path.join(
    "Output",
    "invalid-group-mapping"
);
const TEMPLATE_PATH = "./template_equipment.xlsx";
const DELETE_TEMPLATE_PATH = "./template_delete_equipment.xlsx";
const COLUMN_JSON_PATH = path.join(LOCAL_OUTPUT_ROOT, "column.json");
const VALIDATION_XLSX_PATH = path.join(LOCAL_OUTPUT_ROOT, "validation.xlsx");
const TO_DELETE_ROOT = path.join(LOCAL_OUTPUT_ROOT, "to-delete");
const EQUIPMENT_OUTPUT_START_ROW = 7;

let drive;
let sheets;
let outputDriveFolderId = null;

/** Baris ke-2 per file (index 1) dikumpulkan untuk column.json */
let secondRowDump = [];

/** Satu baris per file PKS untuk validation.xlsx */
let columnValidationRows = [];

/** Primary key lintas file: FUNCTIONAL LOCATION AFTER yang sudah masuk output equipment. */
const usedFuncLocAfter = new Set();

/** Total baris di-skip karena EQUIPMENT GROUP AFTER merah (semua file). */
let totalRedEquipGroupSkipped = 0;

/** Planning Plant + Description → Equipment (new-equipment-number.xlsx). */
/** @type {Map<string, string> | null} */
let equipmentByPlantDesc = null;

/** EQUNR to-delete yang sudah masuk output lintas file. */
const usedToDeleteEqunrGlobal = new Set();

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

/**
 * Hentikan run pada baris equipment pertama tanpa EQART dari equipment-group-mapping.json.
 * @param {{ sourceFile: string, funcLoc: unknown, funcDesc: string, equipmentGroupResource: unknown }} entry
 */
function abortOnMissingEquipmentGroup(entry) {
    const stamp = wibTimestampForFilename();
    const reportPath = path.join(INVALID_GROUP_MAPPING_DIR, `${stamp}.json`);

    fs.mkdirSync(INVALID_GROUP_MAPPING_DIR, { recursive: true });
    fs.writeFileSync(
        reportPath,
        JSON.stringify(
            {
                generatedAtWib: stamp,
                mappingFile: path.basename(EQUIPMENT_GROUP_MAPPING_PATH),
                totalInvalid: 1,
                items: [entry]
            },
            null,
            2
        ),
        "utf8"
    );

    const funcDesc = String(entry.funcDesc ?? "").trim() || "(kosong)";
    const funcLoc = String(entry.funcLoc ?? "").trim() || "(kosong)";

    console.error("\n  [error] Mapping equipment group gagal");
    console.error(`  FUNCTLOC DESC. AFTER: ${funcDesc}`);
    console.error(`  FUNCTIONAL LOCATION AFTER: ${funcLoc}`);
    console.error(`  File: ${entry.sourceFile}`);
    console.error(`  Laporan: ${path.resolve(reportPath)}`);
    console.error("  Proses dihentikan.\n");
    process.exit(1);
}

const EQUIPMENT_GROUP_MAPPING_PATH = path.join(
    __dirname,
    "equipment-group-mapping.json"
);

let equipmentGroupMappingCache = null;
let equipmentGroupFromDescMod = null;

async function ensureEquipmentGroupMapping() {
    if (equipmentGroupMappingCache && equipmentGroupFromDescMod) {
        return {
            mod: equipmentGroupFromDescMod,
            mapping: equipmentGroupMappingCache
        };
    }
    equipmentGroupFromDescMod = await import("./equipmentGroupFromDesc.js");
    equipmentGroupMappingCache =
        equipmentGroupFromDescMod.loadEquipmentGroupMapping(
            EQUIPMENT_GROUP_MAPPING_PATH
        );
    return {
        mod: equipmentGroupFromDescMod,
        mapping: equipmentGroupMappingCache
    };
}

function cellToJson(v) {
    if (v === undefined || v === null || v === "") return null;
    if (v instanceof Date) return v.toISOString();
    return v;
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
            console.log("Token refreshed.");
        }
    } else {

        const newAuth = await authenticate({
            keyfilePath: OAUTH_PATH,
            scopes: SCOPES
        });

        const token = newAuth.credentials;
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
        auth.setCredentials(token);
        console.log("Token saved to token.json");
    }

    drive = google.drive({
        version: "v3",
        auth
    });
    sheets = google.sheets({ version: "v4", auth });
}

async function createFolderIfNotExists(name, parentId) {

    const res = await drive.files.list({
        q: `name='${name}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: "files(id,name)"
    });

    if (res.data.files.length > 0) {
        return res.data.files[0].id;
    }

    const folder = await drive.files.create({
        resource: {
            name,
            mimeType: "application/vnd.google-apps.folder",
            parents: [parentId]
        },
        fields: "id"
    });

    return folder.data.id;
}

async function listFolderChildrenWithRetry(folderId) {
    return withRetry(
        async () => {
            const res = await drive.files.list({
                q: `'${folderId}' in parents and trashed=false`,
                fields: "files(id,name,mimeType)",
                pageSize: 1000
            });
            return res.data.files ?? [];
        },
        { ...DRIVE_RETRY, label: `list folder ${folderId}` }
    );
}

async function downloadSpreadsheetWithRetry(file) {
    return withRetry(
        () => downloadSpreadsheetToTemp(drive, file, TEMP_DOWNLOAD_PATH),
        { ...DRIVE_RETRY, label: `unduh ${file.name}` }
    );
}

async function uploadFile(filePath, fileName, parentId) {

    const existing = await drive.files.list({
        q: `name='${fileName}' and '${parentId}' in parents and trashed=false`,
        fields: "files(id,name)"
    });

    if (existing.data.files.length > 0) {

        for (const f of existing.data.files) {

            await drive.files.delete({
                fileId: f.id
            });

            console.log("Deleted old file:", f.name);
        }
    }

    await drive.files.create({
        resource: {
            name: fileName,
            parents: [parentId]
        },
        media: {
            mimeType:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            body: fs.createReadStream(filePath)
        },
        fields: "id"
    });

    console.log("Uploaded:", fileName);
}

function normalizeHeader(name) {
    if (!name) return "";
    return name.toString().trim().toUpperCase();
}

function isValidPlantCode(code) {
    const s = String(code ?? "").trim().toUpperCase();
    return /^[A-Z0-9]{4}$/.test(s);
}

/** Nilai BEGRU* / plant per file: satu nilai untuk semua baris, tidak kosong jika sumber ada. */
function resolveFileBegru(dataRows, equipmentRows, val, getValueWithParent, idxFuncLoc) {
    for (const r of dataRows) {
        const mp = val(r, "MAINTENANCE PLAN AFTER");
        if (!isEmpty(mp)) return mp;
    }

    for (const r of equipmentRows) {
        const mp = getValueWithParent(r, "MAINTENANCE PLAN AFTER");
        if (!isEmpty(mp)) return mp;
    }

    const scanPlant = rows => {
        for (const r of rows) {
            const p = val(r, "MAINTENANCE PLANT");
            if (!isEmpty(p) && isValidPlantCode(p)) {
                return String(p).trim().toUpperCase();
            }
        }
        for (const r of rows) {
            const p = val(r, "POM");
            if (!isEmpty(p) && isValidPlantCode(p)) {
                return String(p).trim().toUpperCase();
            }
        }
        if (idxFuncLoc !== undefined) {
            for (const r of rows) {
                const raw = r[idxFuncLoc];
                if (!raw) continue;
                const segs = String(raw).trim().split("-").filter(Boolean);
                if (segs.length >= 2) {
                    const plant = segs[1].trim().toUpperCase();
                    if (isValidPlantCode(plant)) return plant;
                }
            }
        }
        return null;
    };

    return scanPlant(dataRows) || scanPlant(equipmentRows);
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

/** Header kolom deskripsi funcloc "after" — wajib FUNCTLOC + DESC + AFTER; bukan varian BEFORE. */
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
    if (withLevel.length >= 1) return withLevel[0];
    return indices[0];
}

async function writeEquipmentOutputRows(outputRowArrays, outPath) {
    const workbookTemplate = new ExcelJS.Workbook();
    await workbookTemplate.xlsx.readFile(TEMPLATE_PATH);

    const sheetTemplate = workbookTemplate.worksheets[0];
    let startRow = EQUIPMENT_OUTPUT_START_ROW;

    for (const rowData of outputRowArrays) {
        const row = sheetTemplate.getRow(startRow++);
        rowData.forEach((v, i) => {
            row.getCell(i + 1).value = v;
        });
        row.commit();
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await workbookTemplate.xlsx.writeFile(outPath);
}

/**
 * @param {Array<{ primary: string, oldName: string }>} rows
 * @param {string} outPath
 */
async function writeToDeleteOutputRows(rows, outPath) {
    await writeDeleteWorkbook(rows, DELETE_TEMPLATE_PATH, outPath, {
        oldNameLookup: true
    });
}

/**
 * @param {Array<{ r: Array<unknown>, rowIndex: number }>} candidates
 * @param {{
 *   sourceName: string,
 *   dataStartExcelRow: number,
 *   idxFuncLoc: number,
 *   idxCostCenter: number | undefined,
 *   funcLocMap: Record<string, Array<unknown>>,
 *   funclocDescRaw: (r: Array<unknown>) => string,
 *   funclocDescText: (r: Array<unknown>) => string
 * }} ctx
 * @returns {Array<{ primary: string, oldName: string }>}
 */
function buildToDeleteRows(candidates, ctx) {
    const seenInFile = new Set();
    /** @type {Array<{ primary: string, oldName: string }>} */
    const output = [];

    for (const { r, rowIndex } of candidates) {
        const funcLocKey = normalizeFuncLocKey(r[ctx.idxFuncLoc]);
        const effectiveCostCenter = getEffectiveCostCenter(
            r,
            funcLocKey,
            ctx.funcLocMap,
            ctx.idxCostCenter
        );
        if (shouldSkipLevel5Stas13Row(funcLocKey, effectiveCostCenter)) {
            continue;
        }

        const rawDesc = ctx.funclocDescRaw(r);
        const plant = getFuncLocSecondSegment(r[ctx.idxFuncLoc]);
        const mappedEqunr = lookupEquipmentNumber(
            plant,
            rawDesc,
            equipmentByPlantDesc
        );
        const equnr =
            mappedEqunr || buildMaintenanceItemFallbackEqunr(plant, rawDesc);
        if (!equnr) continue;

        if (!mappedEqunr) {
            console.warn(
                `  [to-delete] EQUNR tidak ditemukan di mapping — ${ctx.sourceName} baris ${ctx.dataStartExcelRow + rowIndex}: ${rawDesc || "(kosong)"} → fallback ${equnr}`
            );
        }

        if (seenInFile.has(equnr) || usedToDeleteEqunrGlobal.has(equnr)) {
            continue;
        }

        seenInFile.add(equnr);
        usedToDeleteEqunrGlobal.add(equnr);
        output.push({
            primary: equnr,
            oldName: ctx.funclocDescText(r)
        });
    }

    return output;
}

async function buildEquipmentFile(
    sourcePath,
    sourceName,
    outputDir,
    driveFile,
    toDeleteOutputDir = null
) {

    await ensureFunclocDescAliasMap();
    await ensureCwcPlannerLookups();
    const abcRules = ensureAbcRules();

    const workbook = XLSX.readFile(sourcePath, { cellStyles: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const layout = findDataLayout(rows);
    const headerRow = rows[layout.headerRowIndex];

    if (!headerRow) {
        console.log("Header not found:", sourceName);
        return null;
    }

    const columnAudit = auditProtectYColumns(headerRow);

    secondRowDump.push({
        file: sourceName,
        excelRowNumber: layout.headerRowIndex + 1,
        cells: headerRow.map(cellToJson),
        missing_column: columnAudit.missingColumn
    });

    columnValidationRows.push({
        fileName: sourceName,
        resolutions: resolutionsToValidationMap(columnAudit.resolutions)
    });

    const dataRows = rows.slice(layout.headerRowIndex + 1);

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);

    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxFunclocDesc = findFunclocDescAfterColumnIndex(headerRow);

    if (idxFuncLoc === undefined) {
        console.log("No FUNCTIONAL LOCATION column:", sourceName);
        return null;
    }

    if (idxFunclocDesc === undefined) {
        console.log("No FUNCTLOC DESC AFTER column:", sourceName);
    }

    function val(r, name) {
        const idx = colIndex[normalizeHeader(name)];
        return idx !== undefined ? r[idx] ?? null : null;
    }

    const idxCostCenter = findCostCenterColumnIndex(headerRow);
    const equipGroupCol = resolveEquipmentGroupAfterColumn(
        headerRow,
        columnAudit.colIndex
    );
    const idxEquipGroupAfter = equipGroupCol.columnIndex;
    const equipGroupSource = equipGroupCol.source;

    if (idxEquipGroupAfter === undefined) {
        console.warn(
            `  [red-skip] kolom EQUIPMENT GROUP AFTER tidak ditemukan — ${sourceName}`
        );
    } else if (equipGroupSource === "equipment_group_plain") {
        console.log(
            `  [red-skip] kolom alternatif EQUIPMENT GROUP (setelah COST CENTER AFTER) — ${sourceName}`
        );
    }

    const redEquipGroupRowIndices = await collectRedEquipGroupRowIndices({
        sourcePath,
        colIndex0: idxEquipGroupAfter,
        dataRowCount: dataRows.length,
        dataStartExcelRow: layout.dataStartExcelRow,
        driveFile: driveFile ?? null,
        sheetsApi: sheets
    });

    const redSkipCount = redEquipGroupRowIndices.size;
    totalRedEquipGroupSkipped += redSkipCount;
    if (redSkipCount > 0) {
        const redLabel =
            equipGroupSource === "equipment_group_plain"
                ? "EQUIPMENT GROUP alternatif setelah COST CENTER AFTER merah"
                : "EQUIPMENT GROUP AFTER merah";
        console.log(`  [red-skip] ${redSkipCount} baris (${redLabel}) — ${sourceName}`);
    }

    const funcLocMap = {};

    for (const r of dataRows) {
        const code = normalizeFuncLocKey(r[idxFuncLoc]);
        if (code) funcLocMap[code] = r;
    }

    function getValueWithParent(r, columnName) {

        const idx = colIndex[normalizeHeader(columnName)];
        if (idx === undefined) return null;

        let value = r[idx];

        if (value !== null && value !== "" && value !== undefined) {
            return value;
        }

        let funcLoc = r[idxFuncLoc];

        while (funcLoc && funcLoc.includes("-")) {

            funcLoc = funcLoc.substring(0, funcLoc.lastIndexOf("-"));

            const parentRow = funcLocMap[funcLoc];

            if (!parentRow) continue;

            const parentVal = parentRow[idx];

            if (parentVal !== null && parentVal !== "" && parentVal !== undefined) {
                return parentVal;
            }
        }

        return null;
    }

    /**
     * Level = jumlah kelompok pemisah strip (-) pada FUNCTIONAL LOCATION AFTER.
     * Contoh: mPALM-7F06-0005-0001-0001-0003-0006 → 7 level.
     */
    function funcLocLevel(code) {
        if (!code || typeof code !== "string") return 0;
        return code.trim().split("-").filter(s => s.length > 0).length;
    }

    function funclocDescRaw(r) {
        if (idxFunclocDesc === undefined) return "";
        const v = r[idxFunclocDesc];
        return v !== null && v !== undefined && v !== ""
            ? fixFunclocDescSourceTypos(String(v))
            : "";
    }

    /** Deskripsi untuk template SHTXT* (alias jika ada di mapping). */
    function funclocDescText(r) {
        return applyFunclocDescAlias(funclocDescRaw(r));
    }

    /** Level ≥ 5 masuk equipment, kecuali deskripsi mengandung pola pengecualian. */
    function isEquipmentRow(r) {
        const code = r[idxFuncLoc];
        if (funcLocLevel(code) < 5) return false;
        const desc = String(funclocDescRaw(r) ?? "");
        if (/drive\s*unit/i.test(desc)) return false;
        if (/accessories\s+weigh\s*bridge/i.test(desc)) return false;
        return true;
    }

    const equipmentRows = [];
    const rowIndexByRow = new Map();
    for (let i = 0; i < dataRows.length; i++) {
        const r = dataRows[i];
        if (redEquipGroupRowIndices.has(i)) continue;
        if (!isEquipmentRow(r)) continue;
        equipmentRows.push(r);
        rowIndexByRow.set(r, i);
    }

    /** @type {Array<{ r: Array<unknown>, rowIndex: number }>} */
    const toDeleteCandidates = [];
    if (equipGroupSource === "equipment_group_plain") {
        for (let i = 0; i < dataRows.length; i++) {
            const r = dataRows[i];
            if (!redEquipGroupRowIndices.has(i)) continue;
            if (!isEquipmentRow(r)) continue;
            toDeleteCandidates.push({ r, rowIndex: i });
        }
    }

    const { mod: egMod, mapping: egMapping } =
        await ensureEquipmentGroupMapping();
    const egInputRows = equipmentRows.map(r => ({
        rowIndex: rowIndexByRow.get(r),
        funcLocNorm: normalizeFuncLocKey(r[idxFuncLoc]),
        desc: funclocDescRaw(r)
    }));
    const egByRowIndex = egMod.resolveEquipmentGroupAssignmentsForRows(
        egInputRows,
        egMapping
    );

    function equipmentGroupResourceForRow(r) {
        if (idxEquipGroupAfter !== undefined) {
            const v = r[idxEquipGroupAfter];
            if (v !== null && v !== undefined && v !== "") return v;
        }
        return val(r, "EQUIPMENT GROUP AFTER");
    }

    function resolveEqart(r) {
        const rowIndex = rowIndexByRow.get(r);
        const resolved =
            rowIndex !== undefined ? egByRowIndex.get(rowIndex) : null;
        const mapped = resolved?.value ?? "";
        if (!isEmpty(mapped)) return mapped;

        abortOnMissingEquipmentGroup({
            sourceFile: sourceName,
            funcLoc: val(r, "FUNCTIONAL LOCATION AFTER"),
            funcDesc: funclocDescRaw(r),
            equipmentGroupResource: equipmentGroupResourceForRow(r)
        });
    }

    /** BEGRU*, SWERK*, IWERK* & WERGW: satu nilai per file output. */
    const fileBegru = resolveFileBegru(
        dataRows,
        equipmentRows,
        val,
        getValueWithParent,
        idxFuncLoc
    );

    if (
        (equipmentRows.length > 0 || toDeleteCandidates.length > 0) &&
        isEmpty(fileBegru)
    ) {
        console.log("Skip (BEGRU* / plant tidak ditemukan):", sourceName);
        return null;
    }

    await logRegionalMappingForFile(fileBegru, sourceName);

    function buildEquipmentTemplateRow(r, eqartValue, dedupeState) {
        const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);

        if (dedupeState?.usedFuncLocKeys && funcLocKey) {
            if (dedupeState.usedFuncLocKeys.has(funcLocKey)) return null;
        }

        const capacity = normalizeCapacityValue(
            getValueWithParent(r, "CAPACITY")
        );
        const year = resolveBaujjYear(r, getValueWithParent);

        const capOut =
            capacity === dedupeState.lastCapacity ? null : capacity;
        const yearOut = year === dedupeState.lastYear ? null : year;

        const effectiveCostCenter = getEffectiveCostCenter(
            r,
            funcLocKey,
            funcLocMap,
            idxCostCenter
        );
        const plannerGroup = resolvePlannerGroup(
            r,
            funcLocKey,
            effectiveCostCenter,
            val,
            fileBegru
        );
        const workCenter = resolveWorkCenter(
            r,
            funcLocKey,
            effectiveCostCenter,
            val,
            fileBegru
        );

        if (shouldSkipLevel5Stas13Row(funcLocKey, effectiveCostCenter)) {
            return null;
        }

        if (dedupeState?.usedFuncLocKeys && funcLocKey) {
            dedupeState.usedFuncLocKeys.add(funcLocKey);
        }
        dedupeState.lastCapacity = capacity;
        dedupeState.lastYear = year;

        return [
            null,
            1,
            fileBegru,
            capOut,
            eqartValue,
            val(r, "MERK"),
            yearOut,
            null,
            funclocDescText(r),
            null,
            fileBegru,
            "TEK",
            resolveAbcIndicator(r, funcLocKey, val, abcRules),
            "PALM",
            idxCostCenter !== undefined ? r[idxCostCenter] ?? null : null,
            fileBegru,
            plannerGroup,
            workCenter,
            fileBegru,
            "ZPM_PTPN",
            val(r, "FUNCTIONAL LOCATION AFTER"),
            null,
            null
        ];
    }

    const output = [];
    const mainDedupe = { usedFuncLocKeys: usedFuncLocAfter, lastCapacity: null, lastYear: null };

    for (const r of equipmentRows) {
        const row = buildEquipmentTemplateRow(r, resolveEqart(r), mainDedupe);
        if (row) output.push(row);
    }

    const toDeleteOutput = buildToDeleteRows(toDeleteCandidates, {
        sourceName,
        dataStartExcelRow: layout.dataStartExcelRow,
        idxFuncLoc,
        idxCostCenter,
        funcLocMap,
        funclocDescRaw,
        funclocDescText
    });

    const newName = lsmwOutputFileName(sourceName);
    const newPath = path.join(outputDir, newName);

    await writeEquipmentOutputRows(output, newPath);

    let toDeletePath = null;
    if (toDeleteOutputDir && toDeleteOutput.length > 0) {
        fs.mkdirSync(toDeleteOutputDir, { recursive: true });
        toDeletePath = path.join(toDeleteOutputDir, newName);
        await writeToDeleteOutputRows(toDeleteOutput, toDeletePath);
        console.log(
            `  [to-delete] ${toDeleteOutput.length} baris → ${toDeletePath}`
        );
    }

    return {
        path: newPath,
        name: newName,
        outputRows: output,
        toDeleteRows: toDeleteOutput,
        redSkipped: redSkipCount
    };
}

async function processFolder(
    sourceFolderId,
    localDir,
    driveTargetId,
    rowCollector,
    toDeleteRowCollector,
    toDeleteOutputDir = null
) {

    const files = await listFolderChildrenWithRetry(sourceFolderId);

    for (const file of files) {

        if (file.mimeType === "application/vnd.google-apps.folder") {

            const childLocal = path.join(localDir, file.name);
            fs.mkdirSync(childLocal, { recursive: true });

            let childDrive = null;
            if (driveTargetId) {
                childDrive = await createFolderIfNotExists(
                    file.name,
                    driveTargetId
                );
            }

            await processFolder(
                file.id,
                childLocal,
                childDrive,
                rowCollector,
                toDeleteRowCollector,
                toDeleteOutputDir
            );
            continue;
        }

        if (!isProcessableSpreadsheet(file)) {
            if (
                file.mimeType !== "application/vnd.google-apps.folder" &&
                file.mimeType !== "application/vnd.google-apps.shortcut"
            ) {
                console.log(
                    `Skipping unsupported: ${file.name} (${file.mimeType})`
                );
            }
            continue;
        }

        console.log("Downloading:", file.name);

        const { localFile, sourceName } = await downloadSpreadsheetWithRetry(file);
        await sleep(300);

        const equipment = await buildEquipmentFile(
            localFile,
            sourceName,
            localDir,
            file,
            toDeleteOutputDir
        );

        if (!equipment) continue;

        if (rowCollector && equipment.outputRows?.length) {
            rowCollector.push(...equipment.outputRows);
        }

        if (toDeleteRowCollector && equipment.toDeleteRows?.length) {
            toDeleteRowCollector.push(...equipment.toDeleteRows);
        }

        console.log("Saved:", equipment.path);

        if (driveTargetId) {
            await uploadFile(equipment.path, equipment.name, driveTargetId);
        }
    }
}

async function main() {

    secondRowDump = [];
    columnValidationRows = [];
    usedFuncLocAfter.clear();
    usedToDeleteEqunrGlobal.clear();
    totalRedEquipGroupSkipped = 0;

    const equipmentMapping = loadEquipmentNumberMapping();
    equipmentByPlantDesc = equipmentMapping.byPlantDesc;
    console.log(
        `[mapping EQUNR to-delete] ${equipmentByPlantDesc.size} key aktif` +
            (equipmentMapping.duplicateKeys
                ? ` (${equipmentMapping.duplicateKeys} duplikat diabaikan)`
                : "")
    );

    const cli = parseLsmwCli(process.argv.slice(2));
    outputDriveFolderId = cli.outputDriveFolderId;
    const sourceRootFolderId = await resolveProtectYSourceFolderId(cli);

    fs.mkdirSync(LOCAL_OUTPUT_ROOT, { recursive: true });
    prepareTempDownloads();

    try {
        await initDrive();

        console.log("Protect Y source →", sourceRootFolderId);
        if (outputDriveFolderId) {
            console.log("Drive upload enabled →", outputDriveFolderId);
        } else {
            console.log("Local output only →", path.resolve(LOCAL_OUTPUT_ROOT));
        }

        const regionalFolders = await withRetry(
            async () => {
                const res = await drive.files.list({
                    q: `'${sourceRootFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
                    fields: "files(id,name)",
                    pageSize: 1000
                });
                return (res.data.files ?? []).filter(f =>
                    isRegionalFolderName(f.name)
                );
            },
            {
                ...DRIVE_RETRY,
                label: `list regional folders ${sourceRootFolderId}`
            }
        );

        const regionalRowsByName = new Map();
        const allRows = [];
        const toDeleteRegionalRowsByName = new Map();
        const toDeleteAllRows = [];

        for (const folder of regionalFolders) {

            console.log(`Processing ${folder.name}`);

            const localRegionalDir = path.join(LOCAL_OUTPUT_ROOT, folder.name);
            fs.mkdirSync(localRegionalDir, { recursive: true });

            const toDeleteRegionalDir = path.join(TO_DELETE_ROOT, folder.name);
            fs.mkdirSync(toDeleteRegionalDir, { recursive: true });

            let driveRegionalTarget = null;
            if (outputDriveFolderId) {
                driveRegionalTarget = await createFolderIfNotExists(
                    folder.name,
                    outputDriveFolderId
                );
            }

            const regionalRows = [];
            const toDeleteRegionalRows = [];
            await processFolder(
                folder.id,
                localRegionalDir,
                driveRegionalTarget,
                regionalRows,
                toDeleteRegionalRows,
                toDeleteRegionalDir
            );

            if (regionalRows.length > 0) {
                regionalRowsByName.set(folder.name, regionalRows);
                allRows.push(...regionalRows);
            }

            if (toDeleteRegionalRows.length > 0) {
                toDeleteRegionalRowsByName.set(
                    folder.name,
                    toDeleteRegionalRows
                );
                toDeleteAllRows.push(...toDeleteRegionalRows);
            }
        }

        if (regionalFolders.length === 0) {
            console.log("No REGIONAL folder found under source root");
        }

        await writeRegionalAndAllPksAggregates({
            runDir: LOCAL_OUTPUT_ROOT,
            regionalRowsByName,
            allRows,
            label: "baris equipment",
            writeRows: (outPath, rows) => writeEquipmentOutputRows(rows, outPath)
        });

        if (toDeleteAllRows.length > 0) {
            await writeRegionalAndAllPksAggregates({
                runDir: TO_DELETE_ROOT,
                regionalRowsByName: toDeleteRegionalRowsByName,
                allRows: toDeleteAllRows,
                label: "baris to-delete",
                writeRows: (outPath, rows) =>
                    writeToDeleteOutputRows(rows, outPath)
            });
            console.log(
                `to-delete: ${toDeleteAllRows.length} baris di ${path.resolve(TO_DELETE_ROOT)}`
            );
        }

        if (outputDriveFolderId) {
            for (const [regional, rows] of regionalRowsByName) {
                if (!rows || rows.length === 0) continue;
                const aggName = regionalAggregateFileName(regional);
                const aggPath = path.join(LOCAL_OUTPUT_ROOT, aggName);
                await uploadFile(aggPath, aggName, outputDriveFolderId);
            }
            if (allRows.length > 0) {
                const allPksPath = path.join(LOCAL_OUTPUT_ROOT, ALL_PKS_FILE_NAME);
                await uploadFile(
                    allPksPath,
                    ALL_PKS_FILE_NAME,
                    outputDriveFolderId
                );
            }
        }

        fs.writeFileSync(
            COLUMN_JSON_PATH,
            JSON.stringify(secondRowDump, null, 2),
            "utf8"
        );
        console.log("Wrote", COLUMN_JSON_PATH, `(${secondRowDump.length} files)`);

        const filesWithMissingColumns = secondRowDump.filter(
            e => (e.missing_column?.length ?? 0) > 0
        ).length;
        await writeEquipmentColumnValidationExcel(
            VALIDATION_XLSX_PATH,
            columnValidationRows
        );
        console.log(
            "Wrote",
            VALIDATION_XLSX_PATH,
            `(${columnValidationRows.length} files, ${filesWithMissingColumns} dengan missing_column)`
        );
        console.log(
            "Total baris di-skip (EQUIPMENT GROUP AFTER merah):",
            totalRedEquipGroupSkipped
        );

        console.log("DONE");
    } finally {
        cleanupTempDownloads();
    }
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});