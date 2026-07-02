/**
 * gathering-investasi-lpp.cjs
 *
 * Mengumpulkan data dari LPP Resource (seluruh PKS di subfolder Regional*) dan folder Investasi di Google Drive,
 * mencocokkan EQKTU BEFORE dari LPP dengan Nama Mesin/Peralatan & Mesin/Alat No. dari Investasi.
 *
 * Melakukan deduplikasi agar EQKTU BEFORE unik untuk mempercepat proses matching dan menghindari duplikasi data.
 *
 * Menyimpan secara bertahap dalam berkas Excel intermediate:
 * 1. lpp.xlsx (2 kolom data LPP unik)
 * 2. investasi/ — mirror berkas Investasi (style asli + kolom dendogram)
 * 3. match.xlsx (hasil pencocokan deduplikasi)
 * 4. match_per_pks.xlsx (hasil per baris PKS + sheet master nama baru)
 *
 * Jalankan dengan:
 *   pnpm run gathering-investasi-lpp
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const ExcelJS = require("exceljs");
const { authenticate } = require("@google-cloud/local-auth");
const { input } = require("@inquirer/prompts");

const { worksheetToDenseRows } = require("./utils/equipment_excel_io.cjs");
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
    findLppGatheringLayout,
    findFunclocDescAfterLevel123ColumnIndex,
    findRegionalColumnIndex,
    findCostCenterAfterColumnIndex
} = require("./utils/equipment_gathering_columns.cjs");

const { buildColumnIndexFirstWins } = require("./utils/lsmw_lookups.cjs");
const { normalizeHeader } = require("./utils/lsmw_cell_fill.cjs");
const { similarity } = require("./utils/string_similarity.cjs");
const {
    ensureFunclocDescAliasMap,
    applyFunclocDescAlias
} = require("./utils/funcloc_desc_alias.cjs");

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OUTPUT_DIR = path.join("Output", "gathering-investasi-lpp");
const SAWIT_FILE_RE = /sawit/i;
const DENDOGRAM_COL_HEADER = "Nama Mesin/Peralatan Sesuai Dendogram";
const INVESTASI_DATA_SHEET_NAMES = new Set(
    [
        "KSO DJABA",
        "N03",
        "N02",
        "N04",
        "N14",
        "N05",
        "N06",
        "N13",
        "N01",
        "N07"
    ].map(name => name.toUpperCase())
);
const GRAY_STASIUN_FILL = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9D9D9" }
};

let drive;

function isInvestasiDataSheetName(sheetName) {
    return INVESTASI_DATA_SHEET_NAMES.has(
        String(sheetName ?? "").trim().toUpperCase()
    );
}

function normalizeInvestasiHeader(value) {
    return String(value ?? "")
        .replace(/[\r\n\t]+/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

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

async function initDrive() {
    if (drive) return;

    if (!fs.existsSync(OAUTH_PATH)) {
        throw new Error(`File ${OAUTH_PATH} tidak ditemukan. Pastikan file konfigurasi OAuth sudah ada.`);
    }

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

function getVal(row, idx) {
    if (idx === undefined || idx === null || idx < 0 || idx >= row.length) return "";
    const val = row[idx];
    if (val === null || val === undefined) return "";
    return String(val).trim();
}

function findFirstColumnValue(dataRows, idx) {
    if (idx === undefined || idx === null || idx < 0) return "";
    for (const r of dataRows) {
        const v = getVal(r, idx);
        if (v) return v;
    }
    return "";
}

function cellContainsStasiun(value) {
    return /stasiun/i.test(String(value ?? ""));
}

function applyGrayFillIfStasiun(cell, value) {
    if (cellContainsStasiun(value)) {
        cell.fill = GRAY_STASIUN_FILL;
    }
}

function extractCellText(value) {
    if (value === null || value === undefined) return "";
    const t = typeof value;
    if (t === "string" || t === "number" || t === "boolean") return String(value).trim();
    if (value instanceof Date) return String(value).trim();
    if (t === "object") {
        if (Array.isArray(value.richText)) {
            return value.richText
                .map(rt => (rt && typeof rt.text === "string" ? rt.text : ""))
                .join("")
                .trim();
        }
        if (typeof value.text === "string") return value.text.trim();
        if (value.result !== undefined && value.result !== null) {
            return extractCellText(value.result);
        }
    }
    return String(value).trim();
}

function getValFromWorksheetRow(row, idx0) {
    if (idx0 < 0) return "";
    return extractCellText(row.getCell(idx0 + 1).value);
}

function getWorksheetMaxCol(ws) {
    return Math.max(
        ws.actualColumnCount ?? 0,
        ws.columnCount ?? 0,
        ws.dimensions?.right ?? 0
    );
}

function copyExcelCellStyle(srcCell, dstCell) {
    if (!srcCell || !dstCell) return;
    for (const key of ["font", "alignment", "border", "fill", "numFmt", "protection"]) {
        const v = srcCell[key];
        if (v === undefined || v === null) continue;
        dstCell[key] =
            typeof structuredClone === "function"
                ? structuredClone(v)
                : JSON.parse(JSON.stringify(v));
    }
}

function cloneCellFill(cell) {
    const fill = cell?.fill;
    if (!fill) return null;
    return typeof structuredClone === "function"
        ? structuredClone(fill)
        : JSON.parse(JSON.stringify(fill));
}

function captureLppRowFills(wsRow, indices) {
    const get = idx => {
        if (idx === undefined || idx < 0) return null;
        return cloneCellFill(wsRow.getCell(idx + 1));
    };
    return {
        namaBaru: get(indices.funcloc),
        regional: get(indices.regional),
        pom: get(indices.pom),
        costCenterAfter: get(indices.costCenterAfter)
    };
}

function applyCellFill(dstCell, fill) {
    if (fill) dstCell.fill = fill;
}

function flattenWorksheetSharedFormulas(ws) {
    ws.eachRow({ includeEmpty: false }, row => {
        row.eachCell({ includeEmpty: false }, cell => {
            const val = cell.value;
            if (!val || typeof val !== "object" || !val.sharedFormula) return;

            const master = ws.getCell(val.sharedFormula);
            const masterVal = master.value;
            const formula =
                (typeof masterVal === "object" && masterVal?.formula) ||
                master.formula ||
                val.formula;

            cell.value = {
                formula,
                result: val.result ?? master.result ?? masterVal?.result
            };
        });
    });
}

function worksheetHeaderRowsForLayout(ws, limit = 30) {
    const maxCol = getWorksheetMaxCol(ws);
    const sheetRows = Math.min(
        limit,
        Math.max(ws.rowCount || 0, ws.dimensions?.bottom ?? 0)
    );
    const rows = [];
    for (let r = 1; r <= sheetRows; r++) {
        const row = ws.getRow(r);
        const arr = [];
        for (let c = 1; c <= maxCol; c++) {
            arr[c - 1] = extractCellText(row.getCell(c).value);
        }
        rows.push(arr);
    }
    return rows;
}

/**
 * Tambahkan sheet master nama baru (global LPP) ke workbook.
 */
function appendMasterNamaBaruSheet(workbook, lppMasterRows) {
    const masterWs = workbook.addWorksheet("master nama baru", {
        views: [{ state: "frozen", ySplit: 1 }]
    });
    const masterHeaders = ["NAMA BARU", "Regional", "POM", "COST CENTER AFTER"];
    masterWs.addRow(masterHeaders);
    masterWs.getRow(1).font = { bold: true };

    for (const m of lppMasterRows) {
        const masterRow = masterWs.addRow([
            m.namaBaru,
            m.regional,
            m.pom,
            m.costCenterAfter
        ]);
        const masterValues = [m.namaBaru, m.regional, m.pom, m.costCenterAfter];
        const fillKeys = ["namaBaru", "regional", "pom", "costCenterAfter"];
        for (let c = 0; c < masterValues.length; c++) {
            const cell = masterRow.getCell(c + 1);
            applyCellFill(cell, m.fills?.[fillKeys[c]]);
            applyGrayFillIfStasiun(cell, masterValues[c]);
        }
    }

    masterWs.columns = [
        { width: 45 },
        { width: 30 },
        { width: 25 },
        { width: 25 }
    ];

    if (lppMasterRows.length > 0) {
        masterWs.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + lppMasterRows.length, column: 4 }
        };
    }

    return masterWs;
}

/**
 * Salin berkas investasi dengan style asli + kolom dendogram setelah Mesin/Alat No.
 */
async function writeInvestasiMirrorWithDendogram(
    localPath,
    outputPath,
    matchByKey,
    lppMasterRows
) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(localPath);

    for (const ws of wb.worksheets) {
        if (!ws) continue;
        if (!isInvestasiDataSheetName(ws.name)) continue;

        const layout = findInvestasiLayout(worksheetHeaderRowsForLayout(ws));
        if (!layout || layout.idxAlatNo === -1) continue;

        const { idxNamaMesin, idxAlatNo, headerRowIndex } = layout;
        const insertCol = idxAlatNo + 2;
        const headerRowNum = headerRowIndex + 1;
        const rowCount = Math.max(
            ws.rowCount || 0,
            ws.dimensions?.bottom ?? 0,
            headerRowNum + 1
        );
        const emptyInsert = Array.from({ length: rowCount }, () => null);

        flattenWorksheetSharedFormulas(ws);
        ws.spliceColumns(insertCol, 0, emptyInsert);

        const refHeaderCell = ws.getRow(headerRowNum).getCell(insertCol - 1);
        const dendogramHeaderCell = ws.getRow(headerRowNum).getCell(insertCol);
        dendogramHeaderCell.value = DENDOGRAM_COL_HEADER;
        copyExcelCellStyle(refHeaderCell, dendogramHeaderCell);

        const srcColWidth = ws.getColumn(insertCol - 1).width;
        ws.getColumn(insertCol).width = Math.max(srcColWidth || 25, 45);

        for (let r = headerRowNum + 1; r <= rowCount; r++) {
            const row = ws.getRow(r);
            const namaMesin = getValFromWorksheetRow(row, idxNamaMesin);
            const alatNo = getValFromWorksheetRow(row, idxAlatNo);
            const matchKey = buildInvestasiMatchKey(namaMesin, alatNo);
            const namaBaru = matchKey ? (matchByKey.get(matchKey)?.namaBaru ?? "") : "";

            const refCell = row.getCell(insertCol - 1);
            const dendogramCell = row.getCell(insertCol);
            copyExcelCellStyle(refCell, dendogramCell);
            if (namaBaru) {
                dendogramCell.value = namaBaru;
            }
        }
    }

    appendMasterNamaBaruSheet(wb, lppMasterRows);
    await wb.xlsx.writeFile(outputPath);
}

/**
 * Mendeteksi kolom Nama Mesin/Peralatan dan Mesin/Alat No.
 */
function findInvestasiColumns(headerRow) {
    let idxNamaMesin = -1;
    let idxAlatNo = -1;
    let idxNamaPabrik = -1;
    let idxPtpn = -1;
    for (let i = 0; i < headerRow.length; i++) {
        const h = normalizeInvestasiHeader(headerRow[i]);
        if (!h) continue;

        // Cari Nama Mesin/Peralatan
        if (
            h === "NAMA MESIN/PERALATAN" ||
            h === "NAMA MESIN PERALATAN" ||
            h === "NAMA MESIN" ||
            h === "MESIN/PERALATAN" ||
            h.includes("NAMA MESIN") ||
            h.includes("MESIN/PERALATAN") ||
            h.includes("MESIN PERALATAN")
        ) {
            if (idxNamaMesin === -1) idxNamaMesin = i;
        }

        // Cari Mesin/Alat No.
        if (
            h === "MESIN/ALAT NO." ||
            h === "MESIN/ALAT NO" ||
            h === "ALAT NO" ||
            h === "NO. ALAT" ||
            h === "NO ALAT" ||
            h === "NO." ||
            h === "NO" ||
            h.includes("ALAT NO") ||
            h.includes("MESIN NO") ||
            h.includes("MESIN/ALAT NO")
        ) {
            if (idxAlatNo === -1) idxAlatNo = i;
        }

        // Cari Nama Pabrik
        if (
            h === "NAMA PABRIK" ||
            h === "PABRIK" ||
            h === "NAMA PKS" ||
            h === "PKS" ||
            h.includes("NAMA PABRIK") ||
            h.includes("PABRIK") ||
            h.includes("NAMA PKS")
        ) {
            if (idxNamaPabrik === -1) idxNamaPabrik = i;
        }

        if (h === "PTPN" || h.includes("PTPN")) {
            if (idxPtpn === -1) idxPtpn = i;
        }
    }
    return { idxNamaMesin, idxAlatNo, idxNamaPabrik, idxPtpn };
}

/**
 * Mencari layout header Investasi
 */
function findInvestasiLayout(rows) {
    const limit = Math.min(rows.length, 30);
    for (let i = 0; i < limit; i++) {
        const { idxNamaMesin, idxAlatNo, idxNamaPabrik, idxPtpn } = findInvestasiColumns(rows[i]);
        if (idxNamaMesin !== -1) {
            return { headerRowIndex: i, idxNamaMesin, idxAlatNo, idxNamaPabrik, idxPtpn };
        }
    }
    return null;
}

/**
 * Membaca semua worksheet dari file Investasi
 */
async function readAllSheetsRows(localPath) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(localPath);
    const results = [];
    for (const ws of wb.worksheets) {
        if (!ws) continue;
        const rawRows = [];
        const sheetRows = Math.max(ws.rowCount || 0, ws.dimensions?.bottom ?? 0);
        let maxCol = Math.max(ws.actualColumnCount ?? 0, ws.columnCount ?? 0);
        if (ws.dimensions?.right) {
            maxCol = Math.max(maxCol, ws.dimensions.right);
        }

        for (let r = 1; r <= sheetRows; r++) {
            const row = ws.getRow(r);
            const arr = [];
            for (let c = 1; c <= maxCol; c++) {
                const cell = row.getCell(c).value;
                const extract = v => {
                    if (v === null || v === undefined) return "";
                    const t = typeof v;
                    if (t === "string" || t === "number" || t === "boolean") return v;
                    if (v instanceof Date) return v;
                    if (t === "object") {
                        if (Array.isArray(v.richText)) {
                            return v.richText
                                .map(rt => (rt && typeof rt.text === "string" ? rt.text : ""))
                                .join("");
                        }
                        if (typeof v.text === "string") return v.text;
                        if (v.result !== undefined && v.result !== null) {
                            return extract(v.result);
                        }
                    }
                    return v;
                };
                arr[c - 1] = extract(cell);
            }
            rawRows.push(arr);
        }
        if (rawRows.length > 0) {
            results.push({ name: ws.name, rows: rawRows });
        }
    }
    return results;
}

/**
 * Normalisasi string untuk pencocokan (case-insensitive, membuang spasi/simbol/noise word).
 */
function buildInvestasiMatchKey(namaMesin, alatNo) {
    if (!namaMesin) return "";
    const cleanNama = cleanForMatching(namaMesin);
    const cleanAlat = cleanForMatching(alatNo);
    let fullStr = "";
    if (cleanAlat && cleanNama.endsWith(cleanAlat)) {
        fullStr = namaMesin;
    } else {
        fullStr = `${namaMesin} ${alatNo}`.trim();
    }
    return cleanForMatching(fullStr);
}

function cleanForMatching(str) {
    if (!str) return "";
    return String(str)
        .toUpperCase()
        .replace(/\bNO\b\.?/g, " ") // membuang kata "NO" atau "NO."
        .replace(/\bNOMOR\b/g, " ") // membuang kata "NOMOR"
        .replace(/[^A-Z0-9]/g, " ") // membuang karakter non-alfanumerik
        .replace(/\b0+(\d+)\b/g, "$1") // membuang angka nol di depan (mis. 01 -> 1)
        .replace(/\s+/g, "") // membuang semua spasi
        .trim();
}

async function promptUrl(message, extractFn) {
    return input({
        message,
        validate: value => {
            const trimmed = String(value ?? "").trim();
            if (!trimmed) return "URL tidak boleh kosong";
            try {
                extractFn(trimmed);
                return true;
            } catch {
                return "URL Google Drive tidak valid";
            }
        }
    });
}

async function main() {
    const stamp = wibTimestampForFilename();
    const runDir = path.join(OUTPUT_DIR, stamp);

    prepareTempDownloads();

    try {
        await initDrive();
        await ensureFunclocDescAliasMap();

        const lppFolderUrl = await promptUrl(
            "URL folder Google Drive — LPP Resource (subfolder REGIONAL*):",
            extractFolderIdFromUrl
        );
        const investasiFolderUrl = await promptUrl(
            "URL folder Google Drive — Investasi:",
            extractFolderIdFromUrl
        );

        const lppRootFolderId = extractFolderIdFromUrl(String(lppFolderUrl).trim());
        const investasiRootFolderId = extractFolderIdFromUrl(String(investasiFolderUrl).trim());

        // Membuat folder output dengan timestamp WIB
        fs.mkdirSync(runDir, { recursive: true });
        console.log(`\nFolder output dibuat: ${runDir}`);

        // --- TAHAP 1: PROSES LPP ---
        console.log("\nMemuat daftar berkas LPP Resource...");
        const lppEntries = await listRegionalSpreadsheets(lppRootFolderId);
        console.log(`  Menemukan ${lppEntries.length} berkas LPP.`);

        const lppRows = [];
        const lppMasterRows = [];
        const seenEqktu = new Set();

        console.log("\nMembaca berkas LPP...");
        for (const { regional, file } of lppEntries) {
            console.log(`  Membaca: [${regional}] ${file.name}`);
            const { localFile: lppLocal } = await downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                "lpp_"
            );

            const lppWb = new ExcelJS.Workbook();
            await lppWb.xlsx.readFile(lppLocal);
            const lppWs = lppWb.worksheets[0];
            if (!lppWs) {
                console.warn(`     [Skip] Sheet LPP kosong: ${file.name}`);
                continue;
            }

            const lppRawRows = worksheetToDenseRows(lppWs);
            const lppLayout = findLppGatheringLayout(lppRawRows);
            const lppHeader = lppRawRows[lppLayout.headerRowIndex];
            if (!lppHeader) {
                console.warn("     [Skip] Header LPP tidak ditemukan.");
                continue;
            }

            const lppColIndex = buildColumnIndexFirstWins(lppHeader, normalizeHeader);
            const idxEqktuBefore = lppColIndex["EQKTU BEFORE"];
            const idxFunclocLevel123 = findFunclocDescAfterLevel123ColumnIndex(lppHeader);
            const idxRegional = findRegionalColumnIndex(lppHeader);
            const idxPom = lppColIndex.POM;
            const idxCostCenterAfter = findCostCenterAfterColumnIndex(lppHeader);

            if (idxEqktuBefore === undefined || idxFunclocLevel123 === undefined) {
                console.warn("     [Skip] Kolom 'EQKTU BEFORE' atau 'FUNCTLOC DESC. AFTER LEVEL 1,2,3' tidak ditemukan.");
                continue;
            }

            const lppDataRows = lppRawRows.slice(lppLayout.headerRowIndex + 1);
            const fileRegional =
                findFirstColumnValue(lppDataRows, idxRegional) || regional;
            const filePom = findFirstColumnValue(lppDataRows, idxPom);
            let fileExtracted = 0;

            for (let rowIndex = 0; rowIndex < lppDataRows.length; rowIndex++) {
                const r = lppDataRows[rowIndex];
                const eqktuBefore = getVal(r, idxEqktuBefore);
                const funcloc = getVal(r, idxFunclocLevel123);
                const rowRegional = getVal(r, idxRegional) || fileRegional;
                const rowPom = getVal(r, idxPom) || filePom;
                const rowCostCenterAfter = getVal(r, idxCostCenterAfter);
                const sheetRowNum = lppLayout.headerRowIndex + 2 + rowIndex;
                const wsRow = lppWs.getRow(sheetRowNum);

                lppMasterRows.push({
                    namaBaru: applyFunclocDescAlias(funcloc),
                    regional: rowRegional,
                    pom: rowPom,
                    costCenterAfter: rowCostCenterAfter,
                    fills: captureLppRowFills(wsRow, {
                        funcloc: idxFunclocLevel123,
                        regional: idxRegional,
                        pom: idxPom,
                        costCenterAfter: idxCostCenterAfter
                    })
                });

                if (eqktuBefore) {
                    const cleanKey = eqktuBefore.trim().toUpperCase();
                    // Deduplikasi EQKTU BEFORE
                    if (!seenEqktu.has(cleanKey)) {
                        seenEqktu.add(cleanKey);
                        lppRows.push({ eqktuBefore, funcloc });
                        fileExtracted++;
                    }
                }
            }
            console.log(`     Ekstrak ${fileExtracted} baris unik.`);
        }

        // Tulis berkas LPP intermediate
        console.log(`\n  Menulis ${lppRows.length} data LPP unik ke lpp.xlsx...`);
        const lppWb = new ExcelJS.Workbook();
        const lppWs = lppWb.addWorksheet("LPP Data");
        lppWs.addRow(["EQKTU BEFORE", "FUNCTLOC DESC. AFTER LEVEL 1,2,3"]);
        lppWs.getRow(1).font = { bold: true };
        for (const row of lppRows) {
            lppWs.addRow([row.eqktuBefore, row.funcloc]);
        }
        lppWs.columns = [{ width: 45 }, { width: 45 }];
        const lppOutputPath = path.join(runDir, "lpp.xlsx");
        await lppWb.xlsx.writeFile(lppOutputPath);
        console.log(`  [OK] Berhasil menyimpan berkas intermediate LPP: ${lppOutputPath}`);

        // --- TAHAP 2: PROSES INVESTASI ---
        console.log("\nMemuat daftar berkas Investasi...");
        const investasiEntries = await listSpreadsheetsRecursive(investasiRootFolderId);
        console.log(`  Menemukan ${investasiEntries.length} berkas Investasi.`);

        const investasiFilesToProcess = [];
        const investasiFilesSkipped = [];
        for (const file of investasiEntries) {
            if (SAWIT_FILE_RE.test(file.name)) {
                investasiFilesToProcess.push(file);
            } else {
                investasiFilesSkipped.push(file.name);
            }
        }
        console.log(
            `  Diproses (nama mengandung "Sawit"): ${investasiFilesToProcess.length}`
        );
        if (investasiFilesSkipped.length > 0) {
            console.log(
                `  Dilewati (tidak mengandung "Sawit"): ${investasiFilesSkipped.length}`
            );
            for (const name of investasiFilesSkipped.sort()) {
                console.log(`     [Skip] ${name}`);
            }
        }

        const investasiCandidates = [];
        const seenInvestasi = new Set();
        const investasiPerPksRows = [];
        const pksSeqCounter = new Map();
        const warnedMissingPtpn = new Set();
        const investasiMirrorSources = [];

        for (const file of investasiFilesToProcess) {
            console.log(`  Membaca: ${file.name}`);
            const { localFile: invLocal } = await downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                "inv_"
            );
            investasiMirrorSources.push({ fileName: file.name, localPath: invLocal });

            const sheets = await readAllSheetsRows(invLocal);
            let fileExtracted = 0;

            for (const sheet of sheets) {
                if (!isInvestasiDataSheetName(sheet.name)) {
                    console.log(
                        `     [Skip sheet] "${sheet.name}" — bukan sheet data investasi (tidak ada di whitelist).`
                    );
                    continue;
                }

                const layout = findInvestasiLayout(sheet.rows);
                if (!layout) continue;

                const { idxNamaMesin, idxAlatNo, idxNamaPabrik, idxPtpn, headerRowIndex } = layout;
                if (idxPtpn === -1) {
                    const warnKey = `${file.name}::${sheet.name}`;
                    if (!warnedMissingPtpn.has(warnKey)) {
                        warnedMissingPtpn.add(warnKey);
                        console.warn(
                            `     [Warn] Kolom PTPN tidak ditemukan di ${file.name} / sheet "${sheet.name}" — kolom Regional akan kosong.`
                        );
                    }
                }
                const dataRows = sheet.rows.slice(headerRowIndex + 1);

                for (const r of dataRows) {
                    const namaMesin = getVal(r, idxNamaMesin);
                    const alatNo = getVal(r, idxAlatNo);
                    let namaPabrik = "";
                    if (idxNamaPabrik !== -1) {
                        namaPabrik = getVal(r, idxNamaPabrik);
                    }

                    // Hanya proses pabrik yang namanya diawali dengan "PKS" (case-insensitive)
                    if (!namaPabrik.toUpperCase().startsWith("PKS")) {
                        continue;
                    }

                    const namaPks = idxPtpn !== -1 ? getVal(r, idxPtpn) : "";
                    const seqInPks = pksSeqCounter.get(namaPabrik) ?? 0;
                    pksSeqCounter.set(namaPabrik, seqInPks + 1);
                    const matchKey = buildInvestasiMatchKey(namaMesin, alatNo);

                    investasiPerPksRows.push({
                        namaMesin,
                        alatNo,
                        namaPabrik,
                        namaPks,
                        matchKey,
                        seqInPks
                    });

                    if (namaMesin && matchKey && !seenInvestasi.has(matchKey)) {
                        seenInvestasi.add(matchKey);
                        investasiCandidates.push({
                            namaMesin,
                            alatNo,
                            namaPabrik,
                            matchKey,
                            matched: false
                        });
                        fileExtracted++;
                    }
                }
            }
            console.log(`     Ekstrak ${fileExtracted} baris unik.`);
        }

        // --- TAHAP 3: PENCOCOKAN & MATCH ---
        console.log("\nMemulai proses pencocokan...");
        const matchedLppIndices = new Set();

        // 1. Exact Match Pass
        let exactMatchCount = 0;
        for (const c of investasiCandidates) {
            let exactMatchIdx = -1;
            for (let i = 0; i < lppRows.length; i++) {
                if (matchedLppIndices.has(i)) continue;
                const lppCleanKey = cleanForMatching(lppRows[i].eqktuBefore);
                if (lppCleanKey === c.matchKey) {
                    exactMatchIdx = i;
                    break;
                }
            }

            if (exactMatchIdx !== -1) {
                matchedLppIndices.add(exactMatchIdx);
                c.matched = true;
                c.matchedLpp = lppRows[exactMatchIdx];
                c.similarityScore = "PERFECT";
                exactMatchCount++;
            }
        }

        // 2. Forced Similarity Match Pass (untuk Investasi yang belum punya pasangan)
        let forcedMatchCount = 0;
        for (const c of investasiCandidates) {
            if (c.matched) continue;

            let bestLppIdx = -1;
            let bestScore = -1;

            for (let i = 0; i < lppRows.length; i++) {
                if (matchedLppIndices.has(i)) continue;
                const score = similarity(lppRows[i].eqktuBefore, `${c.namaMesin} ${c.alatNo}`);
                if (score > bestScore) {
                    bestScore = score;
                    bestLppIdx = i;
                }
            }

            if (bestLppIdx !== -1) {
                matchedLppIndices.add(bestLppIdx);
                c.matched = true;
                c.matchedLpp = lppRows[bestLppIdx];
                c.similarityScore = `${(bestScore * 100).toFixed(1)}%`;
                forcedMatchCount++;
            } else {
                c.matched = false;
            }
        }

        console.log(`  Hasil analisis pencocokan:`);
        console.log(`    - Exact Match ("PERFECT"): ${exactMatchCount}`);
        console.log(`    - Forced Similarity Match: ${forcedMatchCount}`);
        console.log(`    - Investasi Gagal Match: ${investasiCandidates.filter(c => !c.matched).length}`);

        // Gabungkan semuanya ke match.xlsx
        const finalRows = [];
        for (const c of investasiCandidates) {
            if (c.matched) {
                finalRows.push([
                    c.matchedLpp.eqktuBefore,
                    c.matchedLpp.funcloc,
                    c.namaMesin,
                    c.alatNo,
                    c.namaPabrik,
                    c.similarityScore,
                    applyFunclocDescAlias(c.matchedLpp.funcloc)
                ]);
            } else {
                finalRows.push([
                    "",
                    "",
                    c.namaMesin,
                    c.alatNo,
                    c.namaPabrik,
                    "0%",
                    ""
                ]);
            }
        }

        console.log(`  Menulis ${finalRows.length} baris hasil ke match.xlsx...`);
        const matchWb = new ExcelJS.Workbook();
        const matchWs = matchWb.addWorksheet("Match Results", {
            views: [{ state: "frozen", ySplit: 1 }]
        });

        const headers = [
            "EQKTU BEFORE",
            "FUNCTLOC DESC. AFTER LEVEL 1,2,3",
            "Nama Mesin/Peralatan",
            "Mesin/Alat No.",
            "Nama Pabrik",
            "SIMILARITY",
            "NAMA BARU"
        ];

        matchWs.addRow(headers);
        matchWs.getRow(1).font = { bold: true };

        for (const rowData of finalRows) {
            matchWs.addRow(rowData);
        }

        matchWs.columns = [
            { width: 45 },
            { width: 45 },
            { width: 45 },
            { width: 25 },
            { width: 30 },
            { width: 20 },
            { width: 45 }
        ];

        if (finalRows.length > 0) {
            matchWs.autoFilter = {
                from: { row: 1, column: 1 },
                to: { row: 1 + finalRows.length, column: 7 }
            };
        }

        const matchOutputPath = path.join(runDir, "match.xlsx");
        await matchWb.xlsx.writeFile(matchOutputPath);
        console.log(`  [OK] Berhasil menyimpan berkas hasil: ${matchOutputPath}`);

        const matchByKey = new Map();
        for (const c of investasiCandidates) {
            if (!c.matchKey) continue;
            matchByKey.set(c.matchKey, {
                eqktuBefore: c.matched ? c.matchedLpp.eqktuBefore : "",
                funcloc: c.matched ? c.matchedLpp.funcloc : "",
                similarityScore: c.matched ? c.similarityScore : "0%",
                namaBaru: c.matched ? applyFunclocDescAlias(c.matchedLpp.funcloc) : ""
            });
        }

        const investasiOutDir = path.join(runDir, "investasi");
        fs.mkdirSync(investasiOutDir, { recursive: true });
        console.log(
            `\n  Menulis ${investasiMirrorSources.length} mirror berkas Investasi ke ${investasiOutDir}...`
        );
        for (const { fileName, localPath } of investasiMirrorSources) {
            const outName = path.basename(fileName);
            const outPath = path.join(investasiOutDir, outName);
            await writeInvestasiMirrorWithDendogram(
                localPath,
                outPath,
                matchByKey,
                lppMasterRows
            );
            console.log(`  [OK] Mirror investasi: ${outPath}`);
        }

        const sortedPerPksRows = [...investasiPerPksRows].sort(
            (a, b) =>
                a.namaPabrik.localeCompare(b.namaPabrik) || a.seqInPks - b.seqInPks
        );

        const perPksFinalRows = [];

        for (const row of sortedPerPksRows) {
            const match = row.matchKey ? matchByKey.get(row.matchKey) : null;
            const eqktuBefore = match?.eqktuBefore ?? "";
            const funcloc = match?.funcloc ?? "";
            const similarityScore = match?.similarityScore ?? "0%";
            const namaBaru = match?.namaBaru ?? "";

            perPksFinalRows.push([
                eqktuBefore,
                funcloc,
                row.namaMesin,
                row.alatNo,
                row.namaPabrik,
                similarityScore,
                namaBaru,
                row.namaPks
            ]);
        }

        console.log(
            `  Menulis ${perPksFinalRows.length} baris hasil ke match_per_pks.xlsx (${investasiCandidates.length} baris unik di match.xlsx)...`
        );

        const perPksWb = new ExcelJS.Workbook();
        const perPksWs = perPksWb.addWorksheet("Match Per PKS", {
            views: [{ state: "frozen", ySplit: 1 }]
        });

        const perPksHeaders = [...headers, "Regional"];
        perPksWs.addRow(perPksHeaders);
        perPksWs.getRow(1).font = { bold: true };

        for (const rowData of perPksFinalRows) {
            perPksWs.addRow(rowData);
        }

        perPksWs.columns = [
            { width: 45 },
            { width: 45 },
            { width: 45 },
            { width: 25 },
            { width: 30 },
            { width: 20 },
            { width: 45 },
            { width: 30 }
        ];

        if (perPksFinalRows.length > 0) {
            perPksWs.autoFilter = {
                from: { row: 1, column: 1 },
                to: { row: 1 + perPksFinalRows.length, column: 8 }
            };
        }

        appendMasterNamaBaruSheet(perPksWb, lppMasterRows);

        const perPksOutputPath = path.join(runDir, "match_per_pks.xlsx");
        await perPksWb.xlsx.writeFile(perPksOutputPath);
        console.log(`  [OK] Berhasil menyimpan berkas hasil per PKS: ${perPksOutputPath}`);
        console.log(
            `  [OK] Sheet master nama baru: ${lppMasterRows.length} baris (urutan sama dengan LPP).`
        );

        console.log(`\nSeluruh proses selesai dengan sukses!`);
        console.log(`Folder Output: ${runDir}`);

    } catch (err) {
        console.error(`\n[Error] Terjadi kesalahan fatal:`, err.message || err);
        process.exit(1);
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
