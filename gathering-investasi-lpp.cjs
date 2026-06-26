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
 * 2. investasi.xlsx (2 kolom data Investasi unik)
 * 3. match.xlsx (4 kolom hasil pencocokan & sisa data yang tidak cocok)
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

const { readSpreadsheetRows } = require("./utils/equipment_excel_io.cjs");
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
    findFunclocDescAfterLevel123ColumnIndex
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

/**
 * Mendeteksi kolom Nama Mesin/Peralatan dan Mesin/Alat No.
 */
function findInvestasiColumns(headerRow) {
    let idxNamaMesin = -1;
    let idxAlatNo = -1;
    let idxNamaPabrik = -1;
    for (let i = 0; i < headerRow.length; i++) {
        const h = String(headerRow[i] ?? "").toUpperCase().trim();
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
    }
    return { idxNamaMesin, idxAlatNo, idxNamaPabrik };
}

/**
 * Mencari layout header Investasi
 */
function findInvestasiLayout(rows) {
    const limit = Math.min(rows.length, 30);
    for (let i = 0; i < limit; i++) {
        const { idxNamaMesin, idxAlatNo, idxNamaPabrik } = findInvestasiColumns(rows[i]);
        if (idxNamaMesin !== -1) {
            return { headerRowIndex: i, idxNamaMesin, idxAlatNo, idxNamaPabrik };
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

            const readLpp = await readSpreadsheetRows(lppLocal);
            if (!readLpp.ok) {
                console.warn(`     [Skip] Gagal membaca berkas LPP: ${readLpp.reason}`);
                continue;
            }

            const lppRawRows = readLpp.rawRows;
            const lppLayout = findLppGatheringLayout(lppRawRows);
            const lppHeader = lppRawRows[lppLayout.headerRowIndex];
            if (!lppHeader) {
                console.warn("     [Skip] Header LPP tidak ditemukan.");
                continue;
            }

            const lppColIndex = buildColumnIndexFirstWins(lppHeader, normalizeHeader);
            const idxEqktuBefore = lppColIndex["EQKTU BEFORE"];
            const idxFunclocLevel123 = findFunclocDescAfterLevel123ColumnIndex(lppHeader);

            if (idxEqktuBefore === undefined || idxFunclocLevel123 === undefined) {
                console.warn("     [Skip] Kolom 'EQKTU BEFORE' atau 'FUNCTLOC DESC. AFTER LEVEL 1,2,3' tidak ditemukan.");
                continue;
            }

            const lppDataRows = lppRawRows.slice(lppLayout.headerRowIndex + 1);
            let fileExtracted = 0;

            for (const r of lppDataRows) {
                const eqktuBefore = getVal(r, idxEqktuBefore);
                const funcloc = getVal(r, idxFunclocLevel123);

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

        const investasiCandidates = [];
        const seenInvestasi = new Set();

        for (const file of investasiEntries) {
            console.log(`  Membaca: ${file.name}`);
            const { localFile: invLocal } = await downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                "inv_"
            );

            const sheets = await readAllSheetsRows(invLocal);
            let fileExtracted = 0;

            for (const sheet of sheets) {
                const layout = findInvestasiLayout(sheet.rows);
                if (!layout) continue;

                const { idxNamaMesin, idxAlatNo, idxNamaPabrik, headerRowIndex } = layout;
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

                    if (namaMesin) {
                        const cleanNama = cleanForMatching(namaMesin);
                        const cleanAlat = cleanForMatching(alatNo);
                        let fullStr = "";
                        if (cleanAlat && cleanNama.endsWith(cleanAlat)) {
                            fullStr = namaMesin;
                        } else {
                            fullStr = `${namaMesin} ${alatNo}`.trim();
                        }
                        const matchKey = cleanForMatching(fullStr);

                        // Deduplikasi Investasi (first wins)
                        if (matchKey && !seenInvestasi.has(matchKey)) {
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
            }
            console.log(`     Ekstrak ${fileExtracted} baris unik.`);
        }

        // Tulis berkas Investasi intermediate
        console.log(`\n  Menulis ${investasiCandidates.length} data Investasi unik ke investasi.xlsx...`);
        const invWb = new ExcelJS.Workbook();
        const invWs = invWb.addWorksheet("Investasi Data");
        invWs.addRow(["Nama Mesin/Peralatan", "Mesin/Alat No.", "Nama Pabrik"]);
        invWs.getRow(1).font = { bold: true };
        for (const c of investasiCandidates) {
            invWs.addRow([c.namaMesin, c.alatNo, c.namaPabrik]);
        }
        invWs.columns = [{ width: 45 }, { width: 25 }, { width: 30 }];
        const invOutputPath = path.join(runDir, "investasi.xlsx");
        await invWb.xlsx.writeFile(invOutputPath);
        console.log(`  [OK] Berhasil menyimpan berkas intermediate Investasi: ${invOutputPath}`);

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
