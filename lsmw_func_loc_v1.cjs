/**
 * lsmw_func_loc_v1.cjs
 * ─────────────────────────────────────────────────────────────────────────────
 * Pipeline konversi data PKS → format upload LSMW functional location (template Excel).
 *
 * Sumber: folder Drive (REGIONAL*), file .xlsx atau Google Sheets.
 * Template: template_create_func_loc_lsmw1.xlsx — isi mulai baris 5 (style header tetap).
 *
 * Filter
 *   - Semua baris dengan FUNCTIONAL LOCATION AFTER terisi ikut (stasiun, drive unit, accessories).
 *   - Duplikat FUNCTIONAL LOCATION AFTER lintas file tidak masuk output (primary key).
 *
 * Mapping kolom template (berurutan)
 *   STRNO  ← FUNCTIONAL LOCATION AFTER
 *   ALKEY  ← "1"
 *   TPLKZ  ← "PTPN"
 *   FLTYP  ← "M"
 *   PLTXT  ← FUNCTLOC DESC AFTER; alias mapping-alias-desc-func-loc.xlsx jika cocok
 *   SWERK  ← plant per file (MAINTENANCE PLAN AFTER pertama + fallback; sama semua baris)
 *   BUKRS  ← "PALM"
 *   KOSTL  ← funcloc level 4 (4 segmen): segmen 2 + STAS + 2 digit segmen 4 STRNO (mis. 1F11STAS01);
 *          ← baris lain: COST CENTER AFTER hanya jika flag --with-cost-center
 *   IWERK  ← sama dengan SWERK (satu nilai per file)
 *   GEWRK  ← utils/cwc.csv (menggantikan sumber); kolom regional dari maintenance plant (SWERK)
 *   INGRP  ← utils/planner-group.csv (menggantikan sumber); kolom regional sama
 *   WERGW  ← sama dengan SWERK (satu nilai per file)
 *   IEQUI  ← "X"
 *   TPLMA  ← parent funcloc (FUNCTIONAL LOCATION AFTER tanpa segmen terakhir)
 *
 * Baris sintetis di awal output (sebelum stasiun & level lain), mulai baris 5 template:
 *   Jika ada baris sumber FUNCTLOC DESC mengandung "STATION" dan parent stasiun belum ada
 *   di output file ini, tambah hingga 2 baris (urutan dari atas):
 *     1) STRNO = parent Area Pabrik, PLTXT = nama PKS dari nama file sumber
 *        (mis. "1.PKS AUR GADING.xlsx" → "PKS AUR GADING")
 *     2) STRNO = parent stasiun (TPLMA stasiun), PLTXT = "AREA PABRIK"
 *   GEWRK/INGRP kosong pada kedua baris sintetis.
 *
 * Sumber: folder Protect Y (prompt URL saat dijalankan, atau --protect-y-folder-link).
 *
 *   node lsmw_func_loc_v1.cjs
 *   node lsmw_func_loc_v1.cjs --protect-y-folder-link https://drive.google.com/drive/folders/...
 *   node lsmw_func_loc_v1.cjs --with-cost-center
 *   node lsmw_func_loc_v1.cjs --output-drive-folder-link https://drive.google.com/drive/folders/...
 *
 * Output lokal: Output/2. LSMW Create Functional Location V1/***.xlsx (nama sama dengan file sumber)
 * Output Drive:  hanya dengan --output-drive-folder-link
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const {
    TEMP_DOWNLOAD_PATH,
    XLSX_MIME,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    parseLsmwCli,
    resolveProtectYSourceFolderId,
    listRegionalFolders,
    promptRegionalFolderSelection,
    lsmwOutputFileName
} = require("./utils/lsmw_cli.cjs");
const {
    isEmpty,
    ensureCwcPlannerLookups,
    getEffectiveCostCenter,
    logRegionalMappingForFile,
    buildColumnIndexFirstWins,
    resolvePlannerGroup,
    resolveWorkCenter
} = require("./utils/lsmw_lookups.cjs");
const {
    ensureFunclocDescAliasMap,
    fixFunclocDescSourceTypos,
    applyFunclocDescAlias
} = require("./utils/funcloc_desc_alias.cjs");
const { findCostCenterColumnIndex } = require("./utils/equipment_gathering_columns.cjs");
const { regionalAggregateFileName } = require("./utils/lsmw_pks_aggregates.cjs");
const {
    mergeTemplateBodiesFromFiles,
    sortPksPaths
} = require("./utils/lsmw_func_loc_aggregate.cjs");
const { withRetry, isRetryableDriveError } = require("./utils/lsmw_retry.cjs");

const SCOPES = ["https://www.googleapis.com/auth/drive"];

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";

const LOCAL_OUTPUT_ROOT = path.join("Output", "2. LSMW Create Functional Location V1");
const TEMPLATE_PATH = "./template_create_func_loc_lsmw1.xlsx";
const COLUMN_JSON_PATH = path.join(LOCAL_OUTPUT_ROOT, "column_func_loc.json");

const OUTPUT_START_ROW = 5;

let drive;
let outputDriveFolderId = null;
let withCostCenter = false;

let secondRowDump = [];

/** Primary key lintas file: FUNCTIONAL LOCATION AFTER yang sudah masuk output. */
const usedFuncLocAfter = new Set();

function cellToJson(v) {
    if (v === undefined || v === null || v === "") return null;
    if (v instanceof Date) return v.toISOString();
    return v;
}

function normalizeHeader(name) {
    if (!name) return "";
    return name.toString().trim().toUpperCase();
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

function parentFuncLoc(code) {
    const key = normalizeFuncLocKey(code);
    if (!key || !key.includes("-")) return null;
    return key.substring(0, key.lastIndexOf("-"));
}

const AREA_PABRIK_PLTXT = "AREA PABRIK";
/** Stasiun = funcloc tepat 4 segmen hyphen, mis. PALM-4F04-0005-0001. */
const STATION_FUNCLOC_SEGMENT_COUNT = 4;

/** KOSTL baris stasiun: segmen #2 + STAS + 2 digit terakhir segmen #4 (mis. PALM-1F11-0005-0001 → 1F11STAS01). */
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

function getFuncLocSegmentCount(code) {
    const key = normalizeFuncLocKey(code);
    if (!key) return 0;
    return key.split("-").filter(Boolean).length;
}

function isStationLevelFuncLoc(funcLocKey) {
    return getFuncLocSegmentCount(funcLocKey) === STATION_FUNCLOC_SEGMENT_COUNT;
}

/** FUNCTLOC DESC mengandung kata STATION (bukan sekadar substring acak). */
function descContainsStation(desc) {
    return /\bSTATION\b/i.test(String(desc ?? "").trim());
}

/**
 * Parent funcloc (nilai TPLMA baris stasiun) yang belum punya baris STRNO sendiri di output file ini.
 * Maks. satu per file: ambil stasiun pertama yang parent-nya belum ada di outputKeysThisFile.
 */
function findAreaPabrikFuncLoc(dataRows, idxFuncLoc, funclocDescText, outputKeysThisFile) {
    for (const r of dataRows) {
        const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);
        if (!funcLocKey) continue;
        if (!descContainsStation(funclocDescText(r))) continue;
        const tplma = parentFuncLoc(funcLocKey);
        if (!tplma) continue;
        if (!outputKeysThisFile.has(tplma)) return tplma;
    }
    return null;
}

/** Nama PKS untuk PLTXT baris parent: "1.PKS AUR GADING.xlsx" → "PKS AUR GADING". */
function extractPksPltxtFromSourceName(sourceName) {
    let base = path.basename(String(sourceName ?? ""));
    base = base.replace(/\.xlsx$/i, "").trim();
    base = base.replace(/^\d+\.\s*/, "").trim();
    return base || "PKS";
}

function buildSyntheticFuncLocRow(funcLoc, pltxt, fileSwerk) {
    return [
        funcLoc,
        "1",
        "PTPN",
        "M",
        pltxt,
        fileSwerk,
        "PALM",
        null,
        fileSwerk,
        null,
        null,
        fileSwerk,
        "X",
        parentFuncLoc(funcLoc)
    ];
}

function buildAreaPabrikRow(funcLoc, fileSwerk) {
    return buildSyntheticFuncLocRow(funcLoc, AREA_PABRIK_PLTXT, fileSwerk);
}

function buildPksParentRow(pksFuncLoc, pltxt, fileSwerk) {
    return buildSyntheticFuncLocRow(pksFuncLoc, pltxt, fileSwerk);
}

function isValidPlantCode(code) {
    const s = String(code ?? "").trim().toUpperCase();
    return /^[A-Z0-9]{4}$/.test(s);
}

/** SWERK / IWERK / WERGW: satu nilai per file output. */
function resolveFileSwerk(dataRows, val, getValueWithParent, idxFuncLoc) {
    for (const r of dataRows) {
        const mp = val(r, "MAINTENANCE PLAN AFTER");
        if (!isEmpty(mp)) return mp;
    }

    for (const r of dataRows) {
        if (!normalizeFuncLocKey(r[idxFuncLoc])) continue;
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

    return scanPlant(dataRows);
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

    drive = google.drive({ version: "v3", auth });
}

async function listFolderChildrenWithRetry(folderId) {
    return withRetry(
        async () => {
            const out = [];
            let pageToken;
            do {
                const res = await drive.files.list({
                    q: `'${folderId}' in parents and trashed=false`,
                    fields: "nextPageToken, files(id,name,mimeType)",
                    pageSize: 1000,
                    pageToken
                });
                out.push(...(res.data.files ?? []));
                pageToken = res.data.nextPageToken;
            } while (pageToken);
            return out;
        },
        { label: `list folder ${folderId}`, retryIf: isRetryableDriveError }
    );
}

async function downloadSpreadsheetWithRetry(file) {
    return withRetry(
        () => downloadSpreadsheetToTemp(drive, file, TEMP_DOWNLOAD_PATH),
        { label: `unduh ${file.name}`, retryIf: isRetryableDriveError }
    );
}

async function createFolderIfNotExists(name, parentId) {
    const res = await drive.files.list({
        q: `name='${name.replace(/'/g, "\\'")}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
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

async function uploadFile(filePath, fileName, parentId) {
    const existing = await drive.files.list({
        q: `name='${fileName.replace(/'/g, "\\'")}' and '${parentId}' in parents and trashed=false`,
        fields: "files(id,name)"
    });

    if (existing.data.files.length > 0) {
        for (const f of existing.data.files) {
            await drive.files.delete({ fileId: f.id });
            console.log("Deleted old file:", f.name);
        }
    }

    await drive.files.create({
        resource: {
            name: fileName,
            parents: [parentId]
        },
        media: {
            mimeType: XLSX_MIME,
            body: fs.createReadStream(filePath)
        },
        fields: "id"
    });

    console.log("Uploaded:", fileName);
}

async function writeFuncLocOutputRows(outputRowArrays, outPath) {
    const workbookTemplate = new ExcelJS.Workbook();
    await workbookTemplate.xlsx.readFile(TEMPLATE_PATH);

    const sheetTemplate = workbookTemplate.worksheets[0];
    let startRow = OUTPUT_START_ROW;

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

async function buildFuncLocFile(sourcePath, sourceName, outputDir) {
    await ensureFunclocDescAliasMap();
    await ensureCwcPlannerLookups();

    const workbook = XLSX.readFile(sourcePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    secondRowDump.push({
        file: sourceName,
        excelRowNumber: 2,
        cells: rows[1] ? rows[1].map(cellToJson) : null
    });

    const headerRow = rows[1];

    if (!headerRow) {
        console.log("Header not found:", sourceName);
        return null;
    }

    const dataRows = rows.slice(2);

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);

    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxFunclocDesc = findFunclocDescAfterColumnIndex(headerRow);

    if (idxFuncLoc === undefined) {
        console.log("No FUNCTIONAL LOCATION AFTER column:", sourceName);
        return null;
    }

    if (idxFunclocDesc === undefined) {
        console.log("No FUNCTLOC DESC AFTER column:", sourceName);
    }

    function val(r, name) {
        const idx = colIndex[normalizeHeader(name)];
        return idx !== undefined ? r[idx] ?? null : null;
    }

    function funclocDescRaw(r) {
        if (idxFunclocDesc === undefined) return "";
        const v = r[idxFunclocDesc];
        return v !== null && v !== undefined && v !== ""
            ? fixFunclocDescSourceTypos(String(v))
            : "";
    }

    /** Deskripsi untuk template PLTXT (alias jika ada di mapping). */
    function funclocDescText(r) {
        return applyFunclocDescAlias(funclocDescRaw(r));
    }

    const idxCostCenter = findCostCenterColumnIndex(headerRow);

    const funcLocMap = {};
    for (const r of dataRows) {
        const code = normalizeFuncLocKey(r[idxFuncLoc]);
        if (code) funcLocMap[code] = r;
    }

    function resolveKostl(r, funcLocKey, fileSwerk) {
        if (isStationLevelFuncLoc(funcLocKey)) {
            return deriveKostlFromFuncLoc(funcLocKey);
        }
        if (!withCostCenter) return null;
        return getEffectiveCostCenter(r, funcLocKey, funcLocMap, idxCostCenter);
    }

    function getValueWithParent(r, columnName) {
        const idx = colIndex[normalizeHeader(columnName)];
        if (idx === undefined) return null;

        let value = r[idx];
        if (!isEmpty(value)) return value;

        let funcLoc = r[idxFuncLoc];
        while (funcLoc && funcLoc.includes("-")) {
            funcLoc = funcLoc.substring(0, funcLoc.lastIndexOf("-"));
            const parentRow = funcLocMap[funcLoc];
            if (!parentRow) continue;
            const parentVal = parentRow[idx];
            if (!isEmpty(parentVal)) return parentVal;
        }
        return null;
    }

    const fileSwerk = resolveFileSwerk(
        dataRows,
        val,
        getValueWithParent,
        idxFuncLoc
    );

    const hasOutputRows = dataRows.some(r =>
        normalizeFuncLocKey(r[idxFuncLoc])
    );
    if (hasOutputRows && isEmpty(fileSwerk)) {
        console.log("Skip (SWERK / plant tidak ditemukan):", sourceName);
        return null;
    }

    await logRegionalMappingForFile(fileSwerk, sourceName);

    const output = [];
    const outputKeysThisFile = new Set();

    for (const r of dataRows) {
        const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);
        if (!funcLocKey) continue;
        if (usedFuncLocAfter.has(funcLocKey)) continue;

        usedFuncLocAfter.add(funcLocKey);
        outputKeysThisFile.add(funcLocKey);

        const effectiveCostCenter = getEffectiveCostCenter(
            r,
            funcLocKey,
            funcLocMap,
            idxCostCenter
        );

        output.push([
            funcLocKey,
            "1",
            "PTPN",
            "M",
            funclocDescText(r),
            fileSwerk,
            "PALM",
            resolveKostl(r, funcLocKey, fileSwerk),
            fileSwerk,
            resolvePlannerGroup(r, funcLocKey, effectiveCostCenter, val, fileSwerk),
            resolveWorkCenter(r, funcLocKey, effectiveCostCenter, val, fileSwerk),
            fileSwerk,
            "X",
            parentFuncLoc(funcLocKey)
        ]);
    }

    const areaPabrikFuncLoc = findAreaPabrikFuncLoc(
        dataRows,
        idxFuncLoc,
        funclocDescRaw,
        outputKeysThisFile
    );

    if (areaPabrikFuncLoc) {
        if (!usedFuncLocAfter.has(areaPabrikFuncLoc)) {
            usedFuncLocAfter.add(areaPabrikFuncLoc);
        }
        output.unshift(buildAreaPabrikRow(areaPabrikFuncLoc, fileSwerk));

        const pksParentFuncLoc = parentFuncLoc(areaPabrikFuncLoc);
        const pksPltxt = extractPksPltxtFromSourceName(sourceName);
        if (
            pksParentFuncLoc &&
            pksParentFuncLoc !== areaPabrikFuncLoc &&
            !outputKeysThisFile.has(pksParentFuncLoc)
        ) {
            output.unshift(
                buildPksParentRow(pksParentFuncLoc, pksPltxt, fileSwerk)
            );
            if (!usedFuncLocAfter.has(pksParentFuncLoc)) {
                usedFuncLocAfter.add(pksParentFuncLoc);
            }
            console.log(
                `PKS parent row (PLTXT="${pksPltxt}", STRNO=${pksParentFuncLoc}):`,
                sourceName
            );
        }

        console.log(
            `Area Pabrik row (PLTXT="${AREA_PABRIK_PLTXT}", STRNO=${areaPabrikFuncLoc}):`,
            sourceName
        );
    }

    const newName = lsmwOutputFileName(sourceName);
    const newPath = path.join(outputDir, newName);

    await writeFuncLocOutputRows(output, newPath);

    return { path: newPath, name: newName };
}

async function processFolder(sourceFolderId, localDir, driveTargetId, pksOutputPaths) {
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

            await processFolder(file.id, childLocal, childDrive, pksOutputPaths);
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
        const result = await buildFuncLocFile(localFile, sourceName, localDir);

        if (!result) continue;

        if (pksOutputPaths) {
            pksOutputPaths.push(result.path);
        }

        console.log("Saved:", result.path);

        if (driveTargetId) {
            await uploadFile(result.path, result.name, driveTargetId);
        }
    }
}

async function main() {
    secondRowDump = [];
    usedFuncLocAfter.clear();

    const cli = parseLsmwCli(process.argv.slice(2));
    outputDriveFolderId = cli.outputDriveFolderId;
    withCostCenter = cli.withCostCenter;
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
        console.log(
            "KOSTL (cost center):",
            withCostCenter
                ? "enabled (--with-cost-center); level-4 funcloc KOSTL from STRNO (e.g. 1F11STAS01)"
                : "disabled except level-4 funcloc (STRNO segmen 2+4, e.g. 1F11STAS01)"
        );

        const regionalFolders = await withRetry(
            () => listRegionalFolders(drive, sourceRootFolderId),
            {
                label: `list regional folders ${sourceRootFolderId}`,
                retryIf: isRetryableDriveError
            }
        );

        if (regionalFolders.length === 0) {
            console.log("No REGIONAL folder found under source root");
        }

        const selectedRegionalNames =
            regionalFolders.length > 0
                ? await promptRegionalFolderSelection(regionalFolders)
                : new Set();
        if (selectedRegionalNames.size > 0) {
            console.log(
                `Regional diproses: ${[...selectedRegionalNames].join(", ")}`
            );
        }

        const regionalAggregatePaths = [];

        for (const folder of regionalFolders) {
            if (!selectedRegionalNames.has(folder.name.trim())) continue;

            console.log(`Processing ${folder.name}`);

            const localRegionalDir = path.join(LOCAL_OUTPUT_ROOT, folder.name);
            fs.mkdirSync(localRegionalDir, { recursive: true });

            let driveRegionalTarget = null;
            if (outputDriveFolderId) {
                driveRegionalTarget = await createFolderIfNotExists(
                    folder.name,
                    outputDriveFolderId
                );
            }

            const pksOutputPaths = [];
            await processFolder(
                folder.id,
                localRegionalDir,
                driveRegionalTarget,
                pksOutputPaths
            );

            if (pksOutputPaths.length > 0) {
                const regionalOut = path.join(
                    LOCAL_OUTPUT_ROOT,
                    regionalAggregateFileName(folder.name)
                );
                const { bodyRowCount } = await mergeTemplateBodiesFromFiles({
                    templatePath: TEMPLATE_PATH,
                    outputPath: regionalOut,
                    sourcePaths: sortPksPaths(pksOutputPaths)
                });
                console.log(
                    `Agregat regional: ${regionalOut} (${bodyRowCount} baris func loc)`
                );
                regionalAggregatePaths.push(regionalOut);
            }
        }

        if (outputDriveFolderId) {
            for (const aggPath of regionalAggregatePaths) {
                await uploadFile(
                    aggPath,
                    path.basename(aggPath),
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
        console.log(
            "Unique FUNCTIONAL LOCATION AFTER in output:",
            usedFuncLocAfter.size
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
