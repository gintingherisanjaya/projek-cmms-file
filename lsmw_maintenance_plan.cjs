/**
 * lsmw_maintenance_plan.cjs
 * Konversi Protect Y → template LSMW maintenance plan.
 *
 *   pnpm run lsmw-maintenance-plan
 *
 * Prompt: URL Protect Y, REGIONAL*, URL output measuring-document, mode validasi.
 * Output: Output/lsmw-maintenance-plan/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx + {REGIONAL N}.xlsx + all-pks.xlsx
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const { input, select, checkbox } = require("@inquirer/prompts");
const {
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    extractFolderIdFromUrl,
    isRegionalFolderName,
    listRegionalFolders,
    promptRegionalFolderSelection
} = require("./utils/lsmw_cli.cjs");
const { parseRegionalNumber, compareRegionalFileEntries } = require("./utils/pks_sort_keys.cjs");
const { readSpreadsheetRows } = require("./utils/equipment_excel_io.cjs");
const { findDataLayout, normalizeHeader } = require("./utils/lsmw_cell_fill.cjs");
const {
    buildColumnIndexFirstWins,
    ensureCwcPlannerLookups,
    getEffectiveCostCenter,
    resolveWorkCenter,
    resolvePlannerGroup
} = require("./utils/lsmw_lookups.cjs");
const { collectRedEquipGroupRowIndices } = require("./utils/lsmw_equip_group_red.cjs");
const { withRetry } = require("./utils/lsmw_retry.cjs");
const { buildMaintenancePlanRows } = require("./utils/lsmw_maintenance_plan_transform.cjs");
const {
    loadMaintenancePlanTemplateLimits,
    MAINTENANCE_PLAN_OUTPUT_START_ROW,
    DEFAULT_STRICT_FIELD_COLUMNS,
    formatCharViolation,
    formatFieldViolation
} = require("./utils/lsmw_maintenance_plan_limits.cjs");
const {
    writeMaintenancePlanValidationExcel
} = require("./utils/lsmw_maintenance_plan_validation_excel.cjs");
const {
    loadMaintenanceItemColumnOrder,
    writeMaintenanceItemExcel
} = require("./utils/lsmw_maintenance_item_excel.cjs");
const {
    DEFAULT_MAPPING_PATH,
    loadEquipmentNumberMapping
} = require("./utils/equipment_number_mapping.cjs");
const {
    DEFAULT_IP18_PATH,
    loadMaintItemsByEquipment
} = require("./utils/maintenance_item_ip18_mapping.cjs");
const {
    DEFAULT_IK07_PATH,
    loadMeasuringPointByEquipment
} = require("./utils/measuring_point_ik07_mapping.cjs");
const { findCostCenterColumnIndex } = require("./utils/equipment_gathering_columns.cjs");
const { writeRegionalAndAllPksAggregates } = require("./utils/lsmw_pks_aggregates.cjs");
const { loadHeriSheet } = require("./utils/heri_sheet_loader.cjs");
const { ensureFunclocDescAliasMap } = require("./utils/funcloc_desc_alias.cjs");
const {
    loadMeasuringDocumentCounterFromDrive
} = require("./utils/measuring_document_counter_loader.cjs");

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly"
];
const OUTPUT_DIR = path.join("Output", "lsmw-maintenance-plan");
const TEMPLATE_PATH = "./template_lsmw_maintenance_plan.xlsx";
let drive;
let sheets;

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

function resolveFileBegru(dataRows, val, idxFuncLoc) {
    for (const r of dataRows) {
        const mp = val(r, "MAINTENANCE PLAN AFTER");
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

function stopWithError(message) {
    console.error(`\n  [error] ${message}`);
    console.error("Proses dihentikan.");
    process.exit(1);
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

function pksXlsxFileName(sourceFileName) {
    const base = path
        .basename(String(sourceFileName ?? ""))
        .replace(/\.xlsx$/i, "")
        .trim();
    return `${base.replace(/[/\\?%*:|"<>]/g, "_")}.xlsx`;
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
    sheets = google.sheets({ version: "v4", auth });
}

async function listFolderChildren(folderId) {
    return withRetry(
        async () => {
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
        },
        { label: `list folder ${folderId}` }
    );
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

async function listRegionalSpreadsheets(rootFolderId, selectedRegionalNames) {
    const entries = [];
    const rootChildren = await listFolderChildren(rootFolderId);

    for (const child of rootChildren) {
        if (child.mimeType !== "application/vnd.google-apps.folder") continue;
        if (!isRegionalFolderName(child.name)) continue;

        const regionalName = child.name.trim();
        if (!selectedRegionalNames.has(regionalName)) continue;

        const files = await listSpreadsheetsRecursive(child.id);
        for (const file of files) {
            entries.push({ regional: regionalName, file });
        }
    }

    entries.sort(compareRegionalFileEntries);

    return entries;
}

async function promptMeasuringDocumentFolderUrl() {
    return input({
        message:
            "URL folder Google Drive — output lsmw-measuring-document (folder run, berisi REGIONAL N/*.xlsx):",
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

async function promptProtectYFolderUrl() {
    return input({
        message:
            "URL folder Google Drive — Protect Y (subfolder REGIONAL*):",
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

async function promptValidationMode() {
    return select({
        message:
            "Validasi kolom wajib (EQUNR, MAINTENANCE_ITEM, SZAEH, WPTXT, POINT, Counter Reading) + batas karakter:",
        choices: [
            {
                name: "Strict — hentikan proses (batas karakter selalu strict)",
                value: "strict"
            },
            {
                name: "Warning — lanjutkan + validation.xlsx",
                value: "warning"
            }
        ]
    });
}

async function promptStrictFieldColumns() {
    const selected = await checkbox({
        message:
            "Kolom wajib yang tetap strict (terminate). Uncheck = hanya warning + validation.xlsx:",
        choices: DEFAULT_STRICT_FIELD_COLUMNS.map(col => ({
            name: col,
            value: col,
            checked: true
        })),
        validate: () => true
    });
    return new Set(selected);
}

async function downloadSpreadsheetWithRetry(file) {
    return withRetry(
        () =>
            downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                "mp_"
            ),
        { label: `unduh ${file.name}` }
    );
}

async function readSpreadsheetRowsWithRetry(localFile) {
    return withRetry(() => readSpreadsheetRows(localFile), {
        label: `baca ${path.basename(localFile)}`
    });
}

async function writeValidationExcelIfNeeded(runDir, validationDetailRows) {
    if (!runDir || validationDetailRows.length === 0) return null;
    const validationPath = path.join(runDir, "validation.xlsx");
    await writeMaintenancePlanValidationExcel(
        validationPath,
        validationDetailRows
    );
    return validationPath;
}

async function terminateWithValidation(message, runDir, validationDetailRows) {
    const validationPath = await writeValidationExcelIfNeeded(
        runDir,
        validationDetailRows
    );
    if (validationPath) {
        console.log(
            `Validasi: ${validationPath} (${validationDetailRows.length} masalah)`
        );
    }
    stopWithError(message);
}

async function processValidationResults(built, ctx) {
    const { validationMode, strictFieldColumns, validationDetailRows } = ctx;

    if (built.violations.length > 0) {
        if (validationMode === "strict") {
            await terminateWithValidation(
                formatCharViolation(built.violations[0]),
                ctx.runDir,
                validationDetailRows
            );
        }
        for (const v of built.violations) {
            console.warn(`  [warn] ${formatCharViolation(v)}`);
        }
    }

    for (const issue of built.fieldIssues ?? []) {
        const isStrictColumn =
            validationMode === "strict" &&
            strictFieldColumns.has(issue.column);
        if (isStrictColumn) {
            await terminateWithValidation(
                formatFieldViolation(issue),
                ctx.runDir,
                validationDetailRows
            );
        }
        console.warn(`  [warn] ${formatFieldViolation(issue)}`);
        validationDetailRows.push(issue);
    }

    for (const t of built.maintItemExtras ?? []) {
        console.warn(
            `  [warn] MAINTENANCE_ITEM >1 — ${t.fileName} baris ${t.sourceExcelRow} FUNCLOC=${t.tplnr} (pakai yang pertama)`
        );
    }
}

async function main() {
    const stamp = wibTimestampForFilename();
    const runDir = path.join(OUTPUT_DIR, stamp);

    if (!fs.existsSync(TEMPLATE_PATH)) {
        stopWithError(`Template tidak ditemukan: ${TEMPLATE_PATH}`);
    }
    if (!fs.existsSync(DEFAULT_IP18_PATH)) {
        stopWithError(`File mapping tidak ditemukan: ${DEFAULT_IP18_PATH}`);
    }
    if (!fs.existsSync(DEFAULT_IK07_PATH)) {
        stopWithError(`File mapping tidak ditemukan: ${DEFAULT_IK07_PATH}`);
    }

    prepareTempDownloads();

    const validationDetailRows = [];
    let strictFieldColumns = new Set(DEFAULT_STRICT_FIELD_COLUMNS);

    try {
        await initDrive();

        console.log("Memuat sheet HERI dari Google Spreadsheet...");
        const heriData = await loadHeriSheet(sheets);
        console.log(
            `  ${heriData.equipmentNames.length} kolom alat, ${heriData.byPlant.size} plant`
        );

        const protectYFolderUrl = await promptProtectYFolderUrl();
        const protectYRootFolderId = extractFolderIdFromUrl(
            String(protectYFolderUrl).trim()
        );

        const regionalFolders = await listRegionalFolders(
            drive,
            protectYRootFolderId
        );
        if (regionalFolders.length === 0) {
            stopWithError("Tidak ada subfolder REGIONAL* di folder Protect Y.");
        }

        const selectedRegionalNames =
            await promptRegionalFolderSelection(regionalFolders);

        const measuringDocFolderUrl = await promptMeasuringDocumentFolderUrl();
        const measuringDocRootFolderId = extractFolderIdFromUrl(
            String(measuringDocFolderUrl).trim()
        );

        console.log("Memuat Counter Reading dari output measuring-document...");
        const measuringDocIndex = await loadMeasuringDocumentCounterFromDrive({
            drive,
            rootFolderId: measuringDocRootFolderId,
            isProcessableSpreadsheet,
            listFolderChildren,
            downloadSpreadsheetToTemp: async file => {
                const { localFile } = await downloadSpreadsheetWithRetry(file);
                return localFile;
            },
            readSpreadsheetRows: async file => {
                const sheet = await readSpreadsheetRowsWithRetry(file);
                return sheet.ok ? sheet : { rawRows: [] };
            }
        });
        console.log(
            `[measuring-doc-index] ${measuringDocIndex.counterByEqunr.size} equipment` +
                ` (${measuringDocIndex.filesScanned} file di-scan)`
        );

        const validationMode = await promptValidationMode();
        if (validationMode === "strict") {
            strictFieldColumns = await promptStrictFieldColumns();
            console.log(
                `  Kolom strict: ${[...strictFieldColumns].join(", ") || "(tidak ada)"}`
            );
        }

        console.log("\nMemuat batas karakter dari template maintenance plan...");
        const templateLimits = await loadMaintenancePlanTemplateLimits(
            path.resolve(TEMPLATE_PATH)
        );
        const columnOrder =
            templateLimits.columns.length > 0
                ? templateLimits.columns
                : await loadMaintenanceItemColumnOrder(path.resolve(TEMPLATE_PATH));

        const equipmentMapping = loadEquipmentNumberMapping(
            path.resolve(DEFAULT_MAPPING_PATH)
        );
        console.log(
            `[mapping EQUNR] ${equipmentMapping.byPlantDesc.size} key aktif` +
                ` (${equipmentMapping.duplicateKeys} duplikat diabaikan)`
        );

        const ip18Mapping = loadMaintItemsByEquipment(
            path.resolve(DEFAULT_IP18_PATH)
        );
        console.log(
            `[mapping MAINTENANCE_ITEM] ${ip18Mapping.byEquipment.size} equipment, ${ip18Mapping.totalRows} baris ip18`
        );

        const ik07Mapping = loadMeasuringPointByEquipment(
            path.resolve(DEFAULT_IK07_PATH)
        );
        console.log(
            `[mapping POINT] ${ik07Mapping.byEquipment.size} equipment` +
                ` (${ik07Mapping.duplicateEquipment} duplikat diabaikan, first wins)`
        );

        await ensureFunclocDescAliasMap();
        await ensureCwcPlannerLookups();

        console.log("Memuat file Protect Y (subfolder REGIONAL*)...");
        const entries = await listRegionalSpreadsheets(
            protectYRootFolderId,
            selectedRegionalNames
        );
        const regionalSummary = [...selectedRegionalNames]
            .sort((a, b) => parseRegionalNumber(a) - parseRegionalNumber(b))
            .join(", ");
        console.log(
            `  Regional diproses: ${regionalSummary} (${entries.length} spreadsheet)`
        );

        if (entries.length === 0) {
            console.log("Tidak ada file untuk diproses.");
            return;
        }

        fs.mkdirSync(runDir, { recursive: true });

        let filesOk = 0;
        let filesSkipped = 0;
        let totalRows = 0;
        let totalRedEquipGroupSkipped = 0;
        let totalCharWarnings = 0;
        let totalFieldWarnings = 0;
        let totalMaintItemExtras = 0;

        const seenFuncLocKeys = new Set();
        const regionalRowsByName = new Map();
        const allRows = [];

        for (const { regional, file } of entries) {
            const regionalDir = path.join(runDir, regional);
            fs.mkdirSync(regionalDir, { recursive: true });
            const outPath = path.join(regionalDir, pksXlsxFileName(file.name));

            console.log(`[${regional}] ${file.name}`);

            const { localFile } = await downloadSpreadsheetWithRetry(file);

            const sheet = await readSpreadsheetRowsWithRetry(localFile);
            if (!sheet.ok) {
                console.log(`  [skip] ${sheet.reason}`);
                filesSkipped += 1;
                continue;
            }

            const layout = findDataLayout(sheet.rawRows);
            const headerRow = sheet.rawRows[layout.headerRowIndex];
            const colIndex = headerRow
                ? buildColumnIndexFirstWins(headerRow, normalizeHeader)
                : {};
            const idxEquipGroupAfter =
                colIndex["EQUIPMENT GROUP AFTER"] ?? layout.idxEquipGroupAfter;
            const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
            const idxCostCenter = findCostCenterColumnIndex(headerRow);

            if (idxEquipGroupAfter === undefined) {
                console.warn(
                    `  [red-skip] kolom EQUIPMENT GROUP AFTER tidak ditemukan — ${file.name}`
                );
            }

            const dataRowCount =
                sheet.rawRows.length - layout.headerRowIndex - 1;
            const redEquipGroupRowIndices =
                await collectRedEquipGroupRowIndices({
                    sourcePath: localFile,
                    colIndex0: idxEquipGroupAfter,
                    dataRowCount,
                    dataStartExcelRow: layout.dataStartExcelRow,
                    driveFile: file,
                    sheetsApi: sheets
                });

            const dataRows = sheet.rawRows.slice(layout.headerRowIndex + 1);
            const funcLocMap = {};
            if (idxFuncLoc !== undefined) {
                for (const r of dataRows) {
                    const code = normalizeFuncLocKey(r[idxFuncLoc]);
                    if (code) funcLocMap[code] = r;
                }
            }

            function val(r, name) {
                const idx = colIndex[normalizeHeader(name)];
                return idx !== undefined ? r[idx] ?? null : null;
            }

            const fileBegru = resolveFileBegru(dataRows, val, idxFuncLoc);
            const workCenterByRowIndex = new Map();
            const plannerGroupByRowIndex = new Map();
            const effectiveCostCenterByRowIndex = new Map();

            if (idxFuncLoc !== undefined) {
                for (let i = 0; i < dataRows.length; i++) {
                    const r = dataRows[i] || [];
                    const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);
                    if (!funcLocKey) continue;

                    const effectiveCostCenter = getEffectiveCostCenter(
                        r,
                        funcLocKey,
                        funcLocMap,
                        idxCostCenter
                    );
                    effectiveCostCenterByRowIndex.set(i, effectiveCostCenter);
                    workCenterByRowIndex.set(
                        i,
                        resolveWorkCenter(
                            r,
                            funcLocKey,
                            effectiveCostCenter,
                            val,
                            fileBegru
                        )
                    );
                    plannerGroupByRowIndex.set(
                        i,
                        resolvePlannerGroup(
                            r,
                            funcLocKey,
                            effectiveCostCenter,
                            val,
                            fileBegru
                        )
                    );
                }
            }

            const redSkipCount = redEquipGroupRowIndices.size;
            totalRedEquipGroupSkipped += redSkipCount;
            if (redSkipCount > 0) {
                console.log(
                    `  [red-skip] ${redSkipCount} baris (EQUIPMENT GROUP AFTER merah) — ${file.name}`
                );
            }

            const built = buildMaintenancePlanRows(
                sheet.rawRows,
                file.name,
                templateLimits.limitByTemplate,
                {
                    redEquipGroupRowIndices,
                    seenFuncLocKeys,
                    effectiveCostCenterByRowIndex,
                    workCenterByRowIndex,
                    plannerGroupByRowIndex,
                    equipmentByPlantDesc: equipmentMapping.byPlantDesc,
                    maintByEquipment: ip18Mapping.byEquipment,
                    pointByEquipment: ik07Mapping.byEquipment,
                    heriData,
                    counterByEqunr: measuringDocIndex.counterByEqunr,
                    regional
                }
            );

            if (!built.ok) {
                console.log(`  [skip] ${built.reason}`);
                filesSkipped += 1;
                continue;
            }

            if (built.rows.length === 0) {
                console.log(`  [skip] tidak ada baris maintenance plan yang cocok`);
                filesSkipped += 1;
                continue;
            }

            const rowsBeforeValidation = validationDetailRows.length;
            await processValidationResults(built, {
                validationMode,
                strictFieldColumns,
                validationDetailRows,
                runDir
            });
            totalCharWarnings += built.violations.length;
            totalFieldWarnings +=
                validationDetailRows.length - rowsBeforeValidation;
            totalMaintItemExtras += built.maintItemExtras.length;

            for (const w of built.warnings ?? []) {
                console.warn(`  [warn] ${w}`);
            }

            await writeMaintenanceItemExcel(
                outPath,
                path.resolve(TEMPLATE_PATH),
                built.rows,
                columnOrder,
                MAINTENANCE_PLAN_OUTPUT_START_ROW
            );

            if (!regionalRowsByName.has(regional)) {
                regionalRowsByName.set(regional, []);
            }
            regionalRowsByName.get(regional).push(...built.rows);
            allRows.push(...built.rows);

            console.log(`  → ${built.rows.length} baris → ${outPath}`);
            filesOk += 1;
            totalRows += built.rows.length;
        }

        const templateResolved = path.resolve(TEMPLATE_PATH);
        await writeRegionalAndAllPksAggregates({
            runDir,
            regionalRowsByName,
            allRows,
            label: "baris maintenance plan",
            writeRows: (outPath, rows) =>
                writeMaintenanceItemExcel(
                    outPath,
                    templateResolved,
                    rows,
                    columnOrder,
                    MAINTENANCE_PLAN_OUTPUT_START_ROW
                )
        });

        const validationPath = await writeValidationExcelIfNeeded(
            runDir,
            validationDetailRows
        );
        if (validationPath) {
            console.log(
                `\nValidasi: ${validationPath} (${validationDetailRows.length} masalah)`
            );
        }

        console.log("\nRingkasan:");
        console.log(`  Output: ${path.resolve(runDir)}`);
        console.log(`  Mode validasi: ${validationMode}`);
        if (validationMode === "strict") {
            console.log(
                `  Kolom wajib strict: ${[...strictFieldColumns].join(", ") || "(tidak ada)"}`
            );
        }
        console.log(`  File OK: ${filesOk}, dilewati: ${filesSkipped}`);
        console.log(`  Total baris maintenance plan: ${totalRows}`);
        console.log(
            "  Total baris di-skip (EQUIPMENT GROUP AFTER merah):",
            totalRedEquipGroupSkipped
        );
        if (totalCharWarnings > 0) {
            console.log(`  Peringatan batas karakter: ${totalCharWarnings}`);
        }
        if (totalFieldWarnings > 0) {
            console.log(`  Peringatan kolom wajib: ${totalFieldWarnings}`);
        }
        if (totalMaintItemExtras > 0) {
            console.log(
                `  Peringatan MAINTENANCE_ITEM >1 (pakai pertama): ${totalMaintItemExtras}`
            );
        }
        if (allRows.length === 0) {
            console.log("  all-pks.xlsx tidak dibuat (tidak ada data).");
        }
    } finally {
        cleanupTempDownloads();
    }
}

if (require.main === module) {
    main().catch(err => {
        console.error(err);
        process.exit(1);
    });
}
