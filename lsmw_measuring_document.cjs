/**
 * lsmw_measuring_document.cjs
 * Konversi Protect Y + sheet HERI → template LSMW measuring document.
 *
 *   pnpm run lsmw-measuring-document
 *
 * Prompt: URL Protect Y, pilih subfolder REGIONAL* (checkbox), mode validasi strict/warning + validation.xlsx.
 * Output: Output/lsmw-measuring-document/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx + {REGIONAL N}.xlsx
 *   + Missing Measuring Document/REGIONAL N/{nama PKS}.xlsx + {REGIONAL N}.xlsx + all-pks.xlsx
 *   + validation.xlsx
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
    getEffectiveCostCenter
} = require("./utils/lsmw_lookups.cjs");
const { collectRedEquipGroupRowIndices } = require("./utils/lsmw_equip_group_red.cjs");
const {
    DEFAULT_MAPPING_PATH,
    loadEquipmentNumberMapping
} = require("./utils/equipment_number_mapping.cjs");
const {
    DEFAULT_IK07_PATH,
    loadMeasuringPointByEquipment
} = require("./utils/measuring_point_ik07_mapping.cjs");
const { findCostCenterColumnIndex } = require("./utils/equipment_gathering_columns.cjs");
const { ensureFunclocDescAliasMap } = require("./utils/funcloc_desc_alias.cjs");
const {
    writeRegionalAndAllPksAggregates,
    ALL_PKS_FILE_NAME,
    regionalAggregateFileName
} = require("./utils/lsmw_pks_aggregates.cjs");
const { loadHeriSheet } = require("./utils/heri_sheet_loader.cjs");
const {
    buildMeasuringDocumentRows,
    renumberMeasuringDocumentSeqNo
} = require("./utils/lsmw_measuring_document_transform.cjs");
const {
    loadMeasuringDocumentColumnOrder,
    writeMeasuringDocumentExcel
} = require("./utils/lsmw_measuring_document_excel.cjs");
const {
    DEFAULT_STRICT_FIELD_COLUMNS,
    formatFieldViolation
} = require("./utils/lsmw_measuring_document_limits.cjs");
const {
    writeMeasuringDocumentValidationExcel
} = require("./utils/lsmw_measuring_document_validation_excel.cjs");
const {
    loadMeasuringDocumentCounterFromDrive
} = require("./utils/measuring_document_counter_loader.cjs");

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly"
];
const OUTPUT_DIR = path.join("Output", "lsmw-measuring-document");
const TEMPLATE_PATH = "./template_measuring_document.xlsx";
const MISSING_MEASURING_DOC_DIR = "Missing Measuring Document";

let drive;
let sheets;

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
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

/**
 * @param {Set<string>} selectedRegionalNames
 * @returns {Array<{ regional: string, file: object }>}
 */
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

function findColumnByPrefix(columnOrder, prefix) {
    const upper = prefix.toUpperCase();
    return (
        columnOrder.find(c => c.toUpperCase().startsWith(upper)) ??
        columnOrder.find(c => c.toUpperCase().includes(upper)) ??
        null
    );
}

async function promptPriorMeasuringDocumentFolderUrl() {
    return input({
        message:
            "URL folder Google Drive — output lsmw-measuring-document sebelumnya (untuk missing.xlsx):",
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
            "Validasi kolom wajib (Counter Reading, Equipment Number, Measuring point):",
        choices: [
            {
                name: "Strict — hentikan proses sesuai kolom tercentang",
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

async function writeValidationExcelIfNeeded(runDir, validationDetailRows) {
    if (!runDir || validationDetailRows.length === 0) return null;
    const validationPath = path.join(runDir, "validation.xlsx");
    await writeMeasuringDocumentValidationExcel(
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

    for (const issue of built.fieldIssues ?? []) {
        const isStrictColumn =
            validationMode === "strict" &&
            strictFieldColumns.has(issue.column);
        if (isStrictColumn) {
            await terminateWithValidation(
                formatFieldViolation(issue),
                ctx.runDir,
                [...validationDetailRows, issue]
            );
        }
        console.warn(`  [warn] ${formatFieldViolation(issue)}`);
        validationDetailRows.push(issue);
    }
}

async function withRetry(fn, { label = "operasi", maxAttempts = 4 } = {}) {
    let lastErr;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        try {
            return await fn();
        } catch (err) {
            lastErr = err;
            if (attempt < maxAttempts) {
                console.warn(
                    `  [retry ${attempt}/${maxAttempts - 1}] ${label}: ${err.message}`
                );
                await new Promise(r => setTimeout(r, 1000));
            }
        }
    }
    throw lastErr;
}

async function downloadSpreadsheetWithRetry(file) {
    return withRetry(
        () =>
            downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                "md_"
            ),
        { label: `unduh ${file.name}` }
    );
}

async function readSpreadsheetRowsWithRetry(localFile) {
    return withRetry(() => readSpreadsheetRows(localFile), {
        label: `baca ${path.basename(localFile)}`
    });
}

async function main() {
    const stamp = wibTimestampForFilename();
    const runDir = path.join(OUTPUT_DIR, stamp);

    if (!fs.existsSync(TEMPLATE_PATH)) {
        stopWithError(`Template tidak ditemukan: ${TEMPLATE_PATH}`);
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
        const validationMode = await promptValidationMode();
        if (validationMode === "strict") {
            strictFieldColumns = await promptStrictFieldColumns();
            console.log(
                `  Kolom strict: ${[...strictFieldColumns].join(", ") || "(tidak ada)"}`
            );
        }

        const templateResolved = path.resolve(TEMPLATE_PATH);
        const columnOrder = await loadMeasuringDocumentColumnOrder(
            templateResolved
        );

        const priorMeasuringDocFolderUrl =
            await promptPriorMeasuringDocumentFolderUrl();
        const priorMeasuringDocRootFolderId = extractFolderIdFromUrl(
            String(priorMeasuringDocFolderUrl).trim()
        );

        console.log(
            "Memuat EQUNR dari output measuring-document sebelumnya..."
        );
        const priorMeasuringDocIndex = await loadMeasuringDocumentCounterFromDrive(
            {
                drive,
                rootFolderId: priorMeasuringDocRootFolderId,
                isProcessableSpreadsheet,
                listFolderChildren,
                downloadSpreadsheetToTemp: async file => {
                    const { localFile } = await downloadSpreadsheetWithRetry(
                        file
                    );
                    return localFile;
                },
                readSpreadsheetRows: async file => {
                    const sheet = await readSpreadsheetRowsWithRetry(file);
                    return sheet.ok ? sheet : { rawRows: [] };
                }
            }
        );
        const priorEqunrSet = new Set(
            priorMeasuringDocIndex.counterByEqunr.keys()
        );
        console.log(
            `[prior-md-index] ${priorEqunrSet.size} EQUNR` +
                ` (${priorMeasuringDocIndex.filesScanned} file di-scan)`
        );

        const equipmentMapping = loadEquipmentNumberMapping(
            path.resolve(DEFAULT_MAPPING_PATH)
        );
        console.log(
            `[mapping EQUNR] ${equipmentMapping.byPlantDesc.size} key aktif` +
                ` (${equipmentMapping.duplicateKeys} duplikat diabaikan)`
        );

        const ik07Mapping = loadMeasuringPointByEquipment(
            path.resolve(DEFAULT_IK07_PATH)
        );
        console.log(
            `[mapping POINT] ${ik07Mapping.byEquipment.size} equipment` +
                ` (${ik07Mapping.duplicateEquipment} duplikat diabaikan)`
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
        let filesEmpty = 0;
        let totalRows = 0;
        let totalWarnings = 0;
        let totalRedEquipGroupSkipped = 0;
        let totalEmptyPoint = 0;
        let totalEmptyCounter = 0;

        const seenFuncLocKeys = new Set();
        const usedPlantEquipmentKeys = new Set();
        const regionalRowsByName = new Map();
        const regionalMissingRowsByName = new Map();
        const equnrCol = findColumnByPrefix(columnOrder, "Equipment Number");

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

            const effectiveCostCenterByRowIndex = new Map();
            if (idxFuncLoc !== undefined) {
                for (let i = 0; i < dataRows.length; i += 1) {
                    const r = dataRows[i] || [];
                    const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);
                    if (!funcLocKey) continue;
                    effectiveCostCenterByRowIndex.set(
                        i,
                        getEffectiveCostCenter(
                            r,
                            funcLocKey,
                            funcLocMap,
                            idxCostCenter
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

            const built = buildMeasuringDocumentRows(
                sheet.rawRows,
                file.name,
                columnOrder,
                {
                    heriData,
                    regional,
                    redEquipGroupRowIndices,
                    seenFuncLocKeys,
                    effectiveCostCenterByRowIndex,
                    equipmentByPlantDesc: equipmentMapping.byPlantDesc,
                    pointByEquipment: ik07Mapping.byEquipment,
                    usedPlantEquipmentKeys
                }
            );

            if (!built.ok) {
                console.log(`  [skip] ${built.reason}`);
                filesSkipped += 1;
                continue;
            }

            await processValidationResults(built, {
                validationMode,
                strictFieldColumns,
                validationDetailRows,
                runDir
            });

            for (const w of built.warnings) {
                console.warn(`  [warn] ${w}`);
            }
            totalWarnings += built.warnings.length;
            totalEmptyPoint += built.emptyPointCount;
            totalEmptyCounter += built.emptyCounterCount ?? 0;

            if (built.rows.length === 0) {
                console.log(`  → 0 baris (tidak ada match HERI) — dilewati`);
                filesEmpty += 1;
                continue;
            }

            await writeMeasuringDocumentExcel(
                outPath,
                templateResolved,
                built.rows,
                columnOrder
            );

            if (!regionalRowsByName.has(regional)) {
                regionalRowsByName.set(regional, []);
            }
            regionalRowsByName.get(regional).push(...built.rows);

            if (equnrCol) {
                const missingRows = built.rows.filter(row => {
                    const equnr = String(row[equnrCol] ?? "").trim();
                    return equnr && !priorEqunrSet.has(equnr);
                });
                if (missingRows.length > 0) {
                    const missingRegionalDir = path.join(
                        runDir,
                        MISSING_MEASURING_DOC_DIR,
                        regional
                    );
                    fs.mkdirSync(missingRegionalDir, { recursive: true });
                    const missingPksPath = path.join(
                        missingRegionalDir,
                        pksXlsxFileName(file.name)
                    );
                    await writeMeasuringDocumentExcel(
                        missingPksPath,
                        templateResolved,
                        missingRows,
                        columnOrder
                    );
                    console.log(
                        `  ${MISSING_MEASURING_DOC_DIR}/${regional}/${pksXlsxFileName(file.name)}: ${missingRows.length} baris`
                    );

                    if (!regionalMissingRowsByName.has(regional)) {
                        regionalMissingRowsByName.set(regional, []);
                    }
                    regionalMissingRowsByName.get(regional).push(...missingRows);
                }
            }

            console.log(`  → ${built.rows.length} baris → ${outPath}`);
            filesOk += 1;
            totalRows += built.rows.length;
        }

        await writeRegionalAndAllPksAggregates({
            runDir,
            regionalRowsByName,
            allRows: [],
            label: "baris measuring document",
            writeRows: (outPath, rows) =>
                writeMeasuringDocumentExcel(
                    outPath,
                    templateResolved,
                    renumberMeasuringDocumentSeqNo(rows, columnOrder),
                    columnOrder
                )
        });

        const missingDir = path.join(runDir, MISSING_MEASURING_DOC_DIR);
        const allMissingRows = [];
        let missingRegionalFiles = 0;

        for (const [regional, rows] of regionalMissingRowsByName) {
            if (!rows.length) continue;
            fs.mkdirSync(missingDir, { recursive: true });
            const outPath = path.join(
                missingDir,
                regionalAggregateFileName(regional)
            );
            await writeMeasuringDocumentExcel(
                outPath,
                templateResolved,
                renumberMeasuringDocumentSeqNo(rows, columnOrder),
                columnOrder
            );
            allMissingRows.push(...rows);
            missingRegionalFiles += 1;
            console.log(
                `  ${MISSING_MEASURING_DOC_DIR}/${regionalAggregateFileName(regional)}: ${rows.length} baris`
            );
        }

        if (allMissingRows.length > 0) {
            const allMissingPath = path.join(missingDir, ALL_PKS_FILE_NAME);
            await writeMeasuringDocumentExcel(
                allMissingPath,
                templateResolved,
                renumberMeasuringDocumentSeqNo(allMissingRows, columnOrder),
                columnOrder
            );
            console.log(
                `  ${MISSING_MEASURING_DOC_DIR}/${ALL_PKS_FILE_NAME}: ${allMissingRows.length} baris (${missingRegionalFiles} regional)`
            );
        } else if (totalRows > 0) {
            console.log(
                `  ${MISSING_MEASURING_DOC_DIR}/ tidak dibuat (semua EQUNR run ini sudah ada di folder Drive sebelumnya).`
            );
        }

        const validationPath = await writeValidationExcelIfNeeded(
            runDir,
            validationDetailRows
        );

        console.log("\nRingkasan:");
        console.log(`  Output: ${path.resolve(runDir)}`);
        console.log(
            `  File OK: ${filesOk}, kosong: ${filesEmpty}, dilewati: ${filesSkipped}`
        );
        console.log(`  Total baris measuring document: ${totalRows}`);
        console.log(
            "  Total baris di-skip (EQUIPMENT GROUP AFTER merah):",
            totalRedEquipGroupSkipped
        );
        if (totalEmptyPoint > 0) {
            console.log(`  Baris tanpa measuring point (ik07): ${totalEmptyPoint}`);
        }
        if (totalEmptyCounter > 0) {
            console.log(`  Baris tanpa counter HERI: ${totalEmptyCounter}`);
        }
        if (totalWarnings > 0) {
            console.log(`  Peringatan: ${totalWarnings}`);
        }
        if (validationPath) {
            console.log(
                `  Validasi: ${validationPath} (${validationDetailRows.length} masalah)`
            );
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
