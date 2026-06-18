/**
 * lsmw_tasklist.cjs
 * Konversi Protect Y → template LSMW tasklist (HEADER + OPERATION + SUB_OPERATION).
 *
 *   pnpm run lsmw-tasklist
 *
 * Prompt: URL folder Protect Y (subfolder REGIONAL*).
 * Output: Output/lsmw-tasklist/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx + all-pks.xlsx
 */

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const { input } = require("@inquirer/prompts");
const {
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    extractFolderIdFromUrl,
    isRegionalFolderName
} = require("./utils/lsmw_cli.cjs");
const { readSpreadsheetRows } = require("./utils/equipment_excel_io.cjs");
const { findDataLayout, normalizeHeader } = require("./utils/lsmw_cell_fill.cjs");
const {
    buildColumnIndexFirstWins,
    ensureCwcPlannerLookups,
    logRegionalMappingForFile,
    getEffectiveCostCenter,
    resolveWorkCenter,
    resolvePlannerGroup
} = require("./utils/lsmw_lookups.cjs");
const { withRetry } = require("./utils/lsmw_retry.cjs");
const {
    buildTasklistRows,
    buildSubOperationRows,
    formatSttagWib,
    filterUniqueBundles,
    bundlesToFlatRows
} = require("./utils/lsmw_tasklist_transform.cjs");
const {
    loadTasklistTemplateLayout,
    writeTasklistExcel
} = require("./utils/lsmw_tasklist_excel.cjs");
const {
    loadSubOperationTemplateLimits,
    validateOutputRow,
    formatCharViolation,
    OUTPUT_START_ROW: SUB_OPERATION_OUTPUT_START_ROW
} = require("./utils/lsmw_tasklist_limits.cjs");
const { findCostCenterColumnIndex } = require("./utils/equipment_gathering_columns.cjs");
const { compareRegionalFileEntries } = require("./utils/pks_sort_keys.cjs");

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly"
];
const OUTPUT_DIR = path.join("Output", "lsmw-tasklist");
const TEMPLATE_PATH = "./template_lsmw_tasklist.xlsx";
const ALL_PKS_FILE_NAME = "all-pks.xlsx";

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

function validateSubOperationRows(subOperationRows, limitByTemplate, fileName) {
    for (let i = 0; i < subOperationRows.length; i += 1) {
        const violations = validateOutputRow(
            subOperationRows[i],
            limitByTemplate,
            {
                fileName,
                sourceExcelRow: SUB_OPERATION_OUTPUT_START_ROW + i
            }
        );
        if (violations.length > 0) {
            stopWithError(formatCharViolation(violations[0]));
        }
    }
}

function buildAndValidateSubOperationRows(
    flat,
    templateLayout,
    subOperationLimitByTemplate,
    sttag,
    fileName
) {
    const subOperationRows = buildSubOperationRows({
        operationRows: flat.operationRows,
        operationRowBoldFlags: flat.operationRowBoldFlags ?? [],
        operationColumnOrder: templateLayout.operationColumnOrder,
        subOperationColumnOrder: templateLayout.subOperationColumnOrder,
        subOperationHourMeterColumns:
            templateLayout.subOperationHourMeterColumns,
        sttag
    });
    validateSubOperationRows(
        subOperationRows,
        subOperationLimitByTemplate,
        fileName
    );
    return subOperationRows;
}

function formatTasklistValidationError(fileName, validationError) {
    const toDebugText = value => {
        if (value === null || value === undefined) return "(kosong)";
        if (typeof value === "string") {
            const trimmed = value.trim();
            return trimmed === "" ? "(kosong)" : trimmed;
        }
        try {
            const json = JSON.stringify(value);
            return json ?? String(value);
        } catch {
            return String(value);
        }
    };

    const functionalLocation =
        String(validationError?.functionalLocation ?? "").trim() || "(kosong)";
    const functionalLocationDescription =
        String(validationError?.functionalLocationDescription ?? "").trim() ||
        "(kosong)";
    const costCenterRaw = toDebugText(validationError?.costCenterRaw);
    const effectiveCostCenter = toDebugText(validationError?.effectiveCostCenter);
    const costCenterSuffixForCsv = toDebugText(
        validationError?.costCenterSuffixForCsv
    );
    const missingFields = Array.isArray(validationError?.missingFields)
        ? validationError.missingFields.filter(Boolean)
        : [];
    const missingLabel = missingFields.length
        ? missingFields.join(", ")
        : "main work center / planner group";
    return [
        `Main work center / planner group kosong pada kandidat tasklist.`,
        `File: ${fileName}`,
        `FUNCTIONAL LOCATION AFTER: ${functionalLocation}`,
        `FUNCTLOC DESC. AFTER: ${functionalLocationDescription}`,
        `Cost center raw terbaca: ${costCenterRaw}`,
        `Effective cost center: ${effectiveCostCenter}`,
        `Cost center suffix lookup CSV: ${costCenterSuffixForCsv}`,
        `Field kosong: ${missingLabel}`
    ].join("\n  ");
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

    entries.sort(compareRegionalFileEntries);

    return entries;
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

async function downloadSpreadsheetWithRetry(file) {
    return withRetry(
        () =>
            downloadSpreadsheetToTemp(
                drive,
                file,
                TEMP_DOWNLOAD_PATH,
                "tl_"
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

    try {
        await initDrive();

        const protectYFolderUrl = await promptProtectYFolderUrl();
        const requireUniqueInFile = true;
        const protectYRootFolderId = extractFolderIdFromUrl(
            String(protectYFolderUrl).trim()
        );

        console.log("\nMode: wajib unik per file (HEADER + OPERATION).");

        console.log("\nMemuat layout template tasklist...");
        const templatePath = path.resolve(TEMPLATE_PATH);
        const templateLayout = await loadTasklistTemplateLayout(templatePath);
        const subOperationLimits = await loadSubOperationTemplateLimits(
            templatePath
        );
        const sttag = formatSttagWib(new Date());
        await ensureCwcPlannerLookups();

        console.log("Memuat file Protect Y (subfolder REGIONAL*)...");
        const entries = await listRegionalSpreadsheets(protectYRootFolderId);
        console.log(`  ${entries.length} spreadsheet`);

        if (entries.length === 0) {
            console.log("Tidak ada file untuk diproses.");
            return;
        }

        fs.mkdirSync(runDir, { recursive: true });

        let filesOk = 0;
        let filesSkipped = 0;
        let totalHeaderRows = 0;
        let totalOperationRows = 0;
        let totalSubOperationRows = 0;
        let totalDupSkipBundles = 0;

        const seenFuncLocKeys = new Set();
        /** @type {Array<{ bundle: object, regional: string, fileName: string, sortIndex: number }>} */
        const allBundleEntries = [];
        let allBundleSortIndex = 0;

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
            const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
            const idxCostCenter = findCostCenterColumnIndex(headerRow);

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
            await logRegionalMappingForFile(fileBegru, file.name);
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

            const built = buildTasklistRows(sheet.rawRows, {
                seenFuncLocKeys,
                effectiveCostCenterByRowIndex,
                workCenterByRowIndex,
                plannerGroupByRowIndex,
                headerColumnOrder: templateLayout.headerColumnOrder,
                operationColumnOrder: templateLayout.operationColumnOrder,
                hourMeterColumns: templateLayout.hourMeterColumns,
                requireUniqueInFile
            });

            if (!built.ok) {
                if (built.reason === "MISSING_CWC_PLANNER") {
                    stopWithError(
                        formatTasklistValidationError(
                            file.name,
                            built.validationError
                        )
                    );
                }
                console.log(`  [skip] ${built.reason}`);
                filesSkipped += 1;
                continue;
            }

            if (built.skippedDuplicateBundles > 0) {
                console.log(
                    `  [dup-skip] ${built.skippedDuplicateBundles} bundle (wajib unik) — ${file.name}`
                );
                totalDupSkipBundles += built.skippedDuplicateBundles;
            }

            if (built.headerRows.length === 0) {
                console.log(`  [skip] tidak ada baris tasklist yang cocok`);
                filesSkipped += 1;
                continue;
            }

            const subOperationRows = buildAndValidateSubOperationRows(
                built,
                templateLayout,
                subOperationLimits.limitByTemplate,
                sttag,
                file.name
            );

            await writeTasklistExcel(outPath, templatePath, {
                headerRows: built.headerRows,
                operationRows: built.operationRows,
                operationRowBoldFlags: built.operationRowBoldFlags,
                subOperationRows,
                headerColumnOrder: templateLayout.headerColumnOrder,
                operationColumnOrder: templateLayout.operationColumnOrder,
                subOperationColumnOrder: templateLayout.subOperationColumnOrder
            });

            if (built.bundles) {
                for (const bundle of built.bundles) {
                    allBundleEntries.push({
                        bundle,
                        regional,
                        fileName: file.name,
                        sortIndex: allBundleSortIndex
                    });
                    allBundleSortIndex += 1;
                }
            }

            console.log(
                `  → ${built.headerRows.length} HEADER, ${built.operationRows.length} OPERATION, ${subOperationRows.length} SUB_OPERATION → ${outPath}`
            );
            filesOk += 1;
            totalHeaderRows += built.headerRows.length;
            totalOperationRows += built.operationRows.length;
            totalSubOperationRows += subOperationRows.length;
        }

        if (allBundleEntries.length > 0) {
            allBundleEntries.sort(
                (a, b) =>
                    compareRegionalFileEntries(
                        { regional: a.regional, file: { name: a.fileName } },
                        { regional: b.regional, file: { name: b.fileName } }
                    ) || a.sortIndex - b.sortIndex
            );
            const allBundles = allBundleEntries.map(entry => entry.bundle);
            const filtered = filterUniqueBundles(allBundles, {
                headerColumnOrder: templateLayout.headerColumnOrder,
                operationColumnOrder: templateLayout.operationColumnOrder
            });
            const summaryBundles = filtered.bundles;
            const allPksDupSkip = filtered.skippedDuplicateBundles;
            if (allPksDupSkip > 0) {
                console.log(
                    `\n[dup-skip] ${allPksDupSkip} bundle pada agregat all-pks (wajib unik lintas PKS)`
                );
                totalDupSkipBundles += allPksDupSkip;
            }

            const summaryFlat = bundlesToFlatRows(summaryBundles);
            const summarySubOperationRows = buildAndValidateSubOperationRows(
                summaryFlat,
                templateLayout,
                subOperationLimits.limitByTemplate,
                sttag,
                ALL_PKS_FILE_NAME
            );
            const summaryPath = path.join(runDir, ALL_PKS_FILE_NAME);
            await writeTasklistExcel(summaryPath, templatePath, {
                ...summaryFlat,
                subOperationRows: summarySubOperationRows,
                headerColumnOrder: templateLayout.headerColumnOrder,
                operationColumnOrder: templateLayout.operationColumnOrder,
                subOperationColumnOrder: templateLayout.subOperationColumnOrder
            });
            console.log(`\nAgregat: ${summaryPath}`);
        }

        console.log("\nRingkasan:");
        console.log(`  Output: ${path.resolve(runDir)}`);
        console.log(`  File OK: ${filesOk}, dilewati: ${filesSkipped}`);
        console.log(`  Total baris HEADER: ${totalHeaderRows}`);
        console.log(`  Total baris OPERATION: ${totalOperationRows}`);
        console.log(`  Total baris SUB_OPERATION: ${totalSubOperationRows}`);
        if (totalDupSkipBundles > 0) {
            console.log(
                `  Total bundle di-skip (duplikat HEADER/OPERATION): ${totalDupSkipBundles}`
            );
        }
        if (allBundleEntries.length === 0) {
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
