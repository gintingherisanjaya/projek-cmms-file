/**
 * Runner bersama untuk skrip delete LSMW (func loc / equipment).
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
    resolveSelectedRegionalFolders,
    lsmwOutputFileName
} = require("./lsmw_cli.cjs");
const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");
const { findFunctionalLocationBeforeColumnIndex } = require("./equipment_gathering_columns.cjs");
const {
    loadEquipmentMissingInProtectY,
    lookupMissingByPlantAndTarget,
    plantFromFuncLocAfter
} = require("./equipment_missing_in_protect_y.cjs");
const { loadEqktuByEquipment } = require("./truth_of_data_eqktu_mapping.cjs");
const {
    loadFunclocDescByFunctionalLocation
} = require("./truth_of_data_funcloc_mapping.cjs");
const { loadOpenOrderBlocklist } = require("./equipment_func_loc_open_order_blocklist.cjs");
const { regionalAggregateFileName } = require("./lsmw_pks_aggregates.cjs");

const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";

const OUTPUT_START_ROW = 4;

let drive;
let outputDriveFolderId = null;
let secondRowDump = [];
/** @type {Set<string> | Map<string, string>} */
let usedValuesGlobal;
let oldNameLookupMap;
let skippedBlocklist = 0;

function cellToJson(v) {
    if (v === undefined || v === null || v === "") return null;
    if (v instanceof Date) return v.toISOString();
    return v;
}

function normalizeHeader(name) {
    if (!name) return "";
    return name.toString().trim().toUpperCase();
}

function hasOldNameLookup(config) {
    return Boolean(config.oldNameLookup);
}

function loadOldNameLookupMap(config) {
    if (config.oldNameLookup === "eqktu-by-equipment") {
        const { byEquipment, duplicateEquipment } = loadEqktuByEquipment();
        console.log(
            `Truth of Data (EQKTU): ${byEquipment.size} equipment` +
                (duplicateEquipment
                    ? ` (${duplicateEquipment} duplikat diabaikan)`
                    : "")
        );
        return byEquipment;
    }

    if (config.oldNameLookup === "funcloc-desc-by-funcloc") {
        const { byFunctionalLocation, duplicateKeys } =
            loadFunclocDescByFunctionalLocation();
        console.log(
            `Truth of Data (FunctLocDescrip.): ${byFunctionalLocation.size} functional location` +
                (duplicateKeys ? ` (${duplicateKeys} duplikat diabaikan)` : "")
        );
        return byFunctionalLocation;
    }

    return null;
}

function lookupOldName(primary) {
    if (!oldNameLookupMap) return "";
    return oldNameLookupMap.get(primary) ?? "";
}

function trackGlobalEntry(primary, oldName, config) {
    if (hasOldNameLookup(config)) {
        if (!usedValuesGlobal.has(primary)) {
            usedValuesGlobal.set(primary, oldName);
        }
        return;
    }
    usedValuesGlobal.add(primary);
}

function hasGlobalEntry(primary, config) {
    if (hasOldNameLookup(config)) {
        return usedValuesGlobal.has(primary);
    }
    return usedValuesGlobal.has(primary);
}

function trackRegionalEntry(primary, oldName, config) {
    if (!config.regionalValues) return;
    if (hasOldNameLookup(config)) {
        if (!config.regionalValues.has(primary)) {
            config.regionalValues.set(primary, oldName);
        }
        return;
    }
    config.regionalValues.add(primary);
}

function globalEntriesToRows(config) {
    if (hasOldNameLookup(config)) {
        return [...usedValuesGlobal.entries()]
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([primary, oldName]) => ({ primary, oldName }));
    }
    return [...usedValuesGlobal]
        .sort((a, b) => a.localeCompare(b))
        .map(primary => ({ primary, oldName: "" }));
}

function regionalEntriesToRows(regionalValues, config) {
    if (hasOldNameLookup(config)) {
        return [...regionalValues.entries()]
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([primary, oldName]) => ({ primary, oldName }));
    }
    return [...regionalValues]
        .sort((a, b) => a.localeCompare(b))
        .map(primary => ({ primary, oldName: "" }));
}

function globalUniqueCount(config) {
    return hasOldNameLookup(config)
        ? usedValuesGlobal.size
        : usedValuesGlobal.size;
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

/**
 * @param {Array<{ primary: string, oldName: string }>} rows
 * @param {string} templatePath
 * @param {string} outPath
 * @param {object} config
 */
async function writeDeleteWorkbook(rows, templatePath, outPath, config = {}) {
    const workbookTemplate = new ExcelJS.Workbook();
    await workbookTemplate.xlsx.readFile(templatePath);

    const sheetTemplate = workbookTemplate.worksheets[0];
    sheetTemplate.conditionalFormattings = [];

    const sampleRow = sheetTemplate.getRow(3);
    const clearCols = hasOldNameLookup(config) ? 2 : 1;
    for (let c = 1; c <= clearCols; c += 1) {
        sampleRow.getCell(c).value = null;
    }
    sampleRow.commit();

    let startRow = OUTPUT_START_ROW;

    for (const rowData of rows) {
        const row = sheetTemplate.getRow(startRow++);
        row.getCell(1).value = String(rowData.primary);
        if (hasOldNameLookup(config)) {
            row.getCell(2).value = String(rowData.oldName);
        }
        row.commit();
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await workbookTemplate.xlsx.writeFile(outPath);
}

function resolveSourceColumnIndex(headerRow, colIndex, config) {
    if (config.sourceColumnHeader === "FUNCTIONAL LOCATION BEFORE") {
        return (
            findFunctionalLocationBeforeColumnIndex(headerRow) ??
            colIndex[config.sourceColumnHeader]
        );
    }
    return colIndex[config.sourceColumnHeader];
}

/**
 * @param {Array<Array<unknown>>} dataRows
 * @param {Array<unknown>} headerRow
 * @param {object} config
 * @returns {{ output: Array<{ primary: string, oldName: string }> } | { error: string }}
 */
function collectDeleteValues(dataRows, headerRow, config) {
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxSource = resolveSourceColumnIndex(headerRow, colIndex, config);
    const idxFuncLocAfter = colIndex["FUNCTIONAL LOCATION AFTER"];
    const isEquipmentMode = config.sourceColumnHeader === "EQUIPMENT NUMBER";
    const missingLookup = config.missingLookup;

    if (
        idxSource === undefined &&
        (!missingLookup || idxFuncLocAfter === undefined)
    ) {
        return { error: `No ${config.sourceColumnHeader} column` };
    }

    const seenInFile = new Set();
    const output = [];

    function tryAdd(primary, oldNameOverride) {
        const normalized = config.normalizeValue(primary);
        if (!normalized) return;
        if (seenInFile.has(normalized)) return;
        if (hasGlobalEntry(normalized, config)) return;
        if (config.blockedValues?.has(normalized)) {
            skippedBlocklist += 1;
            return;
        }

        const oldName = oldNameOverride ?? lookupOldName(normalized);

        seenInFile.add(normalized);
        trackGlobalEntry(normalized, oldName, config);
        trackRegionalEntry(normalized, oldName, config);
        output.push({ primary: normalized, oldName });
    }

    for (const r of dataRows) {
        if (idxSource !== undefined) {
            tryAdd(r[idxSource]);
        }

        if (!missingLookup || idxFuncLocAfter === undefined) {
            continue;
        }

        const funcLocAfter = config.normalizeValue(r[idxFuncLocAfter]);
        if (!funcLocAfter) continue;

        const rowPlant = plantFromFuncLocAfter(funcLocAfter);
        if (!rowPlant) continue;

        const matches = lookupMissingByPlantAndTarget(
            missingLookup,
            rowPlant,
            funcLocAfter
        );

        for (const m of matches) {
            if (isEquipmentMode) {
                tryAdd(m.equipment, m.description);
            } else {
                tryAdd(m.functionalLoc, m.descriptionFuncLocOld);
            }
        }
    }

    return { output };
}

async function buildDeleteFile(sourcePath, sourceName, outputDir, config) {
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

    const collected = collectDeleteValues(dataRows, headerRow, config);

    if (collected.error) {
        console.log(`${collected.error}:`, sourceName);
        return null;
    }

    const output = collected.output;
    if (output.length === 0) {
        return null;
    }

    const newName = lsmwOutputFileName(sourceName);
    const newPath = path.join(outputDir, newName);

    await writeDeleteWorkbook(output, config.templatePath, newPath, config);

    return { path: newPath, name: newName };
}

async function processFolder(sourceFolderId, localDir, driveTargetId, config) {
    const res = await drive.files.list({
        q: `'${sourceFolderId}' in parents and trashed=false`,
        fields: "files(id,name,mimeType)",
        pageSize: 1000
    });

    for (const file of res.data.files) {
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

            await processFolder(file.id, childLocal, childDrive, config);
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

        const { localFile, sourceName } = await downloadSpreadsheetToTemp(
            drive,
            file,
            TEMP_DOWNLOAD_PATH
        );

        const result = await buildDeleteFile(
            localFile,
            sourceName,
            localDir,
            config
        );

        if (!result) continue;

        console.log("Saved:", result.path);

        if (driveTargetId) {
            await uploadFile(result.path, result.name, driveTargetId);
        }
    }
}

async function writeAggregateIfNeeded(rows, templatePath, outPath, driveTargetId, config) {
    if (rows.length === 0) return;

    await writeDeleteWorkbook(rows, templatePath, outPath, config);
    console.log("Agregat:", outPath);

    if (driveTargetId) {
        await uploadFile(outPath, path.basename(outPath), driveTargetId);
    }
}

function resolveBlockedValues(blocklist, sourceColumnHeader) {
    if (sourceColumnHeader === "EQUIPMENT NUMBER") {
        return blocklist.blockedEquipment;
    }
    if (sourceColumnHeader === "FUNCTIONAL LOCATION BEFORE") {
        return blocklist.blockedFuncLoc;
    }
    return new Set();
}

async function runLsmwDeleteJob(config) {
    secondRowDump = [];
    skippedBlocklist = 0;
    usedValuesGlobal = hasOldNameLookup(config) ? new Map() : new Set();
    oldNameLookupMap = hasOldNameLookup(config)
        ? loadOldNameLookupMap(config)
        : null;

    const blocklist = loadOpenOrderBlocklist();
    config.blockedValues = resolveBlockedValues(
        blocklist,
        config.sourceColumnHeader
    );
    console.log(
        `Open order blocklist: ${blocklist.equipmentCount} equipment, ` +
            `${blocklist.funcLocCount} functional location`
    );

    const cli = parseLsmwCli(process.argv.slice(2));
    outputDriveFolderId = cli.outputDriveFolderId;

    fs.mkdirSync(config.localOutputRoot, { recursive: true });
    prepareTempDownloads();

    try {
        const missingData = loadEquipmentMissingInProtectY();
        config.missingLookup = missingData.byPlantAndTarget;
        console.log(
            `Missing equipment excel: ${missingData.rows.length} baris, ` +
                `${missingData.byPlantAndTarget.size} key Plant+Target Func Loc`
        );

        await initDrive();

        const sourceRootFolderId = await resolveProtectYSourceFolderId(cli);
        console.log("Sumber Protect Y:", sourceRootFolderId);

        if (outputDriveFolderId) {
            console.log("Drive upload enabled →", outputDriveFolderId);
        } else {
            console.log(
                "Local output only →",
                path.resolve(config.localOutputRoot)
            );
        }

        const { regionalFolders, selected } =
            await resolveSelectedRegionalFolders(drive, sourceRootFolderId);

        if (selected.size > 0) {
            console.log(`Regional diproses: ${[...selected].join(", ")}`);
        }

        for (const folder of regionalFolders) {
            console.log(`Processing ${folder.name}`);

            const localRegionalDir = path.join(
                config.localOutputRoot,
                folder.name
            );
            fs.mkdirSync(localRegionalDir, { recursive: true });

            let driveRegionalTarget = null;
            if (outputDriveFolderId) {
                driveRegionalTarget = await createFolderIfNotExists(
                    folder.name,
                    outputDriveFolderId
                );
            }

            const regionalValues = hasOldNameLookup(config)
                ? new Map()
                : new Set();
            const regionalConfig = { ...config, regionalValues };

            await processFolder(
                folder.id,
                localRegionalDir,
                driveRegionalTarget,
                regionalConfig
            );

            if (config.enableRegionalAggregate) {
                const regionalAggPath = path.join(
                    config.localOutputRoot,
                    regionalAggregateFileName(folder.name)
                );
                const regionalRows = regionalEntriesToRows(
                    regionalValues,
                    config
                );
                const driveRootTarget = outputDriveFolderId || null;
                await writeAggregateIfNeeded(
                    regionalRows,
                    config.templatePath,
                    regionalAggPath,
                    driveRootTarget,
                    config
                );
            }
        }

        if (regionalFolders.length === 0) {
            console.log("No REGIONAL folder found under source root");
        }

        if (config.enableGlobalAggregate) {
            const globalAggName =
                config.globalAggregateFileName || "all-pks.xlsx";
            const globalAggPath = path.join(
                config.localOutputRoot,
                globalAggName
            );
            const globalRows = globalEntriesToRows(config);
            let driveRootTarget = null;
            if (outputDriveFolderId) {
                driveRootTarget = outputDriveFolderId;
            }
            await writeAggregateIfNeeded(
                globalRows,
                config.templatePath,
                globalAggPath,
                driveRootTarget,
                config
            );
        }

        fs.writeFileSync(
            config.columnJsonPath,
            JSON.stringify(secondRowDump, null, 2),
            "utf8"
        );
        console.log(
            "Wrote",
            config.columnJsonPath,
            `(${secondRowDump.length} files)`
        );
        console.log(`Unique ${config.uniqueLabel} in output:`, globalUniqueCount(config));
        if (skippedBlocklist > 0) {
            console.log(
                `Baris di-skip (order belum close): ${skippedBlocklist}`
            );
        }
        console.log("DONE");
    } finally {
        cleanupTempDownloads();
    }
}

module.exports = { runLsmwDeleteJob, writeDeleteWorkbook };
