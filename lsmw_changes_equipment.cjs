/**
 * lsmw_changes_equipment.cjs
 * ─────────────────────────────────────────────────────────────────────────────
 * Pipeline perubahan equipment PKS → template_changes_equipment.xlsx.
 *
 * Sumber: prompt URL folder Protect Y (subfolder REGIONAL*), atau --protect-y-folder-link <url>.
 * Template: template_changes_equipment.xlsx — isi mulai baris 4.
 *
 * Mapping
 *   OLD_NAME ← EQKTU dari TRUTH_OF_DATA_30_JANUARI.xlsx (match Equipment = EQUIPMENT NUMBER)
 *   EQUNR ← EQUIPMENT NUMBER
 *   TPLNR ← FUNCTIONAL LOCATION AFTER
 *   NEW_NAME ← FUNCTLOC DESC. AFTER (opsional; kosong jika kolom tidak ada)
 *
 * Sumber tambahan: equipment-missing-in-protect-y.xlsx (Plant, Equipment, Functional Loc.,
 *   Target Func Loc, Percentage) — match Plant + Target Func Loc = FUNCTIONAL LOCATION AFTER.
 *
 * Filter
 *   - FUNCTIONAL LOCATION AFTER wajib; EQUIPMENT dari Protect Y dan/atau missing excel.
 *   - EQUIPMENT NUMBER unik (per file dan lintas file); duplikat diabaikan.
 *   - Tidak ada filter level / deskripsi (semua baris yang memenuhi syarat ikut).
 *   - Skip equipment atau FUNCTIONAL LOCATION BEFORE di equipment-func-loc-belum-close-order.xlsx (order belum close).
 *
 *   node lsmw_changes_equipment.cjs
 *   node lsmw_changes_equipment.cjs --protect-y-folder-link <url>
 *   node lsmw_changes_equipment.cjs --output-drive-folder-link <url>
 *
 * Output: Output/5. LSMW CHANGE EQUIPMENT SISI/***.xlsx + validation.xlsx
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
} = require("./utils/lsmw_cli.cjs");
const {
    findFunclocDescAfterColumnIndex,
    findFunctionalLocationBeforeColumnIndex
} = require("./utils/equipment_gathering_columns.cjs");
const { loadEqktuByEquipment } = require("./utils/truth_of_data_eqktu_mapping.cjs");
const { loadOpenOrderBlocklist } = require("./utils/equipment_func_loc_open_order_blocklist.cjs");
const {
    writeChangesEquipmentValidationExcel
} = require("./utils/lsmw_changes_equipment_validation_excel.cjs");
const {
    loadEquipmentMissingInProtectY,
    lookupMissingByPlantAndTarget,
    plantFromFuncLocAfter
} = require("./utils/equipment_missing_in_protect_y.cjs");

const SCOPES = ["https://www.googleapis.com/auth/drive"];

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";

const LOCAL_OUTPUT_ROOT = path.join("Output", "5. LSMW CHANGE EQUIPMENT SISI");
const TEMPLATE_PATH = "./template_changes_equipment.xlsx";
const COLUMN_JSON_PATH = path.join(LOCAL_OUTPUT_ROOT, "column.json");

const OUTPUT_START_ROW = 4;

let drive;
let outputDriveFolderId = null;
let eqktuByEquipment;

let secondRowDump = [];
const validationRows = [];

/** Primary key lintas file: EQUIPMENT NUMBER yang sudah masuk output. */
const usedEquipmentNumber = new Set();

/** Equipment / func loc before dengan order belum close — tidak boleh di-change. */
let blockedEquipment = new Set();
let blockedFuncLoc = new Set();
let skippedBlocklist = 0;
/** @type {Map<string, Array<object>> | null} */
let missingLookup = null;

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

function normalizeEquipmentNumber(v) {
    if (v === null || v === undefined || v === "") return "";
    return String(v).trim();
}

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function collectChangeCandidates(r, indices) {
    const { idxFuncLoc, idxEquipment, idxDescAfter, idxFuncLocBefore } =
        indices;
    const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);
    if (!funcLocKey) return null;

    const rowPlant = plantFromFuncLocAfter(funcLocKey);

    const newName =
        idxDescAfter !== undefined ? cellText(r[idxDescAfter]) : "";
    const funcLocBefore =
        idxFuncLocBefore !== undefined
            ? normalizeFuncLocKey(r[idxFuncLocBefore])
            : "";

    const eqFromProtectY =
        idxEquipment !== undefined
            ? normalizeEquipmentNumber(r[idxEquipment])
            : "";

    const missingMatches =
        missingLookup && rowPlant
            ? lookupMissingByPlantAndTarget(
                  missingLookup,
                  rowPlant,
                  funcLocKey
              )
            : [];

    const missingByEquipment = new Map();
    for (const m of missingMatches) {
        const eq = normalizeEquipmentNumber(m.equipment);
        if (eq) missingByEquipment.set(eq, m);
    }

    const candidates = [];

    if (eqFromProtectY) {
        const m = missingByEquipment.get(eqFromProtectY);
        candidates.push({
            equipmentNo: eqFromProtectY,
            percentage: m?.percentage ?? "",
            missingDescription: m?.description ?? "",
            funcLocBefore: m?.functionalLoc || funcLocBefore
        });
    }

    for (const m of missingMatches) {
        const eq = normalizeEquipmentNumber(m.equipment);
        if (!eq || eq === eqFromProtectY) continue;
        candidates.push({
            equipmentNo: eq,
            percentage: m.percentage ?? "",
            missingDescription: m.description ?? "",
            funcLocBefore: m.functionalLoc || funcLocBefore
        });
    }

    if (candidates.length === 0) return null;

    return { funcLocKey, newName, candidates };
}

function buildValidationNote(equipmentNo, oldName, newName) {
    const notes = [];
    if (!eqktuByEquipment.has(equipmentNo)) {
        notes.push("Equipment tidak ditemukan di Truth of Data");
    } else if (!oldName) {
        notes.push("EQKTU kosong di Truth of Data");
    }
    if (!newName) {
        notes.push("DESC AFTER kosong");
    }
    return notes.join("; ");
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

async function buildChangesEquipmentFile(
    sourcePath,
    sourceName,
    outputDir,
    regionalName
) {
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

    const colIndex = {};
    headerRow.forEach((h, i) => {
        const key = normalizeHeader(h);
        if (key) colIndex[key] = i;
    });

    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxEquipment = colIndex["EQUIPMENT NUMBER"];
    const idxFuncLocBefore = findFunctionalLocationBeforeColumnIndex(headerRow);
    const idxDescAfter = findFunclocDescAfterColumnIndex(headerRow);
    if (idxFuncLoc === undefined) {
        console.log("No FUNCTIONAL LOCATION AFTER column:", sourceName);
        return null;
    }

    const indices = {
        idxFuncLoc,
        idxEquipment,
        idxFuncLocBefore,
        idxDescAfter
    };

    const seenInFile = new Set();
    const output = [];

    for (const r of dataRows) {
        const collected = collectChangeCandidates(r, indices);
        if (!collected) continue;

        const { funcLocKey, newName, candidates } = collected;

        for (const candidate of candidates) {
            const equipmentNo = candidate.equipmentNo;
            const funcLocBefore = candidate.funcLocBefore;

            if (blockedEquipment.has(equipmentNo)) {
                skippedBlocklist += 1;
                continue;
            }
            if (funcLocBefore && blockedFuncLoc.has(funcLocBefore)) {
                skippedBlocklist += 1;
                continue;
            }
            if (seenInFile.has(equipmentNo)) continue;
            if (usedEquipmentNumber.has(equipmentNo)) continue;

            seenInFile.add(equipmentNo);
            usedEquipmentNumber.add(equipmentNo);

            const oldName =
                eqktuByEquipment.get(equipmentNo) ??
                candidate.missingDescription ??
                "";

            output.push([
                oldName,
                equipmentNo,
                funcLocKey,
                newName,
                candidate.percentage
            ]);

            const note = buildValidationNote(equipmentNo, oldName, newName);
            if (note) {
                validationRows.push({
                    subFolder: regionalName,
                    fileName: sourceName,
                    equnr: equipmentNo,
                    tplnr: funcLocKey,
                    oldName,
                    newName,
                    note
                });
            }
        }
    }

    const workbookTemplate = new ExcelJS.Workbook();
    await workbookTemplate.xlsx.readFile(TEMPLATE_PATH);

    const sheetTemplate = workbookTemplate.worksheets[0];

    // Template punya conditional formatting (duplicateValues) di B1:C1.
    // ExcelJS menulis ulang tanpa cfRule → Excel menampilkan dialog recovery.
    sheetTemplate.conditionalFormattings = [];

    // Baris 3 template berisi contoh angka (18, 40), bukan data LSMW.
    const sampleRow = sheetTemplate.getRow(3);
    for (let c = 1; c <= 5; c += 1) {
        sampleRow.getCell(c).value = null;
    }
    sampleRow.commit();

    let startRow = OUTPUT_START_ROW;

    for (const rowData of output) {
        const row = sheetTemplate.getRow(startRow++);
        row.getCell(1).value = String(rowData[0]);
        row.getCell(2).value = String(rowData[1]);
        row.getCell(3).value = String(rowData[2]);
        row.getCell(4).value = String(rowData[3]);
        const pct = rowData[4];
        row.getCell(5).value =
            pct === "" || pct === null || pct === undefined
                ? ""
                : typeof pct === "number"
                  ? pct
                  : String(pct);
        row.commit();
    }

    const newName = lsmwOutputFileName(sourceName);
    const newPath = path.join(outputDir, newName);

    fs.mkdirSync(outputDir, { recursive: true });
    await workbookTemplate.xlsx.writeFile(newPath);

    return { path: newPath, name: newName };
}

async function processFolder(
    sourceFolderId,
    localDir,
    driveTargetId,
    regionalName
) {
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

            await processFolder(file.id, childLocal, childDrive, regionalName);
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

        const result = await buildChangesEquipmentFile(
            localFile,
            sourceName,
            localDir,
            regionalName
        );

        if (!result) continue;

        console.log("Saved:", result.path);

        if (driveTargetId) {
            await uploadFile(result.path, result.name, driveTargetId);
        }
    }
}

async function main() {
    secondRowDump = [];
    validationRows.length = 0;
    usedEquipmentNumber.clear();
    skippedBlocklist = 0;

    const cli = parseLsmwCli(process.argv.slice(2));
    outputDriveFolderId = cli.outputDriveFolderId;

    fs.mkdirSync(LOCAL_OUTPUT_ROOT, { recursive: true });
    prepareTempDownloads();

    try {
        const truthMapping = loadEqktuByEquipment();
        eqktuByEquipment = truthMapping.byEquipment;
        console.log(
            `Truth of Data: ${eqktuByEquipment.size} equipment` +
                (truthMapping.duplicateEquipment
                    ? ` (${truthMapping.duplicateEquipment} duplikat diabaikan)`
                    : "")
        );

        const blocklist = loadOpenOrderBlocklist();
        blockedEquipment = blocklist.blockedEquipment;
        blockedFuncLoc = blocklist.blockedFuncLoc;
        console.log(
            `Open order blocklist: ${blocklist.equipmentCount} equipment, ` +
                `${blocklist.funcLocCount} functional location`
        );

        const missingData = loadEquipmentMissingInProtectY();
        missingLookup = missingData.byPlantAndTarget;
        console.log(
            `Missing equipment excel: ${missingData.rows.length} baris, ` +
                `${missingLookup.size} key Plant+Target Func Loc`
        );

        await initDrive();

        const sourceRootFolderId = await resolveProtectYSourceFolderId(cli);
        console.log("Sumber Protect Y:", sourceRootFolderId);

        if (outputDriveFolderId) {
            console.log("Drive upload enabled →", outputDriveFolderId);
        } else {
            console.log(
                "Local output only →",
                path.resolve(LOCAL_OUTPUT_ROOT)
            );
        }

        const { regionalFolders, selected } =
            await resolveSelectedRegionalFolders(drive, sourceRootFolderId);

        if (selected.size > 0) {
            console.log(`Regional diproses: ${[...selected].join(", ")}`);
        }

        for (const folder of regionalFolders) {
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

            await processFolder(
                folder.id,
                localRegionalDir,
                driveRegionalTarget,
                folder.name
            );
        }

        if (regionalFolders.length === 0) {
            console.log("No REGIONAL folder found under source root");
        }

        if (validationRows.length > 0) {
            const validationPath = path.join(
                LOCAL_OUTPUT_ROOT,
                "validation.xlsx"
            );
            await writeChangesEquipmentValidationExcel(
                validationPath,
                validationRows
            );
            console.log(
                `Validasi: ${validationPath} (${validationRows.length} baris)`
            );
        }

        fs.writeFileSync(
            COLUMN_JSON_PATH,
            JSON.stringify(secondRowDump, null, 2),
            "utf8"
        );
        console.log("Wrote", COLUMN_JSON_PATH, `(${secondRowDump.length} files)`);
        console.log(
            "Unique EQUIPMENT NUMBER in output:",
            usedEquipmentNumber.size
        );
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

main().catch(err => {
    console.error(err);
    process.exit(1);
});
