/**
 * CLI & konstanta bersama untuk skrip LSMW (equipment / func loc).
 */

const fs = require("fs");
const path = require("path");
const { input, checkbox } = require("@inquirer/prompts");
const { parseRegionalNumber } = require("./pks_sort_keys.cjs");

/** Fallback bila skrip lain belum memakai prompt/flag sumber dinamis. */
const SOURCE_ROOT_FOLDER_ID = "1qglupEcEi7Q83EC8bR7ja2F9GAn_WU3G";

/** LPP Standart — 1. STANDARISASI NAMA MESIN ALAT PKS ALL REGIONAL */
const LPP_STANDARD_ROOT_FOLDER_ID = "1cCSNlTObRnmUz_7NOO7saYtYh2ehVjys";

/** Unduhan sumber sementara; dibersihkan setelah tiap run LSMW selesai. */
const TEMP_DOWNLOAD_PATH = path.join("Output", "temp_downloads");

const GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet";
const XLSX_MIME =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const LSMW_OUTPUT_ROOTS = [
    path.join("Output", "1. LSMW Create Equipment V.O"),
    path.join("Output", "2. LSMW Create Functional Location V1"),
    path.join("Output", "3. LSMW DELETE MASS FUNCTIONAL LOCATION SISI"),
    path.join("Output", "4. LSMW DELETE EQUIPMENT SISI"),
    path.join("Output", "5. LSMW CHANGE EQUIPMENT SISI")
];

function removeLegacyDotDownloadsFolders() {
    for (const root of LSMW_OUTPUT_ROOTS) {
        const legacy = path.join(root, ".downloads");
        if (fs.existsSync(legacy)) {
            fs.rmSync(legacy, { recursive: true, force: true });
        }
    }
}

function prepareTempDownloads() {
    removeLegacyDotDownloadsFolders();
    if (fs.existsSync(TEMP_DOWNLOAD_PATH)) {
        fs.rmSync(TEMP_DOWNLOAD_PATH, { recursive: true, force: true });
    }
    fs.mkdirSync(TEMP_DOWNLOAD_PATH, { recursive: true });
}

function isProcessableSpreadsheet(file) {
    if (!file || file.name.startsWith("~$")) return false;
    return file.mimeType === XLSX_MIME || file.mimeType === GOOGLE_SHEET_MIME;
}

/**
 * Unduh .xlsx native atau export Google Sheets ke folder temp.
 * @param {string} [localPrefix] Prefiks nama file lokal (hindari tabrakan LSMW vs Protect Y).
 * @returns {{ localFile: string, sourceName: string }}
 */
async function downloadSpreadsheetToTemp(drive, file, tempDir, localPrefix = "") {
    const baseName = file.name.replace(/[/\\?%*:|"<>]/g, "_");
    let localFileName = `${localPrefix}${baseName}`;
    if (!localFileName.toLowerCase().endsWith(".xlsx")) {
        localFileName += ".xlsx";
    }

    const localFile = path.join(tempDir, localFileName);
    const dest = fs.createWriteStream(localFile);

    if (file.mimeType === GOOGLE_SHEET_MIME) {
        const res = await drive.files.export(
            { fileId: file.id, mimeType: XLSX_MIME },
            { responseType: "stream" }
        );
        await new Promise((resolve, reject) => {
            res.data.pipe(dest).on("finish", resolve).on("error", reject);
        });
    } else {
        const res = await drive.files.get(
            { fileId: file.id, alt: "media" },
            { responseType: "stream" }
        );
        await new Promise((resolve, reject) => {
            res.data.pipe(dest).on("finish", resolve).on("error", reject);
        });
    }

    return { localFile, sourceName: localFileName };
}

function cleanupTempDownloads() {
    if (fs.existsSync(TEMP_DOWNLOAD_PATH)) {
        fs.rmSync(TEMP_DOWNLOAD_PATH, { recursive: true, force: true });
    }
}

function extractFolderIdFromUrl(url) {
    const match = String(url).match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (!match) {
        throw new Error(`Invalid Google Drive folder URL: ${url}`);
    }
    return match[1];
}

function extractSpreadsheetIdFromUrl(url) {
    const match = String(url).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
        throw new Error(`Invalid Google Spreadsheet URL: ${url}`);
    }
    return match[1];
}

async function promptDriveFolderUrl(message) {
    return input({
        message,
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
    return promptDriveFolderUrl(
        "URL folder Google Drive — Protect Y yang akan diproses (subfolder REGIONAL*):"
    );
}

async function promptCheckRowsFolderUrl() {
    return promptDriveFolderUrl(
        "URL folder Google Drive (subfolder REGIONAL*):"
    );
}

function parseCheckRowsCli(argv) {
    const out = { folderId: null, onMismatch: null };
    const idx = argv.indexOf("--folder-link");
    if (idx !== -1 && argv[idx + 1]) {
        out.folderId = extractFolderIdFromUrl(argv[idx + 1]);
    }
    const mismatchIdx = argv.indexOf("--on-mismatch");
    if (mismatchIdx !== -1 && argv[mismatchIdx + 1]) {
        const raw = String(argv[mismatchIdx + 1]).trim().toLowerCase();
        if (raw === "terminate" || raw === "warn" || raw === "warning") {
            out.onMismatch = raw === "terminate" ? "terminate" : "warn";
        }
    }
    return out;
}

async function resolveCheckRowsFolderId(cli) {
    if (cli?.folderId) return cli.folderId;
    const url = await promptCheckRowsFolderUrl();
    return extractFolderIdFromUrl(String(url).trim());
}

/**
 * @param {string[]} argv
 * @returns {{ outputDriveFolderId: string | null, sourceRootFolderId: string | null, withCostCenter: boolean }}
 */
function parseLsmwCli(argv) {
    const out = {
        outputDriveFolderId: null,
        sourceRootFolderId: null,
        withCostCenter: argv.includes("--with-cost-center")
    };
    const idx = argv.indexOf("--output-drive-folder-link");
    if (idx !== -1 && argv[idx + 1]) {
        out.outputDriveFolderId = extractFolderIdFromUrl(argv[idx + 1]);
    }
    const srcIdx = argv.indexOf("--protect-y-folder-link");
    if (srcIdx !== -1 && argv[srcIdx + 1]) {
        out.sourceRootFolderId = extractFolderIdFromUrl(argv[srcIdx + 1]);
    }
    return out;
}

/** Prompt URL Protect Y bila belum ada di CLI (--protect-y-folder-link). */
async function resolveProtectYSourceFolderId(cli) {
    if (cli?.sourceRootFolderId) return cli.sourceRootFolderId;
    const url = await promptProtectYFolderUrl();
    return extractFolderIdFromUrl(String(url).trim());
}

function isRegionalFolderName(name) {
    return name.toUpperCase().trim().startsWith("REGIONAL");
}

/**
 * Subfolder REGIONAL* langsung di bawah root Drive (tanpa rekursif).
 * @returns {Promise<Array<{ id: string, name: string }>>}
 */
async function listRegionalFolders(drive, rootFolderId) {
    const out = [];
    let pageToken;
    do {
        const res = await drive.files.list({
            q: `'${rootFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
            fields: "nextPageToken, files(id,name)",
            pageSize: 1000,
            pageToken,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
        });
        for (const f of res.data.files ?? []) {
            if (isRegionalFolderName(f.name)) {
                out.push({ id: f.id, name: String(f.name).trim() });
            }
        }
        pageToken = res.data.nextPageToken;
    } while (pageToken);

    out.sort((a, b) => {
        const byNum = parseRegionalNumber(a.name) - parseRegionalNumber(b.name);
        if (byNum !== 0) return byNum;
        return a.name.localeCompare(b.name, "id");
    });
    return out;
}

/**
 * @param {Array<{ id: string, name: string }>} regionalFolders
 * @returns {Promise<Set<string>>}
 */
async function promptRegionalFolderSelection(regionalFolders) {
    const selected = await checkbox({
        message: "Pilih subfolder REGIONAL yang akan diproses:",
        choices: regionalFolders.map(f => ({
            name: f.name,
            value: f.name,
            checked: true
        })),
        validate: value =>
            (Array.isArray(value) && value.length > 0) ||
            "Pilih minimal satu regional"
    });
    return new Set(selected);
}

/**
 * Daftar REGIONAL* di Drive lalu prompt checkbox pemilihan.
 * @returns {Promise<{ all: Array<{ id: string, name: string }>, selected: Set<string>, regionalFolders: Array<{ id: string, name: string }> }>}
 */
async function resolveSelectedRegionalFolders(drive, rootFolderId) {
    const all = await listRegionalFolders(drive, rootFolderId);
    if (all.length === 0) {
        return { all, selected: new Set(), regionalFolders: [] };
    }
    const selected = await promptRegionalFolderSelection(all);
    const regionalFolders = all.filter(f => selected.has(f.name.trim()));
    return { all, selected, regionalFolders };
}

/** Nama file output = nama file sumber (tanpa akhiran _EQUIPMENT, dll.). */
function lsmwOutputFileName(sourceName) {
    const base = path.basename(sourceName);
    if (base.toLowerCase().endsWith(".xlsx")) return base;
    return `${base}.xlsx`;
}

module.exports = {
    SOURCE_ROOT_FOLDER_ID,
    LPP_STANDARD_ROOT_FOLDER_ID,
    TEMP_DOWNLOAD_PATH,
    GOOGLE_SHEET_MIME,
    XLSX_MIME,
    prepareTempDownloads,
    cleanupTempDownloads,
    removeLegacyDotDownloadsFolders,
    isProcessableSpreadsheet,
    downloadSpreadsheetToTemp,
    extractFolderIdFromUrl,
    extractSpreadsheetIdFromUrl,
    promptProtectYFolderUrl,
    promptCheckRowsFolderUrl,
    parseCheckRowsCli,
    resolveCheckRowsFolderId,
    parseLsmwCli,
    resolveProtectYSourceFolderId,
    isRegionalFolderName,
    listRegionalFolders,
    promptRegionalFolderSelection,
    resolveSelectedRegionalFolders,
    lsmwOutputFileName
};
