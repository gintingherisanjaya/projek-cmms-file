/**
 * Enumerasi file PKS per regional dari folder Drive output measuring document.
 */

const path = require("path");
const {
    shouldSkipMeasuringDocAggregateFile
} = require("./measuring_document_counter_loader.cjs");

const SKIP_ROOT_FOLDER_NAMES = new Set(["Missing Measuring Document"]);

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function pksNameFromFileName(fileName) {
    return path
        .basename(String(fileName ?? ""))
        .replace(/\.xlsx$/i, "")
        .trim();
}

function pksXlsxFileName(pksName) {
    const base = String(pksName ?? "").trim();
    return `${base.replace(/[/\\?%*:|"<>]/g, "_")}.xlsx`;
}

/**
 * @param {{
 *   drive: import("googleapis").drive_v3.Drive,
 *   rootFolderId: string,
 *   isRegionalFolderName: (name: string) => boolean,
 *   isProcessableSpreadsheet: (file: object) => boolean,
 *   skipMissingFolder?: boolean
 * }} deps
 * @returns {Promise<Array<{ regional: string, pksName: string, file: object }>>}
 */
async function listMeasuringDocumentPksEntries(deps) {
    const {
        drive,
        rootFolderId,
        isRegionalFolderName,
        isProcessableSpreadsheet,
        skipMissingFolder = false
    } = deps;

    const res = await drive.files.list({
        q: `'${rootFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: "files(id,name)",
        pageSize: 1000,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
    });

    const regionalFolders = (res.data.files || []).filter(f => {
        const name = cellText(f.name);
        if (!isRegionalFolderName(name)) return false;
        if (skipMissingFolder && SKIP_ROOT_FOLDER_NAMES.has(name)) return false;
        return true;
    });

    /** @type {Array<{ regional: string, pksName: string, file: object }>} */
    const entries = [];

    for (const folder of regionalFolders) {
        const regional = cellText(folder.name);
        const childRes = await drive.files.list({
            q: `'${folder.id}' in parents and trashed=false`,
            fields: "files(id,name,mimeType)",
            pageSize: 1000,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
        });

        for (const item of childRes.data.files || []) {
            if (item.mimeType === "application/vnd.google-apps.folder") continue;
            if (!isProcessableSpreadsheet(item)) continue;
            if (shouldSkipMeasuringDocAggregateFile(item.name)) continue;

            entries.push({
                regional,
                pksName: pksNameFromFileName(item.name),
                file: item
            });
        }
    }

    entries.sort((a, b) => {
        const ra = a.regional.localeCompare(b.regional, "id");
        if (ra !== 0) return ra;
        return a.pksName.localeCompare(b.pksName, "id");
    });

    return entries;
}

module.exports = {
    SKIP_ROOT_FOLDER_NAMES,
    pksNameFromFileName,
    pksXlsxFileName,
    listMeasuringDocumentPksEntries
};
