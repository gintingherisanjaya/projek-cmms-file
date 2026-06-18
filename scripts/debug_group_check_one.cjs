/**
 * Debug satu file LSMW vs Protect Y dari Drive.
 *   node scripts/debug_group_check_one.cjs "3.PKS DOLOK SINUMBA.xlsx"
 */
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const {
    SOURCE_ROOT_FOLDER_ID,
    TEMP_DOWNLOAD_PATH,
    prepareTempDownloads,
    cleanupTempDownloads,
    downloadSpreadsheetToTemp
} = require("../utils/lsmw_cli.cjs");
const {
    readProtectYRows,
    readLsmwRows,
    compareMaps
} = require("../lsmw_group_result_checking.cjs");

const LSMW_FOLDER_ID = "1aVCwpidscVY8bmXkEsgvH5BvG6OiJvLC";
const OAUTH_PATH = path.join(__dirname, "..", "oauth.json");
const TOKEN_PATH = path.join(__dirname, "..", "token.json");

const targetName = process.argv[2] || "3.PKS DOLOK SINUMBA.xlsx";

async function initDrive() {
    const credentials = JSON.parse(fs.readFileSync(OAUTH_PATH));
    const { client_id, client_secret, redirect_uris } = credentials.installed;
    const auth = new google.auth.OAuth2(
        client_id,
        client_secret,
        redirect_uris[0]
    );
    auth.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
    return google.drive({ version: "v3", auth });
}

async function findFileByName(drive, folderId, name, out = []) {
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
        for (const f of res.data.files ?? []) {
            if (f.mimeType === "application/vnd.google-apps.folder") {
                await findFileByName(drive, f.id, name, out);
            } else if (
                f.name === name ||
                f.name === name.replace(/\.xlsx$/i, "")
            ) {
                out.push(f);
            }
        }
        pageToken = res.data.nextPageToken;
    } while (pageToken);
    return out;
}

async function main() {
    prepareTempDownloads();
    const drive = await initDrive();

    const lsmwHits = await findFileByName(drive, LSMW_FOLDER_ID, targetName);
    const pyHits = await findFileByName(
        drive,
        SOURCE_ROOT_FOLDER_ID,
        targetName
    );

    console.log("LSMW:", lsmwHits.map(f => `${f.name} [${f.mimeType}]`));
    console.log("Protect Y:", pyHits.map(f => `${f.name} [${f.mimeType}]`));

    if (!lsmwHits[0] || !pyHits[0]) {
        console.error("File tidak ditemukan di Drive");
        process.exit(1);
    }

    const { localFile: lsmwLocal } = await downloadSpreadsheetToTemp(
        drive,
        lsmwHits[0],
        TEMP_DOWNLOAD_PATH,
        "lsmw_"
    );
    const { localFile: pyLocal } = await downloadSpreadsheetToTemp(
        drive,
        pyHits[0],
        TEMP_DOWNLOAD_PATH,
        "py_"
    );

    const lsmw = readLsmwRows(lsmwLocal);
    console.log("\nreadLsmwRows:", lsmw.ok ? `OK (${lsmw.byFuncLoc.size})` : lsmw.reason);
    if (lsmw.ok) {
        console.log("  layout:", lsmw.layoutSource);
        console.log("  sample:", [...lsmw.byFuncLoc.entries()].slice(0, 2));
    }

    const py = readProtectYRows(pyLocal);
    console.log("\nreadProtectYRows:", py.ok ? `OK (${py.byFuncLoc.size})` : py.reason);
    if (py.ok) {
        console.log("  sample:", [...py.byFuncLoc.entries()].slice(0, 2));
    }

    if (lsmw.ok && py.ok) {
        const { items, differenceCount } = compareMaps(
            py.byFuncLoc,
            lsmw.byFuncLoc
        );
        console.log("\nBaris:", items.length, "berbeda:", differenceCount);
        console.log(items.filter((r, i, arr) => {
            const egS = String(r.equipmentGroupSource ?? "").toUpperCase();
            const egL = String(r.equipmentGroupLsmw ?? "").toUpperCase();
            return egS !== egL;
        }).slice(0, 5));
    }

    cleanupTempDownloads();
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
