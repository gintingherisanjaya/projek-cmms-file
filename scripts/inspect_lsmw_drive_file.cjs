const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { google } = require("googleapis");
const {
    downloadSpreadsheetToTemp,
    prepareTempDownloads
} = require("../utils/lsmw_cli.cjs");

const LSMW_FOLDER_ID = "1aVCwpidscVY8bmXkEsgvH5BvG6OiJvLC";
const nameFilter = process.argv[2] || "BERANGIR";

async function initDrive() {
    const credentials = JSON.parse(fs.readFileSync("oauth.json"));
    const { client_id, client_secret, redirect_uris } = credentials.installed;
    const auth = new google.auth.OAuth2(
        client_id,
        client_secret,
        redirect_uris[0]
    );
    auth.setCredentials(JSON.parse(fs.readFileSync("token.json")));
    return google.drive({ version: "v3", auth });
}

async function findContains(drive, folderId, needle, out = []) {
    let pageToken;
    do {
        const res = await drive.files.list({
            q: `'${folderId}' in parents and trashed=false`,
            fields: "nextPageToken, files(id,name,mimeType,parents)",
            pageSize: 1000,
            pageToken,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
        });
        for (const f of res.data.files ?? []) {
            if (f.mimeType === "application/vnd.google-apps.folder") {
                await findContains(drive, f.id, needle, out);
            } else if (f.name.toUpperCase().includes(needle.toUpperCase())) {
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
    const hits = await findContains(drive, LSMW_FOLDER_ID, nameFilter);
    console.log("Files:", hits.map(f => f.name));
    for (const file of hits.slice(0, 2)) {
        console.log("\n===", file.name, "===");
        const { localFile } = await downloadSpreadsheetToTemp(
            drive,
            file,
            path.join("Output", "temp_downloads")
        );
        const wb = XLSX.readFile(localFile);
        const s = wb.Sheets[wb.SheetNames[0]];
        for (let r = 1; r <= 10; r++) {
            const parts = [];
            for (let c = 1; c <= 22; c++) {
                const a = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
                const cell = s[a];
                if (cell) parts.push(`c${c}:${cell.w ?? cell.v}`);
            }
            if (parts.length) console.log(`r${r}`, parts.join(" | "));
        }
    }
}

main().catch(console.error);
