const fs = require("fs");
const { google } = require("googleapis");

const LSMW_FOLDER_ID = "1aVCwpidscVY8bmXkEsgvH5BvG6OiJvLC";

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

async function walk(drive, folderId, pathParts, out) {
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
            const p = [...pathParts, f.name].join("/");
            if (f.mimeType === "application/vnd.google-apps.folder") {
                await walk(drive, f.id, [...pathParts, f.name], out);
            } else if (
                f.name.includes("BERANGIR") ||
                f.name.includes("DOLOK SINUMBA")
            ) {
                out.push({ path: p, name: f.name, mimeType: f.mimeType, id: f.id });
            }
        }
        pageToken = res.data.nextPageToken;
    } while (pageToken);
}

async function main() {
    const drive = await initDrive();
    const out = [];
    await walk(drive, LSMW_FOLDER_ID, [], out);
    console.log(JSON.stringify(out, null, 2));
}

main().catch(console.error);
