import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import xlsx from 'xlsx';

import {
    initGoogleDrive,
    listFilesInFolder,
    listFilesRecursively,
    downloadFile,
    extractFolderIdFromUrl,
    getDrive,
} from './utils/googleDrive.js';
import { runWithConcurrency } from './utils/concurrencyHelpers.js';
import { createWriteStream } from 'node:fs';
import { pipeline } from 'node:stream/promises';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const OAUTH_JSON = path.join(__dirname, 'oauth.json');

const getJakartaStamp = () => {
    const fmt = new Intl.DateTimeFormat('sv-SE', {
        timeZone: 'Asia/Jakarta',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false,
    });
    const parts = fmt.formatToParts(new Date());
    const map = new Map(parts.map((p) => [p.type, p.value]));
    return `${map.get('year')}${map.get('month')}${map.get('day')}_${map.get('hour')}${map.get('minute')}${map.get('second')}`;
};

const ROOT_FOLDER_URL = 'https://drive.google.com/drive/folders/1qglupEcEi7Q83EC8bR7ja2F9GAn_WU3G';

async function downloadExcelOrGoogleSheet(file, destPath) {
    const drive = getDrive();
    fs.mkdirSync(path.dirname(destPath), { recursive: true });

    if (file.mimeType === 'application/vnd.google-apps.spreadsheet') {
        const res = await drive.files.export({
            fileId: file.id,
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }, { responseType: 'stream' });
        await pipeline(res.data, createWriteStream(destPath));
    } else {
        await downloadFile(file.id, destPath);
    }
}

const TEMP_DIR = path.join(__dirname, 'Output', 'temp_downloads');
const OUT_DIR = path.join(__dirname, 'Output', 'template-checking');

async function main() {
    await initGoogleDrive(OAUTH_JSON);

    if (!fs.existsSync(TEMP_DIR)) {
        fs.mkdirSync(TEMP_DIR, { recursive: true });
    }
    if (!fs.existsSync(OUT_DIR)) {
        fs.mkdirSync(OUT_DIR, { recursive: true });
    }

    const rootFolderId = extractFolderIdFromUrl(ROOT_FOLDER_URL);
    console.log(`Listing files in root folder ${rootFolderId}...`);

    const rootItems = await listFilesInFolder(rootFolderId);
    const regionalFolders = rootItems.filter(
        (item) => item.mimeType === 'application/vnd.google-apps.folder' && item.name.toUpperCase().startsWith('REGIONAL')
    );

    console.log(`Found ${regionalFolders.length} REGIONAL folders.`);

    const allXlsxFiles = [];

    for (const folder of regionalFolders) {
        console.log(`Scanning folder: ${folder.name}...`);
        const files = await listFilesRecursively(folder.id);
        const xlsxFiles = files.filter((f) => {
            const isExcel = f.name.toLowerCase().endsWith('.xlsx') && !f.name.startsWith('~$');
            const isGoogleSheet = f.mimeType === 'application/vnd.google-apps.spreadsheet';
            return isExcel || isGoogleSheet;
        });
        for (const f of xlsxFiles) {
            allXlsxFiles.push({
                ...f,
                regionalFolder: folder.name,
            });
        }
    }

    console.log(`Found ${allXlsxFiles.length} excel files to process.`);

    const result = {};

    // Download and process with concurrency
    await runWithConcurrency(allXlsxFiles, 3, async (file) => {
        const tempPath = path.join(TEMP_DIR, `${file.id}.xlsx`);

        try {
            console.log(`Downloading ${file.name}...`);
            await downloadExcelOrGoogleSheet(file, tempPath);

            const wb = xlsx.readFile(tempPath);
            for (const sheetName of wb.SheetNames) {
                const sheet = wb.Sheets[sheetName];
                const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

                let headerV = null;
                if (rows.length > 1) {
                    const secondRow = rows[1];
                    // Column V is index 21
                    if (secondRow && secondRow.length > 21) {
                        headerV = secondRow[20];
                    }
                }

                const headerKey = headerV ? String(headerV).trim() : 'EMPTY_OR_MISSING';

                if (!result[headerKey]) {
                    result[headerKey] = [];
                }

                result[headerKey].push(`${file.regionalFolder}/${file.path} - ${sheetName}`);
            }
        } catch (err) {
            console.error(`Error processing ${file.name}:`, err.message);
        } finally {
            // Clean up temp file
            if (fs.existsSync(tempPath)) {
                fs.unlinkSync(tempPath);
            }
        }
    });

    const timestamp = getJakartaStamp();
    const outPath = path.join(OUT_DIR, `${timestamp}.json`);

    const fileContent = `${JSON.stringify(result, null, 2)}\n`;
    fs.writeFileSync(outPath, fileContent, 'utf-8');

    console.log(`Processing complete. Results saved to ${outPath}`);
}

main().catch((err) => {
    console.error(err);
    process.exit(1);
});
