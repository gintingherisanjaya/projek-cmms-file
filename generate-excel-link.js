/**
 * generate-excel-link.js
 * Daftar link publik semua file Excel/Google Sheets di folder REGIONAL* (protect-y).
 *
 *   node generate-excel-link.js
 *   node generate-excel-link.js --protect-y-folder-link <url>
 *   pnpm run generate-excel-link
 *
 * Output: Output/PKS_LINK_{timestamp}.xlsx
 * Kolom: nama folder | nama pks | link
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { input } from "@inquirer/prompts";
import {
    initGoogleDrive,
    listFilesInFolder,
    listFilesRecursively,
    extractFolderIdFromUrl
} from "./utils/googleDrive.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OAUTH_JSON = path.join(__dirname, "oauth.json");

const GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet";
const XLSX_MIME =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const OUTPUT_DIR = path.join(__dirname, "Output");

function getJakartaStamp() {
    const fmt = new Intl.DateTimeFormat("sv-SE", {
        timeZone: "Asia/Jakarta",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false
    });
    const parts = fmt.formatToParts(new Date());
    const map = new Map(parts.map(p => [p.type, p.value]));
    return `${map.get("year")}${map.get("month")}${map.get("day")}_${map.get("hour")}${map.get("minute")}${map.get("second")}`;
}

function isRegionalFolderName(name) {
    return name.toUpperCase().trim().startsWith("REGIONAL");
}

function isProcessableSpreadsheet(file) {
    if (!file?.name || file.name.startsWith("~$")) return false;
    return (
        file.mimeType === XLSX_MIME ||
        file.mimeType === GOOGLE_SHEET_MIME ||
        file.name.toLowerCase().endsWith(".xlsx")
    );
}

/** Link view yang bisa diklik (folder sudah shared). */
function publicLink(fileId, mimeType) {
    if (mimeType === GOOGLE_SHEET_MIME) {
        return `https://docs.google.com/spreadsheets/d/${fileId}/edit?usp=sharing`;
    }
    return `https://drive.google.com/file/d/${fileId}/view?usp=sharing`;
}

/** Nama PKS: nama file tanpa ekstensi .xlsx */
function pksNameFromFile(fileName) {
    return fileName.replace(/\.xlsx$/i, "").trim();
}

function parseCli(argv) {
    const idx = argv.indexOf("--protect-y-folder-link");
    if (idx === -1) return null;

    const url = argv[idx + 1];
    if (!url) {
        throw new Error("--protect-y-folder-link memerlukan URL folder Drive");
    }

    const folderId = extractFolderIdFromUrl(url);
    if (!folderId) {
        throw new Error("URL folder Protect Y tidak valid");
    }

    return folderId;
}

async function resolveProtectYFolderId(argv) {
    const fromCli = parseCli(argv);
    if (fromCli) return fromCli;

    const url = await input({
        message:
            "URL folder Google Drive — Protect Y yang akan diproses (subfolder REGIONAL*):",
        validate: value => {
            const trimmed = String(value ?? "").trim();
            if (!trimmed) return "URL tidak boleh kosong";
            if (!extractFolderIdFromUrl(trimmed)) {
                return "URL folder Drive tidak valid";
            }
            return true;
        }
    });

    return extractFolderIdFromUrl(String(url).trim());
}

async function collectRows(rootFolderId) {
    const rootItems = await listFilesInFolder(rootFolderId);
    const regionalFolders = rootItems.filter(
        item =>
            item.mimeType === "application/vnd.google-apps.folder" &&
            isRegionalFolderName(item.name)
    );

    const rows = [];

    for (const folder of regionalFolders) {
        console.log("Scanning:", folder.name);
        const files = await listFilesRecursively(folder.id);

        for (const file of files) {
            if (!isProcessableSpreadsheet(file)) continue;

            rows.push({
                regional: folder.name,
                pks: pksNameFromFile(file.name),
                link: publicLink(file.id, file.mimeType)
            });
        }
    }

    rows.sort((a, b) => {
        const r = a.regional.localeCompare(b.regional, "id");
        if (r !== 0) return r;
        return a.pks.localeCompare(b.pks, "id");
    });

    return rows;
}

async function writeOutputExcel(rows, outPath) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("PKS Links");

    ws.columns = [
        { header: "nama folder", key: "regional", width: 18 },
        { header: "nama pks", key: "pks", width: 36 },
        { header: "link", key: "link", width: 72 }
    ];

    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true };

    for (const row of rows) {
        const excelRow = ws.addRow({
            regional: row.regional,
            pks: row.pks,
            link: row.link
        });
        excelRow.getCell(3).value = {
            text: row.link,
            hyperlink: row.link
        };
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await wb.xlsx.writeFile(outPath);
}

async function main() {
    const rootFolderId = await resolveProtectYFolderId(process.argv.slice(2));
    if (!rootFolderId) {
        throw new Error("URL folder Protect Y tidak valid");
    }

    await initGoogleDrive(OAUTH_JSON);

    console.log("Protect Y source →", rootFolderId);

    const rows = await collectRows(rootFolderId);
    console.log("Total files:", rows.length);

    const outPath = path.join(
        OUTPUT_DIR,
        `PKS_LINK_${getJakartaStamp()}.xlsx`
    );
    await writeOutputExcel(rows, outPath);

    console.log("Saved:", outPath);
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
