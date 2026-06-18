/**
 * Cek panjang karakter vs template LSMW (murni, tanpa alias/transform).
 *
 *   pnpm run template-lsmw-check
 *
 * 1. Pilih template: Protect Y | Equipment | Func Loc
 * 2. URL file sumber Google Drive
 *
 * - Protect Y: semua kolom sumber yang dipetakan
 * - Equipment / Func Loc: hanya FUNCTLOC DESC (sumber + output lokal jika ada)
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const { select, input } = require("@inquirer/prompts");
const {
    prepareTempDownloads,
    cleanupTempDownloads,
    downloadSpreadsheetToTemp,
    isProcessableSpreadsheet,
    GOOGLE_SHEET_MIME
} = require("./utils/lsmw_cli.cjs");
const { CHECK_MODES } = require("./utils/template_lsmw_check_config.cjs");
const { runTemplateCheck, cellPlainText } = require("./utils/template_lsmw_pure_check.cjs");

const OAUTH_PATH = "oauth.json";
const TOKEN_PATH = "token.json";
const SCOPES = ["https://www.googleapis.com/auth/drive"];
const OUTPUT_DIR = path.join("Output", "template-lsmw-cheking");

const EXCEL_TABLE_HEADERS = [
    "nama item",
    "jenis data",
    "nama kolom asal",
    "baris",
    "nilai",
    "panjang karakter",
    "panjang maksimal karakter"
];

let drive;

async function initDrive() {
    if (drive) return;

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
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(newAuth.credentials));
        auth.setCredentials(newAuth.credentials);
        console.log("Token saved to token.json");
    }

    drive = google.drive({ version: "v3", auth });
}

function extractFileIdFromUrl(raw) {
    if (!raw || typeof raw !== "string") return null;
    const s = raw.trim();
    if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
    const m =
        s.match(/\/d\/([a-zA-Z0-9_-]+)/) ||
        s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
        s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/) ||
        s.match(/[?&#]id=([a-zA-Z0-9_-]+)/);
    return m ? m[1] : null;
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

async function promptCheckMode() {
    const modeId = await select({
        message: "Pilih jenis template pengecekan:",
        choices: [
            {
                name: `1. Template Protect Y — ${CHECK_MODES["protect-y"].description}`,
                value: "protect-y"
            },
            {
                name: `2. Template Equipment — ${CHECK_MODES.equipment.description}`,
                value: "equipment"
            },
            {
                name: `3. Template Func Loc — ${CHECK_MODES["func-loc"].description}`,
                value: "func-loc"
            }
        ]
    });
    return CHECK_MODES[modeId];
}

async function promptSourceUrl() {
    const url = await input({
        message: "URL atau ID file Google Drive / Google Sheets (file sumber PKS):",
        validate: value => {
            const trimmed = String(value ?? "").trim();
            if (!trimmed) return "URL tidak boleh kosong";
            if (!extractFileIdFromUrl(trimmed)) {
                return "URL/ID tidak valid";
            }
            return true;
        }
    });
    return String(url).trim();
}

async function downloadSourceFile(sourceUrl, tempDir) {
    const fileId = extractFileIdFromUrl(sourceUrl);
    if (!fileId) throw new Error(`Tidak bisa parse file id: ${sourceUrl}`);

    await initDrive();
    const meta = await drive.files.get({
        fileId,
        fields: "id,name,mimeType,webViewLink",
        supportsAllDrives: true
    });

    const file = meta.data;
    if (!isProcessableSpreadsheet(file)) {
        throw new Error(
            `Bukan spreadsheet: ${file.name} (${file.mimeType})`
        );
    }

    console.log("Downloading:", file.name);
    const { localFile, sourceName } = await downloadSpreadsheetToTemp(
        drive,
        file,
        tempDir
    );

    const sourceUrlOut =
        file.mimeType === GOOGLE_SHEET_MIME
            ? `https://docs.google.com/spreadsheets/d/${file.id}/edit`
            : file.webViewLink ||
              `https://drive.google.com/file/d/${file.id}/view`;

    return { localFile, sourceName, sourceUrl: sourceUrlOut };
}

function summarizeOverLimitByColumn(violatingRows) {
    const byColumn = new Map();
    for (const row of violatingRows) {
        for (const v of row.columnsOverLimit) {
            const key = v.column;
            if (!byColumn.has(key)) {
                byColumn.set(key, {
                    column: v.column,
                    maxLength: v.maxLength,
                    rowCount: 0
                });
            }
            byColumn.get(key).rowCount += 1;
        }
    }
    return [...byColumn.values()].sort((a, b) => b.rowCount - a.rowCount);
}

function resolveNamaItem(vr) {
    if (vr.functionalLocation) return vr.functionalLocation;
    const desc = vr.columnsOverLimit[0];
    return desc?.value ?? "";
}

function groupViolationsByColumn(violatingRows) {
    const byColumn = new Map();
    for (const vr of violatingRows) {
        const namaItem = resolveNamaItem(vr);
        const jenisData = vr.target === "source" ? "sumber" : "output";
        const baris =
            vr.target === "source"
                ? vr.sourceExcelRow
                : vr.outputExcelRow;

        for (const v of vr.columnsOverLimit) {
            if (!byColumn.has(v.column)) byColumn.set(v.column, []);
            byColumn.get(v.column).push({
                namaItem,
                jenisData,
                namaKolomAsal: v.sourceColumn,
                baris,
                nilai: v.value,
                panjangKarakter: v.actualLength,
                panjangMaksimal: v.maxLength
            });
        }
    }
    return byColumn;
}

function sanitizeSheetName(name, usedNames) {
    let base = String(name)
        .replace(/[\\/*?:\[\]]/g, "_")
        .trim();
    if (!base) base = "Kolom";
    if (base.length > 31) base = base.slice(0, 31);
    let candidate = base;
    let n = 2;
    while (usedNames.has(candidate)) {
        const suffix = `_${n}`;
        candidate = base.slice(0, 31 - suffix.length) + suffix;
        n += 1;
    }
    usedNames.add(candidate);
    return candidate;
}

async function writeOverLimitWorkbook(filePath, violatingRows, overLimitByColumn) {
    const wb = new ExcelJS.Workbook();
    const usedSheetNames = new Set();
    const byColumn = groupViolationsByColumn(violatingRows);

    const columnOrder = overLimitByColumn.map(c => c.column);
    for (const col of byColumn.keys()) {
        if (!columnOrder.includes(col)) columnOrder.push(col);
    }

    for (const templateColumn of columnOrder) {
        const rows = byColumn.get(templateColumn);
        if (!rows || rows.length === 0) continue;

        const ws = wb.addWorksheet(
            sanitizeSheetName(templateColumn, usedSheetNames)
        );
        ws.addRow(EXCEL_TABLE_HEADERS);
        ws.getRow(1).font = { bold: true };
        for (const r of rows) {
            ws.addRow([
                r.namaItem,
                r.jenisData,
                r.namaKolomAsal,
                r.baris,
                r.nilai,
                r.panjangKarakter,
                r.panjangMaksimal
            ]);
        }
        ws.columns = [
            { width: 42 },
            { width: 10 },
            { width: 36 },
            { width: 10 },
            { width: 48 },
            { width: 16 },
            { width: 22 }
        ];
        ws.views = [{ state: "frozen", ySplit: 1 }];
    }

    if (wb.worksheets.length === 0) {
        const ws = wb.addWorksheet("Tidak ada pelanggaran");
        ws.addRow(EXCEL_TABLE_HEADERS);
        ws.getRow(1).font = { bold: true };
    }

    await wb.xlsx.writeFile(filePath);
}

async function main() {
    console.log("Template LSMW — cek panjang karakter (murni, tanpa alias)\n");

    const modeConfig = await promptCheckMode();
    console.log(`\n[ok] Mode: ${modeConfig.label}\n`);

    const sourceInput = await promptSourceUrl();

    prepareTempDownloads();
    const tempDir = path.join("Output", "temp_downloads");

    try {
        const { localFile, sourceName, sourceUrl } =
            await downloadSourceFile(sourceInput, tempDir);

        console.log(`[info] Cek sumber: ${sourceName}`);
        const result = await runTemplateCheck(
            modeConfig,
            localFile,
            sourceName
        );

        if (!result.ok) {
            console.error(`[err] ${result.reason}`);
            process.exit(1);
        }

        const { violatingRows, sourceScan, outputScan, outputPath } = result;
        const sourceViolations = violatingRows.filter(
            v => v.target === "source"
        );
        const outputViolations = violatingRows.filter(
            v => v.target === "output"
        );
        const overLimitByColumn = summarizeOverLimitByColumn(violatingRows);

        const generatedAt = wibTimestampForFilename();
        const safeMode = modeConfig.id.replace(/[^a-z0-9-]/gi, "-");
        const outBase = path.join(
            OUTPUT_DIR,
            `${generatedAt} - ${safeMode}`
        );
        const outFile = `${outBase}.json`;
        const outXlsx = `${outBase}.xlsx`;

        const report = {
            generatedAt,
            checkMode: modeConfig.id,
            checkModeLabel: modeConfig.label,
            template: result.template,
            source: {
                name: sourceName,
                url: sourceUrl,
                localPath: localFile,
                headerRowIndex: sourceScan.headerRowIndex,
                dataStartExcelRow: sourceScan.dataStartExcelRow,
                totalRows: sourceScan.totalRows,
                checksApplied: sourceScan.checksApplied
            },
            output: outputPath
                ? {
                      path: outputPath,
                      found: true,
                      totalRows: outputScan?.totalRows ?? 0,
                      checksApplied: outputScan?.checksApplied ?? [],
                      scanOk: outputScan?.ok ?? false,
                      scanError: outputScan?.ok
                          ? null
                          : outputScan?.reason ?? null
                  }
                : {
                      found: false,
                      hint:
                          modeConfig.outputRoot &&
                          `Output lokal tidak ditemukan di ${modeConfig.outputRoot} untuk ${sourceName}`
                  },
            summary: {
                totalSourceRows: sourceScan.totalRows,
                sourceRowsWithViolations: sourceViolations.length,
                totalOutputRows: outputScan?.totalRows ?? 0,
                outputRowsWithViolations: outputViolations.length,
                totalViolations: violatingRows.length,
                overLimitByColumn
            },
            violatingRows
        };

        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
        fs.writeFileSync(outFile, JSON.stringify(report, null, 2), "utf8");
        await writeOverLimitWorkbook(
            outXlsx,
            violatingRows,
            overLimitByColumn
        );

        console.log(
            `[ok] Sumber: ${sourceViolations.length} baris melebihi batas (dari ${sourceScan.totalRows})`
        );
        if (modeConfig.outputRoot) {
            if (outputPath) {
                console.log(
                    `[ok] Output: ${outputViolations.length} baris melebihi batas (dari ${outputScan?.totalRows ?? 0}) — ${outputPath}`
                );
            } else {
                console.log(
                    `[warn] File output LSMW tidak ditemukan untuk ${sourceName}`
                );
            }
        }
        if (overLimitByColumn.length > 0) {
            console.log("[ok] Over limit per kolom template:");
            for (const c of overLimitByColumn) {
                console.log(
                    `     ${c.column}: ${c.rowCount} baris (max ${c.maxLength})`
                );
            }
        }
        console.log(`[ok] JSON:  ${path.resolve(outFile)}`);
        console.log(`[ok] Excel: ${path.resolve(outXlsx)}`);
    } finally {
        cleanupTempDownloads();
    }
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
