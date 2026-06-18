/**
 * equipment-group-checking.js
 * Bandingkan equipment group SAP (Excel) vs mapping lokal (JSON).
 *
 *   pnpm run equipment-group-check
 *
 * Input:
 *   equipment-group-mapping-sap.xlsx  — kolom ID, NAME
 *   equipment-group-mapping.json      — childrens + parents (key = NAME, value = ID)
 *
 * Output:
 *   Output/equipment-group-checking/{YYYY-MM-DD_HH-mm-ss WIB}.xlsx
 *   Kolom: ID | NAME | STATUS
 *   STATUS: MATCH | UNUSED | NEW
 *   Satu baris per ID per STATUS; beberapa NAME digabung dengan koma jika sama STATUS.
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";
import ExcelJS from "exceljs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const SAP_XLSX = path.join(__dirname, "equipment-group-mapping-sap.xlsx");
const MAPPING_JSON = path.join(__dirname, "equipment-group-mapping.json");
const OUTPUT_DIR = path.join(__dirname, "Output", "equipment-group-checking");

const STATUS = {
    MATCH: "MATCH",
    UNUSED: "UNUSED",
    NEW: "NEW"
};

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

function normalizeId(value) {
    return String(value ?? "").trim().toUpperCase();
}

function normalizeName(value) {
    return String(value ?? "").trim();
}

function loadSapRows() {
    if (!fs.existsSync(SAP_XLSX)) {
        throw new Error(`File tidak ditemukan: ${SAP_XLSX}`);
    }

    const workbook = XLSX.readFile(SAP_XLSX);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const byId = new Map();

    for (const row of rows) {
        const id = normalizeId(row.ID);
        const name = normalizeName(row.NAME);
        if (!id || !name) continue;
        if (!byId.has(id)) {
            byId.set(id, new Set());
        }
        byId.get(id).add(name);
    }

    return byId;
}

function joinNames(nameSet) {
    return [...nameSet]
        .sort((a, b) => a.localeCompare(b))
        .join(", ");
}

function loadJsonMappings() {
    if (!fs.existsSync(MAPPING_JSON)) {
        throw new Error(`File tidak ditemukan: ${MAPPING_JSON}`);
    }

    const raw = JSON.parse(fs.readFileSync(MAPPING_JSON, "utf8"));
    const pairs = [];

    for (const section of ["childrens", "parents"]) {
        for (const [name, id] of Object.entries(raw[section] ?? {})) {
            const normalizedId = normalizeId(id);
            const normalizedName = normalizeName(name);
            if (!normalizedId || !normalizedName) continue;
            pairs.push({ id: normalizedId, name: normalizedName });
        }
    }

    const idsInJson = new Set(pairs.map(p => p.id));
    const namesById = new Map();

    for (const { id, name } of pairs) {
        if (!namesById.has(id)) {
            namesById.set(id, new Set());
        }
        namesById.get(id).add(name);
    }

    return { pairs, idsInJson, namesById };
}

function buildResultRows(sapById, { idsInJson, namesById }) {
    /** @type {Map<string, Set<string>>} key = `${id}\0${status}` */
    const groups = new Map();

    const addNames = (id, status, names) => {
        const key = `${id}\0${status}`;
        if (!groups.has(key)) {
            groups.set(key, new Set());
        }
        const bucket = groups.get(key);
        for (const name of names) {
            const normalized = normalizeName(name);
            if (normalized) bucket.add(normalized);
        }
    };

    for (const [id, sapNames] of sapById) {
        const status = idsInJson.has(id) ? STATUS.MATCH : STATUS.UNUSED;
        addNames(id, status, sapNames);

        if (status === STATUS.MATCH) {
            addNames(id, status, namesById.get(id) ?? []);
        }
    }

    for (const [id, jsonNames] of namesById) {
        if (sapById.has(id)) continue;
        addNames(id, STATUS.NEW, jsonNames);
    }

    const statusOrder = { MATCH: 0, UNUSED: 1, NEW: 2 };
    const rows = [];

    for (const [key, nameSet] of groups) {
        const sep = key.indexOf("\0");
        const id = key.slice(0, sep);
        const status = key.slice(sep + 1);
        rows.push({ id, name: joinNames(nameSet), status });
    }

    rows.sort(
        (a, b) =>
            statusOrder[a.status] - statusOrder[b.status] ||
            a.id.localeCompare(b.id)
    );

    return rows;
}

async function writeOutputXlsx(rows, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Equipment Group Checking");

    sheet.columns = [
        { header: "ID", key: "id", width: 18 },
        { header: "NAME", key: "name", width: 40 },
        { header: "STATUS", key: "status", width: 12 }
    ];

    sheet.getRow(1).font = { bold: true };

    for (const row of rows) {
        sheet.addRow(row);
    }

    sheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1 + rows.length, column: 3 }
    };

    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    await workbook.xlsx.writeFile(outputPath);
}

async function main() {
    const sapById = loadSapRows();
    const jsonData = loadJsonMappings();
    const rows = buildResultRows(sapById, jsonData);

    const outputPath = path.join(
        OUTPUT_DIR,
        `${wibTimestampForFilename()}.xlsx`
    );

    await writeOutputXlsx(rows, outputPath);

    const counts = rows.reduce((acc, row) => {
        acc[row.status] = (acc[row.status] ?? 0) + 1;
        return acc;
    }, {});

    console.log("Equipment group checking selesai.");
    console.log(`  SAP (Excel): ${sapById.size} ID`);
    console.log(`  JSON: ${jsonData.idsInJson.size} ID unik, ${jsonData.pairs.length} pasangan`);
    console.log(`  MATCH: ${counts.MATCH ?? 0}`);
    console.log(`  UNUSED: ${counts.UNUSED ?? 0}`);
    console.log(`  NEW: ${counts.NEW ?? 0}`);
    console.log(`  Total baris: ${rows.length}`);
    console.log(`  Output: ${outputPath}`);
}

main().catch(err => {
    console.error(err.message || err);
    process.exit(1);
});
