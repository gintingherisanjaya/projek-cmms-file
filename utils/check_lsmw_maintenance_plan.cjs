/**
 * Util check maintenance plan (SZAEH) vs CDS CountReadng (ZPP jam jalan).
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");
const { excelCellToPlain } = require("./equipment_excel_io.cjs");
const { findColumnIndexByPrefix } = require("./measuring_document_counter_loader.cjs");
const { MAINTENANCE_PLAN_OUTPUT_START_ROW } = require("./lsmw_maintenance_plan_limits.cjs");

const CDS_DATA_START_ROW = 1;
const HEADER_SCAN_COLS = 30;

const CHECK_HEADERS = [
    "nama regional",
    "nama pks",
    "measuring point",
    "equipment",
    "functional loc",
    "SZAEH",
    "CountReadng",
    "tanggal",
    "status"
];

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

function isFuncLocDescendant(childFuncLoc, ancestorFuncLoc) {
    const child = normalizeFuncLocKey(childFuncLoc);
    const ancestor = normalizeFuncLocKey(ancestorFuncLoc);
    return Boolean(child && ancestor && child.startsWith(`${ancestor}-`));
}

/**
 * @param {import("exceljs").Cell} cell
 * @returns {string}
 */
function getExcelCellDisplayText(cell) {
    const text = cellText(cell?.text);
    if (text) return text;

    const plain = excelCellToPlain(cell?.value);
    if (plain === null || plain === undefined || plain === "") return "";
    if (typeof plain === "string") return plain.trim();
    if (typeof plain === "number" && Number.isFinite(plain)) {
        if (Number.isInteger(plain)) return String(plain);
        return "";
    }
    return cellText(plain);
}

/**
 * @param {import("xlsx").CellObject | undefined} cell
 * @returns {string}
 */
function getXlsxCellDisplayText(cell) {
    if (!cell) return "";
    const w = cell.w;
    if (w !== undefined && w !== null && String(w).trim() !== "") {
        return String(w).trim();
    }
    if (cell.t === "s") return String(cell.v ?? "").trim();
    if (cell.t === "n" && Number.isInteger(cell.v)) return String(cell.v);
    if (cell.t === "b") return String(cell.v);
    return "";
}

/**
 * @param {import("exceljs").Worksheet} ws
 * @param {number} rowNumber
 * @returns {string[]}
 */
function buildHeaderRowFromSheet(ws, rowNumber) {
    const headerRow = [];
    for (let c = 1; c <= HEADER_SCAN_COLS; c += 1) {
        headerRow.push(getExcelCellDisplayText(ws.getRow(rowNumber).getCell(c)));
    }
    return headerRow;
}

/**
 * @param {string} display
 * @returns {string}
 */
function canonicalCounterForExactCompare(display) {
    let s = cellText(display).replace(/\s/g, "");
    if (!s) return "0";
    if (s.includes(",")) return s.replace(/\./g, "");
    if (/^-?\d+\.\d+$/.test(s)) return s.replace(".", ",");
    return s;
}

/**
 * @param {string} displayA
 * @param {string} displayB
 * @returns {boolean}
 */
function countersEqualExact(displayA, displayB) {
    return (
        canonicalCounterForExactCompare(displayA) ===
        canonicalCounterForExactCompare(displayB)
    );
}

/**
 * @param {Map<string, string>} byEquipment
 * @returns {Map<string, string[]>}
 */
function buildPointToEquipmentsMap(byEquipment) {
    /** @type {Map<string, string[]>} */
    const byPoint = new Map();

    for (const [equipment, point] of byEquipment) {
        const key = cellText(point);
        if (!key) continue;
        if (!byPoint.has(key)) byPoint.set(key, []);
        byPoint.get(key).push(cellText(equipment));
    }

    for (const list of byPoint.values()) {
        list.sort((a, b) => a.localeCompare(b, "id"));
    }

    return byPoint;
}

/**
 * @param {string} cdsPath
 * @returns {Promise<Map<string, string>>}
 */
async function loadCdsCountReadngIndex(cdsPath) {
    const byEquipment = await loadCdsEntriesByEquipment(cdsPath);
    /** @type {Map<string, string>} */
    const index = new Map();
    for (const [equipment, entries] of byEquipment) {
        if (entries.length > 0) {
            index.set(equipment, entries[0].countReadng);
        }
    }
    return index;
}

/**
 * @param {string} cdsPath
 * @returns {Promise<Map<string, Array<{ countReadng: string, date: string }>>>}
 */
async function loadCdsEntriesByEquipment(cdsPath) {
    const resolved = path.resolve(cdsPath);
    if (!fs.existsSync(resolved)) {
        throw new Error(`File CDS tidak ditemukan: ${resolved}`);
    }

    const wb = XLSX.readFile(resolved);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    if (!sheet || !sheet["!ref"]) {
        throw new Error("CDS tidak memiliki sheet");
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const headerRow = rows[0] || [];
    const idxEquipment = findColumnIndexByPrefix(headerRow, "Equipment");
    const idxCount = findColumnIndexByPrefix(headerRow, "CountReadng");
    const idxDate = findColumnIndexByPrefix(headerRow, "Date");

    if (idxEquipment < 0 || idxCount < 0) {
        throw new Error(
            "Kolom wajib CDS tidak ditemukan: Equipment, CountReadng"
        );
    }

    /** @type {Map<string, Array<{ countReadng: string, date: string }>>} */
    const byEquipment = new Map();

    for (let r = CDS_DATA_START_ROW; r < rows.length; r += 1) {
        const eqCell = sheet[XLSX.utils.encode_cell({ r, c: idxEquipment })];
        const cntCell = sheet[XLSX.utils.encode_cell({ r, c: idxCount })];
        const dateCell =
            idxDate >= 0
                ? sheet[XLSX.utils.encode_cell({ r, c: idxDate })]
                : undefined;
        const equipment =
            getXlsxCellDisplayText(eqCell) ||
            cellText(rows[r]?.[idxEquipment]);
        if (!equipment) continue;

        const countReadng = getXlsxCellDisplayText(cntCell);
        const date =
            getXlsxCellDisplayText(dateCell) ||
            (idxDate >= 0 ? cellText(rows[r]?.[idxDate]) : "");

        if (!byEquipment.has(equipment)) byEquipment.set(equipment, []);
        byEquipment.get(equipment).push({ countReadng, date });
    }

    return byEquipment;
}

/**
 * @param {Array<{ countReadng: string, date: string }> | undefined} entries
 * @param {string} szaeh
 * @returns {{ status: string, countReadng: string, tanggal: string }}
 */
function resolveCdsCheck(entries, szaeh) {
    if (!entries || entries.length === 0) {
        return { status: "cds_missing", countReadng: "", tanggal: "" };
    }

    if (entries.length === 1) {
        const { countReadng, date } = entries[0];
        const match = countersEqualExact(szaeh, countReadng);
        return {
            status: match ? "match" : "mismatch",
            countReadng,
            tanggal: match ? "" : date
        };
    }

    const canonicals = entries.map(e =>
        canonicalCounterForExactCompare(e.countReadng)
    );
    const allSame = canonicals.every(c => c === canonicals[0]);

    if (allSame) {
        const { countReadng, date } = entries[0];
        const match = countersEqualExact(szaeh, countReadng);
        return {
            status: match ? "match" : "mismatch",
            countReadng,
            tanggal: match ? "" : date
        };
    }

    const mismatchEntries = entries.filter(
        e => !countersEqualExact(szaeh, e.countReadng)
    );
    const tanggalDates = [
        ...new Set(
            mismatchEntries.map(e => cellText(e.date)).filter(Boolean)
        )
    ];
    const displayEntry = mismatchEntries[0] ?? entries[0];

    return {
        status: "mismatch",
        countReadng: displayEntry.countReadng,
        tanggal: tanggalDates.join("; ")
    };
}

/**
 * @param {string} localPath
 * @returns {Promise<Array<{ point: string, szaeh: string }>>}
 */
async function extractMaintenancePlanRowsFromPath(localPath) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(localPath);
    const ws = wb.worksheets[0];
    if (!ws) return [];

    const headerRow = buildHeaderRowFromSheet(ws, 1);
    const idxPoint = findColumnIndexByPrefix(headerRow, "POINT");
    const idxSzaeh = findColumnIndexByPrefix(headerRow, "SZAEH");

    if (idxPoint < 0 || idxSzaeh < 0) {
        return [];
    }

    const pointCol = idxPoint + 1;
    const szaehCol = idxSzaeh + 1;
    const lastRow = ws.rowCount || 0;

    /** @type {Array<{ point: string, szaeh: string }>} */
    const rows = [];

    for (let r = MAINTENANCE_PLAN_OUTPUT_START_ROW; r <= lastRow; r += 1) {
        const row = ws.getRow(r);
        const point = getExcelCellDisplayText(row.getCell(pointCol));
        const szaeh = getExcelCellDisplayText(row.getCell(szaehCol));
        if (!point || szaeh === "") continue;
        rows.push({ point, szaeh });
    }

    return rows;
}

/**
 * @param {string} regional
 * @param {string} pksName
 * @param {Array<{ point: string, szaeh: string }>} planRows
 * @param {Map<string, string[]>} pointToEquipments
 * @param {Map<string, Array<{ countReadng: string, date: string }>>} cdsByEquipment
 * @param {{ byEquipment: Map<string, string>, entries: Array<{ equipment: string, funcLoc: string }> }} equipmentFuncLocIndex
 */
function buildMaintenancePlanCheckRows(
    regional,
    pksName,
    planRows,
    pointToEquipments,
    cdsByEquipment,
    equipmentFuncLocIndex
) {
    /** @type {Array<object>} */
    const out = [];
    /** @type {Array<object>} */
    const invalidParents = [];
    /** @type {Set<string>} */
    const emittedEquipment = new Set();

    const funcLocByEquipment = equipmentFuncLocIndex.byEquipment;
    const funcLocEntries = equipmentFuncLocIndex.entries;

    for (const { point, szaeh } of planRows) {
        const equipments = pointToEquipments.get(point) || [];

        if (equipments.length === 0) {
            out.push({
                regional,
                pksName,
                measuringPoint: point,
                equipment: "",
                functionalLoc: "",
                szaeh,
                countReadng: "",
                tanggal: "",
                status: "no_equipment"
            });
            continue;
        }

        for (const equipment of equipments) {
            const eqKey = cellText(equipment);
            const cdsEntries = cdsByEquipment.get(eqKey);
            const { status, countReadng, tanggal } = resolveCdsCheck(
                cdsEntries,
                szaeh
            );
            const functionalLoc = funcLocByEquipment.get(eqKey) ?? "";

            const row = {
                regional,
                pksName,
                measuringPoint: point,
                equipment: eqKey,
                functionalLoc,
                szaeh,
                countReadng,
                tanggal,
                status
            };
            out.push(row);
            emittedEquipment.add(eqKey);

            if (status !== "match" && status !== "no_equipment") {
                invalidParents.push({
                    equipment: eqKey,
                    funcLoc: functionalLoc,
                    measuringPoint: point,
                    szaeh,
                    countReadng,
                    tanggal,
                    status
                });
            }
        }
    }

    /** @type {Set<string>} */
    const emittedChildEquipment = new Set();

    for (const parent of invalidParents) {
        if (!parent.funcLoc) continue;

        for (const { equipment, funcLoc } of funcLocEntries) {
            if (equipment === parent.equipment) continue;
            if (!isFuncLocDescendant(funcLoc, parent.funcLoc)) continue;
            if (emittedEquipment.has(equipment)) continue;
            if (emittedChildEquipment.has(equipment)) continue;

            emittedChildEquipment.add(equipment);
            out.push({
                regional,
                pksName,
                measuringPoint: parent.measuringPoint,
                equipment,
                functionalLoc: funcLoc,
                szaeh: parent.szaeh,
                countReadng: parent.countReadng,
                tanggal: parent.tanggal,
                status: parent.status
            });
        }
    }

    return out;
}

/**
 * @param {string} outPath
 * @param {Array<object>} rows
 */
async function writeCheckMaintenancePlanExcel(outPath, rows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Check", {
        views: [{ state: "frozen", ySplit: 1 }]
    });

    ws.addRow(CHECK_HEADERS);
    ws.getRow(1).font = { bold: true };

    for (const row of rows) {
        ws.addRow([
            row.regional,
            row.pksName,
            row.measuringPoint,
            row.equipment,
            row.functionalLoc,
            row.szaeh,
            row.countReadng,
            row.tanggal,
            row.status
        ]);
    }

    ws.columns = [
        { width: 18 },
        { width: 36 },
        { width: 20 },
        { width: 18 },
        { width: 36 },
        { width: 14 },
        { width: 14 },
        { width: 16 },
        { width: 14 }
    ];

    if (rows.length > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1 + rows.length, column: CHECK_HEADERS.length }
        };
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await wb.xlsx.writeFile(outPath);
}

module.exports = {
    CHECK_HEADERS,
    getExcelCellDisplayText,
    canonicalCounterForExactCompare,
    countersEqualExact,
    isFuncLocDescendant,
    buildPointToEquipmentsMap,
    loadCdsCountReadngIndex,
    loadCdsEntriesByEquipment,
    resolveCdsCheck,
    extractMaintenancePlanRowsFromPath,
    buildMaintenancePlanCheckRows,
    writeCheckMaintenancePlanExcel
};
