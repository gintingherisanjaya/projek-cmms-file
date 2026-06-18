/**
 * Deteksi duplikat EQUNR per file (template_changes_equipment.xlsx layout).
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const OUTPUT_START_ROW = 4;
const HEADER_ROW = 1;

const RED_FILL = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFF0000" }
};

function normalizeHeader(name) {
    if (!name) return "";
    return String(name).trim().toUpperCase();
}

function normalizeEquipmentNumber(v) {
    if (v === null || v === undefined || v === "") return "";
    return String(v).trim();
}

function cellText(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) return value.toISOString();
    if (typeof value === "object" && value !== null) {
        if (value.text !== undefined) return String(value.text).trim();
        if (value.result !== undefined) return String(value.result).trim();
        if (value.richText) {
            return value.richText.map(t => t.text ?? "").join("").trim();
        }
    }
    return String(value).trim();
}

function getLastRowNumber(worksheet) {
    return worksheet.rowCount || worksheet.lastRow?.number || 0;
}

function findColumnIndexInRow(worksheet, rowNum, headerName) {
    const row = worksheet.getRow(rowNum);
    const target = normalizeHeader(headerName);
    const lastCol = Math.max(row.cellCount || 0, 10);

    for (let c = 1; c <= lastCol; c += 1) {
        const header = normalizeHeader(cellText(row.getCell(c).value));
        if (header === target) return c;
    }

    return undefined;
}

/**
 * @param {import('exceljs').Worksheet} worksheet
 * @param {string} [fileLabel]
 * @returns {number} 1-based column index EQUNR
 */
function assertEqunrHeaderRow1(worksheet, fileLabel = "") {
    const equnrColumn = findColumnIndexInRow(worksheet, HEADER_ROW, "EQUNR");
    if (equnrColumn === undefined) {
        const suffix = fileLabel ? `: ${fileLabel}` : "";
        throw new Error(`Kolom EQUNR tidak ditemukan di baris 1${suffix}`);
    }
    return equnrColumn;
}

/**
 * @param {import('exceljs').Worksheet} worksheet
 * @param {number} equnrColumn
 * @returns {Array<{ oldName: string, equnr: string, tplnr: string, newName: string }>}
 */
function readChangesEquipmentRows(worksheet, equnrColumn) {
    const oldNameCol =
        findColumnIndexInRow(worksheet, HEADER_ROW, "OLD_NAME") ?? 1;
    const tplnrCol = findColumnIndexInRow(worksheet, HEADER_ROW, "TPLNR") ?? 3;
    const newNameCol =
        findColumnIndexInRow(worksheet, HEADER_ROW, "NEW_NAME") ?? 4;

    const rows = [];
    const lastRow = getLastRowNumber(worksheet);

    for (let rowNum = OUTPUT_START_ROW; rowNum <= lastRow; rowNum += 1) {
        const row = worksheet.getRow(rowNum);
        const equnr = normalizeEquipmentNumber(
            cellText(row.getCell(equnrColumn).value)
        );
        if (!equnr) continue;

        rows.push({
            oldName: cellText(row.getCell(oldNameCol).value),
            equnr,
            tplnr: cellText(row.getCell(tplnrCol).value),
            newName: cellText(row.getCell(newNameCol).value)
        });
    }

    return rows;
}

/**
 * @param {Array<{ equnr: string }>} rows
 * @returns {{ duplicateEqunrCount: number, duplicateRowIndexes: Set<number> }}
 */
function findDuplicateRowIndexes(rows) {
    /** @type {Map<string, number[]>} */
    const byEqunr = new Map();

    rows.forEach((row, index) => {
        const list = byEqunr.get(row.equnr) ?? [];
        list.push(index);
        byEqunr.set(row.equnr, list);
    });

    const duplicateRowIndexes = new Set();
    let duplicateEqunrCount = 0;

    for (const [, indexes] of byEqunr) {
        if (indexes.length <= 1) continue;
        duplicateEqunrCount += 1;
        for (const index of indexes) {
            duplicateRowIndexes.add(index);
        }
    }

    return { duplicateEqunrCount, duplicateRowIndexes };
}

/**
 * @param {import('exceljs').Worksheet} worksheet
 * @param {Set<number>} rowNumbers 1-based Excel row numbers
 * @param {number} [colCount]
 */
function applyRedFillToRows(worksheet, rowNumbers, colCount = 4) {
    for (const rowNum of rowNumbers) {
        const row = worksheet.getRow(rowNum);
        for (let c = 1; c <= colCount; c += 1) {
            row.getCell(c).fill = RED_FILL;
        }
        row.commit();
    }
}

function cellHasRedFill(cell) {
    if (!cell || !cell.fill) return false;
    const argb = cell.fill.fgColor?.argb;
    return String(argb ?? "").toUpperCase() === "FFFF0000";
}

/**
 * Tulis ulang dari template bersih; hanya baris EQUNR duplikat di-fill merah.
 * @param {string} templatePath
 * @param {Array<{ oldName: string, equnr: string, tplnr: string, newName: string }>} rows
 * @param {Set<number>} duplicateRowIndexes 0-based index ke `rows`
 * @param {string} outPath
 */
async function writeChangesEquipmentCheckWorkbook(
    templatePath,
    rows,
    duplicateRowIndexes,
    outPath
) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const sheet = workbook.worksheets[0];
    if (!sheet) {
        throw new Error("Sheet not found in template");
    }

    sheet.conditionalFormattings = [];

    const sampleRow = sheet.getRow(3);
    for (let c = 1; c <= 4; c += 1) {
        sampleRow.getCell(c).value = null;
    }
    sampleRow.commit();

    const duplicateExcelRows = new Set();
    let startRow = OUTPUT_START_ROW;

    for (let index = 0; index < rows.length; index += 1) {
        const excelRow = startRow;
        const rowData = rows[index];
        const row = sheet.getRow(startRow);
        row.getCell(1).value = String(rowData.oldName);
        row.getCell(2).value = String(rowData.equnr);
        row.getCell(3).value = String(rowData.tplnr);
        row.getCell(4).value = String(rowData.newName);
        row.commit();

        if (duplicateRowIndexes.has(index)) {
            duplicateExcelRows.add(excelRow);
        }

        startRow += 1;
    }

    if (duplicateExcelRows.size > 0) {
        applyRedFillToRows(sheet, duplicateExcelRows);
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await workbook.xlsx.writeFile(outPath);
}

module.exports = {
    OUTPUT_START_ROW,
    HEADER_ROW,
    RED_FILL,
    normalizeHeader,
    normalizeEquipmentNumber,
    cellText,
    assertEqunrHeaderRow1,
    readChangesEquipmentRows,
    findDuplicateRowIndexes,
    applyRedFillToRows,
    writeChangesEquipmentCheckWorkbook,
    cellHasRedFill
};

if (require.main === module) {
    const os = require("os");

    (async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Test");
        ws.getRow(1).getCell(1).value = "OLD_NAME";
        ws.getRow(1).getCell(2).value = "EQUNR";
        ws.getRow(1).getCell(3).value = "TPLNR";
        ws.getRow(1).getCell(4).value = "NEW_NAME";
        ws.getRow(4).getCell(2).value = "100";
        ws.getRow(5).getCell(2).value = "100";
        ws.getRow(6).getCell(2).value = "200";

        assertEqunrHeaderRow1(ws);
        const rows = readChangesEquipmentRows(ws, 2);
        const dup = findDuplicateRowIndexes(rows);

        if (rows.length !== 3 || dup.duplicateEqunrCount !== 1) {
            console.error("FAIL duplicate detection", rows.length, dup);
            process.exit(1);
        }

        const templatePath = path.join(
            __dirname,
            "..",
            "template_changes_equipment.xlsx"
        );
        const outPath = path.join(os.tmpdir(), "check_lsmw_dup_smoke.xlsx");
        await writeChangesEquipmentCheckWorkbook(
            templatePath,
            rows,
            dup.duplicateRowIndexes,
            outPath
        );

        const outWb = new ExcelJS.Workbook();
        await outWb.xlsx.readFile(outPath);
        const outWs = outWb.worksheets[0];

        for (const rowNum of [4, 5]) {
            for (let c = 1; c <= 4; c += 1) {
                if (!cellHasRedFill(outWs.getRow(rowNum).getCell(c))) {
                    console.error(`FAIL row ${rowNum} col ${c} should be red`);
                    process.exit(1);
                }
            }
        }
        for (let c = 1; c <= 4; c += 1) {
            if (cellHasRedFill(outWs.getRow(6).getCell(c))) {
                console.error(`FAIL row 6 col ${c} should not be red`);
                process.exit(1);
            }
        }

        fs.unlinkSync(outPath);

        try {
            assertEqunrHeaderRow1(ws.getRow(1).worksheet || ws, "bad");
        } catch {
            // ws has EQUNR - test missing header separately
        }

        const badWb = new ExcelJS.Workbook();
        const badWs = badWb.addWorksheet("Bad");
        badWs.getRow(1).getCell(1).value = "OLD_NAME";
        let threw = false;
        try {
            assertEqunrHeaderRow1(badWs, "missing-equnr.xlsx");
        } catch (err) {
            threw = true;
            if (!/EQUNR tidak ditemukan/i.test(err.message)) {
                console.error("FAIL bad header message", err.message);
                process.exit(1);
            }
        }
        if (!threw) {
            console.error("FAIL expected throw for missing EQUNR header");
            process.exit(1);
        }

        console.log("OK check_lsmw_changes_equipment_dup smoke");
    })().catch(err => {
        console.error(err);
        process.exit(1);
    });
}
