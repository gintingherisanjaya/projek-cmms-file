/**
 * Pengecekan panjang karakter: file sumber PKS & output LSMW vs template.
 * Tanpa alias / transform (murni nilai di sel).
 */

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");
const { findDataLayout } = require("./lsmw_cell_fill.cjs");

const SOURCE_ALL_COLUMNS = [
    { templateCol: "TPLNR", sourceKeys: ["FUNCTIONAL LOCATION AFTER"] },
    { templateCol: "HERST", sourceKeys: ["MERK"] },
    { templateCol: "GROES", sourceKeys: ["CAPACITY"] },
    {
        templateCol: "BAUJJ",
        sourceKeys: ["STANDART UMUR TEKNIS (TAHUN)", "CONSTRUCTION YEAR"]
    },
    { templateCol: "EQART", sourceKeys: ["EQUIPMENT GROUP AFTER"] },
    { templateCol: "KOSTL*", sourceKeys: ["COST CENTER AFTER"] },
    {
        templateCol: "SWERK*",
        sourceKeys: ["MAINTENANCE PLAN AFTER", "MAINTENANCE PLANT", "POM"]
    },
    {
        templateCol: "BEGRU*",
        sourceKeys: ["MAINTENANCE PLAN AFTER", "MAINTENANCE PLANT", "POM"]
    },
    {
        templateCol: "IWERK*",
        sourceKeys: ["MAINTENANCE PLAN AFTER", "MAINTENANCE PLANT", "POM"]
    },
    {
        templateCol: "WERGW",
        sourceKeys: ["MAINTENANCE PLAN AFTER", "MAINTENANCE PLANT", "POM"]
    },
    { templateCol: "INGRP", sourceKeys: ["PLANNER GROUP", "PLANNER GROUP AFTER"] },
    {
        templateCol: "GEWRK",
        sourceKeys: ["WORK CENTER", "WORK CENTER AFTER"]
    }
];

function normalizeHeader(name) {
    return String(name ?? "")
        .trim()
        .toUpperCase();
}

function cellPlainText(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) return value.toISOString();
    if (typeof value === "object" && value !== null) {
        if (value.text !== undefined) return String(value.text);
        if (value.result !== undefined) return String(value.result);
        if (value.richText) {
            return value.richText.map(t => t.text ?? "").join("");
        }
    }
    return String(value);
}

function parseMaxLength(raw) {
    if (raw === null || raw === undefined || raw === "") return null;
    const n = parseInt(String(raw).trim(), 10);
    return Number.isFinite(n) && n > 0 ? n : null;
}

function isFunclocDescAfterHeader(header) {
    const h = normalizeHeader(header);
    if (!h) return false;
    if (/\bBEFORE\b/.test(h)) return false;
    if (!h.includes("FUNCTLOC")) return false;
    if (!/\bDESC\b/.test(h)) return false;
    if (!/\b(AFTER|AFTRER)\b/.test(h)) return false;
    return true;
}

function findFunclocDescAfterColumnIndex(headerRow) {
    const indices = [];
    for (let i = 0; i < headerRow.length; i++) {
        if (isFunclocDescAfterHeader(headerRow[i])) indices.push(i);
    }
    if (indices.length === 0) return undefined;
    if (indices.length === 1) return indices[0];
    const withLevel = indices.filter(i =>
        /\bLEVEL\b/.test(normalizeHeader(headerRow[i]))
    );
    if (withLevel.length >= 1) return withLevel[0];
    return indices[0];
}

async function loadTemplateLimits(
    templatePath,
    headerRowNum,
    maxLenRowNum
) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const sheet = wb.worksheets[0];
    if (!sheet) throw new Error("Template tidak memiliki sheet");

    const nameRow = sheet.getRow(headerRowNum);
    const maxRow = sheet.getRow(maxLenRowNum);
    const colCount = Math.max(nameRow.cellCount, maxRow.cellCount, 30);
    const columns = [];
    const maxLengths = [];

    for (let c = 1; c <= colCount; c += 1) {
        const name = cellPlainText(nameRow.getCell(c).value).trim();
        if (!name) break;
        columns.push(name);
        maxLengths.push(parseMaxLength(maxRow.getCell(c).value));
    }

    if (columns.length === 0) {
        throw new Error(
            `Kolom template tidak ditemukan (baris ${headerRowNum})`
        );
    }

    return { columns, maxLengths, limitByTemplate: buildLimitMap(columns, maxLengths) };
}

function buildLimitMap(columns, maxLengths) {
    const limitByTemplate = {};
    for (let i = 0; i < columns.length; i += 1) {
        limitByTemplate[columns[i]] = maxLengths[i];
    }
    return limitByTemplate;
}

function buildSourceChecks(headerRow, limitByTemplate, scope, descTemplateCol) {
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxFunclocDesc = findFunclocDescAfterColumnIndex(headerRow);
    const checks = [];

    if (scope === "all-source-columns") {
        for (const spec of SOURCE_ALL_COLUMNS) {
            const maxLen = limitByTemplate[spec.templateCol];
            if (!maxLen) continue;
            for (const key of spec.sourceKeys) {
                const idx = colIndex[normalizeHeader(key)];
                if (idx !== undefined) {
                    checks.push({
                        templateCol: spec.templateCol,
                        sourceIdx: idx,
                        sourceLabel:
                            cellPlainText(headerRow[idx]).trim() || key
                    });
                    break;
                }
            }
        }
        if (idxFunclocDesc !== undefined && limitByTemplate["SHTXT*"]) {
            const hasShtxt = checks.some(c => c.templateCol === "SHTXT*");
            if (!hasShtxt) {
                checks.push({
                    templateCol: "SHTXT*",
                    sourceIdx: idxFunclocDesc,
                    sourceLabel:
                        cellPlainText(headerRow[idxFunclocDesc]).trim() ||
                        "FUNCTLOC DESC AFTER"
                });
            }
        }
    } else if (scope === "funcloc-desc-only" && descTemplateCol) {
        const maxLen = limitByTemplate[descTemplateCol];
        if (idxFunclocDesc !== undefined && maxLen) {
            checks.push({
                templateCol: descTemplateCol,
                sourceIdx: idxFunclocDesc,
                sourceLabel:
                    cellPlainText(headerRow[idxFunclocDesc]).trim() ||
                    "FUNCTLOC DESC AFTER"
            });
        }
    }

    return { checks, idxFuncLoc, colIndex };
}

function evaluateRowViolations(r, checks, limitByTemplate, idxFuncLoc) {
    const functionalLocation =
        idxFuncLoc !== undefined
            ? cellPlainText(r[idxFuncLoc]).trim()
            : "";

    const columnViolations = [];
    for (const chk of checks) {
        const text = cellPlainText(r[chk.sourceIdx]).trim();
        if (!text) continue;
        const max = limitByTemplate[chk.templateCol];
        const len = text.length;
        if (max && len > max) {
            columnViolations.push({
                column: chk.templateCol,
                sourceColumn: chk.sourceLabel,
                maxLength: max,
                actualLength: len,
                value: text
            });
        }
    }

    if (columnViolations.length === 0) return null;

    return {
        functionalLocation: functionalLocation || null,
        columnsOverLimit: columnViolations
    };
}

function scanSourceWorkbook(
    sourcePath,
    limitByTemplate,
    scope,
    descTemplateCol
) {
    const workbook = XLSX.readFile(sourcePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const layout = findDataLayout(rows);
    const headerRow = rows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header sumber tidak ditemukan" };
    }

    const dataRows = rows.slice(layout.headerRowIndex + 1);
    const { checks, idxFuncLoc } = buildSourceChecks(
        headerRow,
        limitByTemplate,
        scope,
        descTemplateCol
    );

    const violatingRows = [];

    for (let i = 0; i < dataRows.length; i++) {
        const hit = evaluateRowViolations(
            dataRows[i],
            checks,
            limitByTemplate,
            idxFuncLoc
        );
        if (!hit) continue;
        violatingRows.push({
            target: "source",
            sourceExcelRow: layout.dataStartExcelRow + i,
            outputExcelRow: null,
            ...hit
        });
    }

    return {
        ok: true,
        headerRowIndex: layout.headerRowIndex,
        dataStartExcelRow: layout.dataStartExcelRow,
        totalRows: dataRows.length,
        checksApplied: checks.map(c => ({
            templateColumn: c.templateCol,
            sourceColumn: c.sourceLabel
        })),
        violatingRows
    };
}

async function scanOutputLsmwFile(
    outputPath,
    modeConfig,
    limitByTemplate
) {
    const descCol = modeConfig.outputDescColumn;
    const keyCol = modeConfig.outputKeyColumn;
    const maxLen = limitByTemplate[descCol];
    if (!maxLen) {
        return { ok: false, reason: `Batas template ${descCol} tidak ada` };
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(outputPath);
    const sheet = wb.worksheets[0];
    if (!sheet) return { ok: false, reason: "Sheet output kosong" };

    const headerRow = sheet.getRow(modeConfig.outputHeaderRow);
    const columns = [];
    const colIndexByName = {};
    for (let c = 1; c <= headerRow.cellCount; c += 1) {
        const name = cellPlainText(headerRow.getCell(c).value).trim();
        if (!name) break;
        columns.push(name);
        colIndexByName[name] = c;
    }

    const idxDesc = colIndexByName[descCol];
    const idxKey = colIndexByName[keyCol];
    if (idxDesc === undefined) {
        return {
            ok: false,
            reason: `Kolom ${descCol} tidak ada di output`
        };
    }

    const violatingRows = [];
    let rowNum = modeConfig.outputDataStartRow;
    let totalRows = 0;

    while (rowNum <= sheet.rowCount) {
        const row = sheet.getRow(rowNum);
        const descText = cellPlainText(row.getCell(idxDesc).value).trim();
        const keyText =
            idxKey !== undefined
                ? cellPlainText(row.getCell(idxKey).value).trim()
                : "";

        const rowEmpty =
            !descText &&
            !keyText &&
            columns.every(name => {
                const ci = colIndexByName[name];
                return !cellPlainText(row.getCell(ci).value).trim();
            });

        if (rowEmpty && totalRows > 0) break;

        if (descText || keyText) {
            totalRows += 1;
            const len = descText.length;
            if (len > maxLen) {
                violatingRows.push({
                    target: "output",
                    sourceExcelRow: null,
                    outputExcelRow: rowNum,
                    functionalLocation: keyText || null,
                    columnsOverLimit: [
                        {
                            column: descCol,
                            sourceColumn: `${descCol} (output LSMW)`,
                            maxLength: maxLen,
                            actualLength: len,
                            value: descText
                        }
                    ]
                });
            }
        }

        rowNum += 1;
        if (rowNum > modeConfig.outputDataStartRow + 500000) break;
    }

    return {
        ok: true,
        totalRows,
        outputPath,
        checksApplied: [
            {
                templateColumn: descCol,
                sourceColumn: `${descCol} (output LSMW)`
            }
        ],
        violatingRows
    };
}

function findLocalOutputFile(outputRoot, sourceFileName) {
    if (!outputRoot || !fs.existsSync(outputRoot)) return null;

    const hits = [];
    const walk = dir => {
        for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
            const full = path.join(dir, ent.name);
            if (ent.isDirectory()) walk(full);
            else if (ent.name === sourceFileName) hits.push(full);
        }
    };
    walk(outputRoot);
    return hits.length > 0 ? hits[0] : null;
}

/**
 * @param {object} modeConfig from template_lsmw_check_config
 * @param {string} sourcePath downloaded source xlsx
 * @param {string} sourceFileName basename for output lookup
 */
async function runTemplateCheck(modeConfig, sourcePath, sourceFileName) {
    const { columns, maxLengths, limitByTemplate } =
        await loadTemplateLimits(
            modeConfig.templatePath,
            modeConfig.templateHeaderRow,
            modeConfig.templateMaxLenRow
        );

    const sourceScan = scanSourceWorkbook(
        sourcePath,
        limitByTemplate,
        modeConfig.scope,
        modeConfig.sourceDescTemplateCol
    );

    if (!sourceScan.ok) {
        return { ok: false, reason: sourceScan.reason };
    }

    const violatingRows = [...sourceScan.violatingRows];
    let outputScan = null;
    let outputPath = null;

    if (modeConfig.outputRoot) {
        outputPath = findLocalOutputFile(
            modeConfig.outputRoot,
            sourceFileName
        );
        if (outputPath) {
            outputScan = await scanOutputLsmwFile(
                outputPath,
                modeConfig,
                limitByTemplate
            );
            if (outputScan.ok) {
                violatingRows.push(...outputScan.violatingRows);
            }
        }
    }

    return {
        ok: true,
        template: {
            path: modeConfig.templatePath,
            headerRow: modeConfig.templateHeaderRow,
            maxLengthRow: modeConfig.templateMaxLenRow,
            columns,
            maxLengths
        },
        sourceScan,
        outputScan,
        outputPath,
        violatingRows
    };
}

module.exports = {
    cellPlainText,
    loadTemplateLimits,
    scanSourceWorkbook,
    scanOutputLsmwFile,
    findLocalOutputFile,
    runTemplateCheck
};
