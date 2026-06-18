/**
 * Transformasi baris Protect Y → baris HEADER + OPERATION (LSMW tasklist).
 */

const fs = require("fs");
const path = require("path");
const { buildColumnIndexFirstWins } = require("./lsmw_lookups.cjs");
const { findDataLayout, normalizeHeader } = require("./lsmw_cell_fill.cjs");
const {
    findFunclocDescAfterColumnIndex,
    findCostCenterColumnIndex
} = require("./equipment_gathering_columns.cjs");

const JSON_PATH = path.join(
    __dirname,
    "..",
    "peralatan_dengan_jam_notif_overhaul.json"
);

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function isEmptyCell(value) {
    return value === null || value === undefined || String(value).trim() === "";
}

function extractCostCenterSuffixForCsv(costCenter) {
    const cleaned = String(costCenter ?? "")
        .replace(/[`]/g, "")
        .replace(/\s+/g, "")
        .trim()
        .toUpperCase();
    if (!cleaned) return "";

    const stasMatch = cleaned.match(/STAS(\d{2,3})/);
    if (stasMatch) {
        return `STAS${String(parseInt(stasMatch[1], 10)).padStart(2, "0")}`;
    }

    const noStars = cleaned.replace(/^\*{4}/, "");
    if (/^STAS\d{2}$/i.test(noStars)) return noStars.toUpperCase();

    if (/^[A-Z0-9]{4}/.test(noStars) && noStars.length > 4) {
        return noStars.slice(4).replace(/^\*{4}/, "");
    }

    return noStars;
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

function getFuncLocSecondSegment(raw) {
    const parts = String(raw ?? "")
        .split("-")
        .map(p => p.trim())
        .filter(Boolean);
    return parts.length >= 2 ? parts[1] : "";
}

function loadPreventiveMaintenanceRules() {
    const raw = JSON.parse(fs.readFileSync(JSON_PATH, "utf8"));
    const items = raw.preventive_maintenance ?? [];
    const byFilter = new Map();

    for (const item of items) {
        const filter = cellText(item.filter);
        if (!filter) {
            const label = cellText(item.text) || cellText(item.equipment) || "?";
            console.warn(
                `[preventive JSON] entri tanpa filter dilewati: ${label}`
            );
            continue;
        }
        const key = filter.toUpperCase();
        if (!byFilter.has(key)) byFilter.set(key, []);
        byFilter.get(key).push(item);
    }

    const filterKeywords = [...byFilter.keys()].sort(
        (a, b) => b.length - a.length
    );

    return { items, byFilter, filterKeywords };
}

/**
 * Longest filter keyword contained in desc (case-insensitive).
 * @returns {string|null} filter key (uppercase)
 */
function matchFilterKeyword(desc, filterKeywords) {
    const upper = String(desc ?? "").toUpperCase();
    for (const keyword of filterKeywords) {
        if (upper.includes(keyword)) return keyword;
    }
    return null;
}

/**
 * Pilih satu entri JSON untuk baris sumber (satu HEADER + dua OPERATION).
 */
function pickJsonEntryForRow(entries, desc) {
    if (!entries || entries.length === 0) return null;
    if (entries.length === 1) return entries[0];

    const upper = String(desc ?? "").toUpperCase();

    const hmMatches = entries.filter(entry => {
        const hm = String(entry.hour_meter ?? "");
        if (!hm) return false;
        if (upper.includes(`${hm} HM`)) return true;
        if (upper.includes(hm)) return true;
        return false;
    });

    if (hmMatches.length === 1) return hmMatches[0];
    if (hmMatches.length > 1) {
        return hmMatches.sort(
            (a, b) =>
                String(b.hour_meter).length - String(a.hour_meter).length
        )[0];
    }

    const textMatches = entries.filter(entry => {
        const t = cellText(entry.text).toUpperCase();
        return t && upper.includes(t);
    });
    if (textMatches.length === 1) return textMatches[0];

    return entries[0];
}

/**
 * Entri JSON untuk tasklist: satu entri, atau semua dengan hour_meter unik (multi HM).
 * @param {object[]} entries
 * @returns {object[]}
 */
function resolveJsonEntriesForFilter(entries) {
    if (!entries || entries.length === 0) return [];
    if (entries.length === 1) return [entries[0]];

    const seen = new Set();
    const result = [];
    for (const entry of entries) {
        const hm = entry.hour_meter;
        if (hm === null || hm === undefined || hm === "") continue;
        const key = String(hm);
        if (seen.has(key)) continue;
        seen.add(key);
        result.push(entry);
    }

    if (result.length > 0) return result;
    return [entries[0]];
}

function resolveTaskListGroupSuffix(plant, taskListGroup) {
    const group = cellText(taskListGroup).toUpperCase();
    const p = cellText(plant);
    if (group === "PRS1" && p.startsWith("3")) {
        return "PRS2";
    }
    return cellText(taskListGroup);
}

function buildTaskListGroup(plant, taskListGroup) {
    return `${cellText(plant)}${resolveTaskListGroupSuffix(plant, taskListGroup)}`;
}

function hourMeterColumnName(hourMeter) {
    return `${cellText(hourMeter)} HM`;
}

function normalizeCellValue(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "number" && Number.isFinite(value)) {
        return String(value);
    }
    return String(value).trim();
}

/**
 * @param {Record<string, unknown>} row
 * @param {string[]} columnOrder
 * @param {{ excludeColumns?: string[] }} [options]
 */
function rowFingerprint(row, columnOrder, options = {}) {
    const exclude = new Set(options.excludeColumns ?? []);
    const parts = [];
    for (const col of columnOrder) {
        if (exclude.has(col)) continue;
        parts.push(`${col}\x1f${normalizeCellValue(row[col])}`);
    }
    return parts.join("\x1e");
}

/**
 * @typedef {{
 *   headerRow: Record<string, unknown>,
 *   operationRows: Array<Record<string, unknown>>,
 *   operationRowBoldFlags: boolean[]
 * }} TasklistBundle
 */

/**
 * @param {TasklistBundle[]} bundles
 * @param {{ headerColumnOrder: string[], operationColumnOrder: string[] }} options
 */
function filterUniqueBundles(bundles, options) {
    const { headerColumnOrder, operationColumnOrder } = options;
    const noColumn = headerColumnOrder[0];
    const seenHeader = new Set();
    const seenOperation = new Set();
    const kept = [];
    let skippedDuplicateBundles = 0;

    for (const bundle of bundles) {
        const headerFp = rowFingerprint(bundle.headerRow, headerColumnOrder, {
            excludeColumns: noColumn ? [noColumn] : []
        });

        let duplicate = seenHeader.has(headerFp);
        const opFps = [];

        if (!duplicate) {
            for (const opRow of bundle.operationRows) {
                const opFp = rowFingerprint(opRow, operationColumnOrder);
                if (seenOperation.has(opFp)) {
                    duplicate = true;
                    break;
                }
                opFps.push(opFp);
            }
        }

        if (duplicate) {
            skippedDuplicateBundles += 1;
            continue;
        }

        seenHeader.add(headerFp);
        for (const opFp of opFps) seenOperation.add(opFp);
        kept.push(bundle);
    }

    return { bundles: kept, skippedDuplicateBundles };
}

/**
 * @param {TasklistBundle[]} bundles
 */
function bundlesToFlatRows(bundles) {
    const headerRows = [];
    const operationRows = [];
    const operationRowBoldFlags = [];

    for (const bundle of bundles) {
        headerRows.push(bundle.headerRow);
        operationRows.push(...bundle.operationRows);
        operationRowBoldFlags.push(...bundle.operationRowBoldFlags);
    }

    return { headerRows, operationRows, operationRowBoldFlags };
}

/** Tanggal run WIB untuk kolom STTAG (DD.MM.YYYY). */
function formatSttagWib(date = new Date()) {
    const parts = new Intl.DateTimeFormat("en-GB", {
        timeZone: "Asia/Jakarta",
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
    }).formatToParts(date);
    const get = type => parts.find(p => p.type === type)?.value ?? "00";
    return `${get("day")}.${get("month")}.${get("year")}`;
}

/**
 * Satu baris SUB_OPERATION per baris OPERATION tasklist (bold, operation_no 0010).
 * @param {{
 *   operationRows: Array<Record<string, unknown>>,
 *   operationRowBoldFlags: boolean[],
 *   operationColumnOrder: string[],
 *   subOperationColumnOrder: string[],
 *   subOperationHourMeterColumns: string[],
 *   sttag: string
 * }} params
 */
function buildSubOperationRows(params) {
    const {
        operationRows,
        operationRowBoldFlags,
        operationColumnOrder,
        subOperationColumnOrder,
        subOperationHourMeterColumns,
        sttag
    } = params;

    const taskListGroupCol = operationColumnOrder[0];
    const groupCounterCol = operationColumnOrder[1];
    const operationNoCol = operationColumnOrder[2];
    const subOpRows = [];

    for (let i = 0; i < operationRows.length; i += 1) {
        if (!operationRowBoldFlags[i]) continue;

        const opRow = operationRows[i];
        if (operationNoCol && cellText(opRow[operationNoCol]) !== "0010") continue;
        const row = {};
        for (const col of subOperationColumnOrder) row[col] = null;

        if (taskListGroupCol) row["PLNNR"] = opRow[taskListGroupCol] ?? null;
        row["STTAG"] = sttag;
        if (groupCounterCol) row["PLNAL"] = opRow[groupCounterCol] ?? null;
        row["ENTRY_ACT"] = null;
        row["FLG_SEL_01"] = null;

        for (const hmCol of subOperationHourMeterColumns) {
            row[hmCol] = opRow[hmCol] === "X" ? "X" : null;
        }

        subOpRows.push(row);
    }

    return subOpRows;
}

function buildHeaderRowForMatch(params) {
    const {
        plant,
        jsonEntry,
        workCenter,
        plannerGroup,
        headerColumnOrder
    } = params;

    const taskListGroup = buildTaskListGroup(plant, jsonEntry.task_list_group);
    const equipment = cellText(jsonEntry.equipment);

    const headerRow = {};
    for (const col of headerColumnOrder) headerRow[col] = null;

    headerRow[headerColumnOrder[0]] = null; // No — diisi saat write (autoincrement)
    if (headerColumnOrder[1] !== undefined)
        headerRow[headerColumnOrder[1]] = taskListGroup;
    if (headerColumnOrder[2] !== undefined) headerRow[headerColumnOrder[2]] = "1";
    if (headerColumnOrder[3] !== undefined)
        headerRow[headerColumnOrder[3]] = `PREVENTIVE ${equipment}`;
    if (headerColumnOrder[4] !== undefined) headerRow[headerColumnOrder[4]] = plant;
    if (headerColumnOrder[5] !== undefined)
        headerRow[headerColumnOrder[5]] = workCenter;
    if (headerColumnOrder[6] !== undefined) headerRow[headerColumnOrder[6]] = plant;
    if (headerColumnOrder[7] !== undefined)
        headerRow[headerColumnOrder[7]] = plannerGroup;
    if (headerColumnOrder[8] !== undefined)
        headerRow[headerColumnOrder[8]] = "PALMPB";

    return headerRow;
}

/**
 * @returns {{ tasklistLine: Record<string, unknown>, operationLine: Record<string, unknown> }}
 */
function buildOperationPairForEntry(params) {
    const {
        plant,
        jsonEntry,
        workCenter,
        operationColumnOrder,
        hourMeterColumns
    } = params;

    const taskListGroup = buildTaskListGroup(plant, jsonEntry.task_list_group);
    const operation = jsonEntry.operations?.[0] ?? {};
    const hmColName = hourMeterColumnName(jsonEntry.hour_meter);

    function baseOperationRow(isTasklistLine) {
        const row = {};
        for (const col of operationColumnOrder) row[col] = null;

        if (operationColumnOrder[0] !== undefined)
            row[operationColumnOrder[0]] = taskListGroup;
        if (operationColumnOrder[1] !== undefined)
            row[operationColumnOrder[1]] = "1";

        if (isTasklistLine) {
            if (operationColumnOrder[2] !== undefined)
                row[operationColumnOrder[2]] = cellText(operation.operation_no);
            if (operationColumnOrder[3] !== undefined)
                row[operationColumnOrder[3]] = null;
            if (operationColumnOrder[4] !== undefined)
                row[operationColumnOrder[4]] = cellText(jsonEntry.text);
            if (operationColumnOrder[7] !== undefined)
                row[operationColumnOrder[7]] = cellText(jsonEntry.control_key);
        } else {
            if (operationColumnOrder[2] !== undefined)
                row[operationColumnOrder[2]] = cellText(operation.operation_no);
            if (operationColumnOrder[3] !== undefined)
                row[operationColumnOrder[3]] = cellText(
                    operation.sub_operation_no
                );
            if (operationColumnOrder[4] !== undefined)
                row[operationColumnOrder[4]] = cellText(
                    operation.operation_text
                );
            if (operationColumnOrder[7] !== undefined)
                row[operationColumnOrder[7]] = cellText(
                    operation.operation_control_key
                );
        }

        if (operationColumnOrder[5] !== undefined)
            row[operationColumnOrder[5]] = workCenter;
        if (operationColumnOrder[6] !== undefined)
            row[operationColumnOrder[6]] = plant;

        for (const hmCol of hourMeterColumns) {
            row[hmCol] = isTasklistLine && hmCol === hmColName ? "X" : null;
        }

        return row;
    }

    return {
        tasklistLine: baseOperationRow(true),
        operationLine: baseOperationRow(false)
    };
}

/**
 * @param {object} params
 * @returns {{
 *   headerRow: Record<string, unknown>,
 *   operationRows: Array<Record<string, unknown>>,
 *   operationRowBoldFlags: boolean[]
 * }}
 */
function buildOutputRowsForMatch(params) {
    const {
        plant,
        jsonEntries,
        workCenter,
        plannerGroup,
        headerColumnOrder,
        operationColumnOrder,
        hourMeterColumns
    } = params;

    const headerRow = buildHeaderRowForMatch({
        plant,
        jsonEntry: jsonEntries[0],
        workCenter,
        plannerGroup,
        headerColumnOrder
    });

    const operationRows = [];
    const operationRowBoldFlags = [];

    for (const jsonEntry of jsonEntries) {
        const pair = buildOperationPairForEntry({
            plant,
            jsonEntry,
            workCenter,
            operationColumnOrder,
            hourMeterColumns
        });
        operationRows.push(pair.tasklistLine, pair.operationLine);
        operationRowBoldFlags.push(true, false);
    }

    return { headerRow, operationRows, operationRowBoldFlags };
}

/**
 * @param {Array<Array<unknown>>} rawRows
 * @param {{
 *   seenFuncLocKeys?: Set<string>,
 *   workCenterByRowIndex?: Map<number, unknown>,
 *   plannerGroupByRowIndex?: Map<number, unknown>,
 *   headerColumnOrder: string[],
 *   operationColumnOrder: string[],
 *   hourMeterColumns: string[],
 *   requireUniqueInFile?: boolean
 * }} options
 */
function buildTasklistRows(rawRows, options) {
    const layout = findDataLayout(rawRows);
    const headerRow = rawRows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header tidak ditemukan" };
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxDesc = findFunclocDescAfterColumnIndex(headerRow);
    const idxCostCenter = findCostCenterColumnIndex(headerRow);

    if (idxFuncLoc === undefined) {
        return {
            ok: false,
            reason: "Kolom FUNCTIONAL LOCATION AFTER tidak ditemukan"
        };
    }

    if (idxDesc === undefined) {
        return {
            ok: false,
            reason: "Kolom FUNCTLOC DESC. AFTER tidak ditemukan"
        };
    }

    const { byFilter, filterKeywords } = loadPreventiveMaintenanceRules();
    const dataRows = rawRows.slice(layout.headerRowIndex + 1);

    /** @type {TasklistBundle[]} */
    const bundles = [];

    for (let i = 0; i < dataRows.length; i++) {
        const r = dataRows[i] || [];
        const funcLoc = cellText(r[idxFuncLoc]);
        if (!funcLoc) continue;

        const desc = cellText(r[idxDesc]);

        const funcLocKey = normalizeFuncLocKey(funcLoc);
        if (funcLocKey && options.seenFuncLocKeys?.has(funcLocKey)) continue;

        const workCenter = options.workCenterByRowIndex?.get(i);

        const filterKey = matchFilterKeyword(desc, filterKeywords);
        if (!filterKey) continue;

        const jsonEntries = resolveJsonEntriesForFilter(
            byFilter.get(filterKey)
        );
        if (jsonEntries.length === 0) continue;

        const plannerGroup = options.plannerGroupByRowIndex?.get(i) ?? null;
        const missingFields = [];
        if (isEmptyCell(workCenter)) missingFields.push("main work center");
        if (isEmptyCell(plannerGroup)) missingFields.push("planner group");
        if (missingFields.length > 0) {
            const costCenterRaw =
                idxCostCenter !== undefined ? r[idxCostCenter] ?? null : null;
            const effectiveCostCenter =
                options.effectiveCostCenterByRowIndex?.get(i) ?? null;
            return {
                ok: false,
                reason: "MISSING_CWC_PLANNER",
                validationError: {
                    functionalLocation: funcLoc,
                    functionalLocationDescription: desc,
                    missingFields,
                    costCenterRaw,
                    effectiveCostCenter,
                    costCenterSuffixForCsv: extractCostCenterSuffixForCsv(
                        effectiveCostCenter ?? costCenterRaw
                    )
                }
            };
        }

        if (funcLocKey) options.seenFuncLocKeys?.add(funcLocKey);

        const plant = getFuncLocSecondSegment(funcLoc);

        const built = buildOutputRowsForMatch({
            plant,
            jsonEntries,
            workCenter,
            plannerGroup,
            headerColumnOrder: options.headerColumnOrder,
            operationColumnOrder: options.operationColumnOrder,
            hourMeterColumns: options.hourMeterColumns
        });

        bundles.push({
            headerRow: built.headerRow,
            operationRows: built.operationRows,
            operationRowBoldFlags: built.operationRowBoldFlags
        });
    }

    let skippedDuplicateBundles = 0;
    let finalBundles = bundles;

    if (options.requireUniqueInFile) {
        const filtered = filterUniqueBundles(bundles, {
            headerColumnOrder: options.headerColumnOrder,
            operationColumnOrder: options.operationColumnOrder
        });
        finalBundles = filtered.bundles;
        skippedDuplicateBundles = filtered.skippedDuplicateBundles;
    }

    const flat = bundlesToFlatRows(finalBundles);

    return {
        ok: true,
        bundles: finalBundles,
        ...flat,
        skippedDuplicateBundles
    };
}

module.exports = {
    loadPreventiveMaintenanceRules,
    matchFilterKeyword,
    pickJsonEntryForRow,
    resolveJsonEntriesForFilter,
    resolveTaskListGroupSuffix,
    buildTaskListGroup,
    buildTasklistRows,
    buildSubOperationRows,
    formatSttagWib,
    rowFingerprint,
    filterUniqueBundles,
    bundlesToFlatRows,
    getFuncLocSecondSegment,
    hourMeterColumnName
};
