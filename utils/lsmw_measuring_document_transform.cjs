/**
 * Transformasi baris Protect Y + data HERI → baris output LSMW measuring document.
 */

const {
    buildColumnIndexFirstWins,
    shouldSkipLevel5Stas13Row
} = require("./lsmw_lookups.cjs");
const { findDataLayout, normalizeHeader } = require("./lsmw_cell_fill.cjs");
const { findFunclocDescAfterColumnIndex } = require("./equipment_gathering_columns.cjs");
const { getFuncLocSecondSegment } = require("./lsmw_measuring_point_transform.cjs");
const {
    lookupEquipmentNumber,
    buildMaintenanceItemFallbackEqunr
} = require("./lsmw_maintenance_item_transform.cjs");
const {
    matchHeriEquipmentName,
    formatHeriCounterOutput,
    isSterilizerHeriAlias,
    resolveHeriCounterForEquipment
} = require("./heri_sheet_loader.cjs");
const { validateMeasuringDocumentFields } = require("./lsmw_measuring_document_limits.cjs");
const {
    fixFunclocDescSourceTypos,
    applyFunclocDescAlias
} = require("./funcloc_desc_alias.cjs");

const MEASUREMENT_DATE = "09.06.2026";
const MEASUREMENT_TIME = "10:00:00";

function cellText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

function funcLocLevel(code) {
    if (!code || typeof code !== "string") return 0;
    return code.trim().split("-").filter(s => s.length > 0).length;
}

function isExcludedDescription(desc) {
    const d = String(desc ?? "");
    if (/drive\s*unit/i.test(d)) return true;
    if (/accessories\s+weigh\s*bridge/i.test(d)) return true;
    return false;
}

function isMeasuringDocumentSourceRow(funcLoc, desc) {
    if (funcLocLevel(funcLoc) < 5) return false;
    if (isExcludedDescription(desc)) return false;
    return true;
}

function isFuncLocDescendant(childFuncLoc, ancestorFuncLoc) {
    const child = normalizeFuncLocKey(childFuncLoc);
    const ancestor = normalizeFuncLocKey(ancestorFuncLoc);
    return Boolean(child && ancestor && child.startsWith(ancestor + "-"));
}

function descForEquipmentOutput(rawDesc) {
    return applyFunclocDescAlias(fixFunclocDescSourceTypos(rawDesc));
}

/**
 * @param {string[]} columnOrder
 * @param {string} prefix
 */
function findColumnByPrefix(columnOrder, prefix) {
    const upper = prefix.toUpperCase();
    return (
        columnOrder.find(c => c.toUpperCase().startsWith(upper)) ??
        columnOrder.find(c => c.toUpperCase().includes(upper)) ??
        null
    );
}

/**
 * @param {string[]} columnOrder
 * @param {{
 *   seqNo: number,
 *   plant: string,
 *   desc: string,
 *   counterValue?: number | null,
 *   equipmentByPlantDesc?: Map<string, string>,
 *   pointByEquipment?: Map<string, string>
 * }} params
 */
function buildOutputRow(columnOrder, params) {
    const {
        seqNo,
        plant,
        desc,
        counterValue,
        equipmentByPlantDesc,
        pointByEquipment
    } = params;

    const mappedEqunr = lookupEquipmentNumber(plant, desc, equipmentByPlantDesc);
    const equnr = mappedEqunr || buildMaintenanceItemFallbackEqunr(plant, desc);
    const point =
        equnr && pointByEquipment ? pointByEquipment.get(equnr) ?? "" : "";

    const colNoUrut = findColumnByPrefix(columnOrder, "No. Urut");
    const colEqunr = findColumnByPrefix(columnOrder, "Equipment Number");
    const colDesc = findColumnByPrefix(columnOrder, "Equipment Description");
    const colPoint = findColumnByPrefix(columnOrder, "Measuring point");
    const colDate = findColumnByPrefix(columnOrder, "Measurement Date");
    const colTime = findColumnByPrefix(columnOrder, "Measurement Time");
    const colCounter = findColumnByPrefix(columnOrder, "Counter Reading");
    const colUom = findColumnByPrefix(columnOrder, "Unit of Measure");
    const colDiff = findColumnByPrefix(columnOrder, "Difference");
    const colReadBy = findColumnByPrefix(columnOrder, "Read By");
    const colShortText = findColumnByPrefix(columnOrder, "Short Text");

    const row = {};
    if (colNoUrut) row[colNoUrut] = seqNo;
    if (colEqunr) row[colEqunr] = equnr;
    if (colDesc) row[colDesc] = descForEquipmentOutput(desc);
    if (colPoint) row[colPoint] = point || null;
    if (colDate) row[colDate] = MEASUREMENT_DATE;
    if (colTime) row[colTime] = MEASUREMENT_TIME;
    if (colCounter) row[colCounter] = formatHeriCounterOutput(counterValue);
    if (colUom) row[colUom] = null;
    if (colDiff) row[colDiff] = "x";
    if (colReadBy) row[colReadBy] = "regional";
    if (colShortText) row[colShortText] = "IB sd 31 Mei 2026";

    return { row, equnr, point, equnrFromMapping: Boolean(mappedEqunr) };
}

/**
 * @param {Array<Array<unknown>>} rawRows
 * @param {string} fileName
 * @param {string[]} columnOrder
 * @param {{
 *   heriData: object,
 *   regional?: string,
 *   redEquipGroupRowIndices?: Set<number>,
 *   seenFuncLocKeys?: Set<string>,
 *   effectiveCostCenterByRowIndex?: Map<number, unknown>,
 *   equipmentByPlantDesc?: Map<string, string>,
 *   pointByEquipment?: Map<string, string>,
 *   usedPlantEquipmentKeys?: Set<string>
 * }} options
 */
function buildMeasuringDocumentRows(rawRows, fileName, columnOrder, options = {}) {
    const layout = findDataLayout(rawRows);
    const headerRow = rawRows[layout.headerRowIndex];
    if (!headerRow) {
        return { ok: false, reason: "Header tidak ditemukan" };
    }

    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxDesc = findFunclocDescAfterColumnIndex(headerRow);

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

    const heriData = options.heriData;
    if (!heriData) {
        return { ok: false, reason: "Data HERI tidak tersedia" };
    }

    const dataRows = rawRows.slice(layout.headerRowIndex + 1);
    const outputRows = [];
    const fieldIssues = [];
    const warnings = [];
    const usedPlantEquipmentKeys = options.usedPlantEquipmentKeys ?? new Set();
    let seqNo = 0;
    let emptyPointCount = 0;
    let emptyCounterCount = 0;

    // Aturan 1: filter Protect Y umum (level, merah, STAS13, seen, drive unit).
    /** @type {Array<{ rowIndex: number, funcLoc: string, funcLocKey: string, desc: string, plant: string }>} */
    const filter1Rows = [];

    for (let i = 0; i < dataRows.length; i += 1) {
        if (options.redEquipGroupRowIndices?.has(i)) continue;

        const r = dataRows[i] || [];
        const funcLoc = cellText(r[idxFuncLoc]);
        if (!funcLoc) continue;

        const desc = cellText(r[idxDesc]);
        if (!isMeasuringDocumentSourceRow(funcLoc, desc)) continue;

        const funcLocKey = normalizeFuncLocKey(funcLoc);
        if (funcLocKey && options.seenFuncLocKeys?.has(funcLocKey)) continue;

        const effectiveCostCenter =
            options.effectiveCostCenterByRowIndex?.get(i);
        if (shouldSkipLevel5Stas13Row(funcLocKey, effectiveCostCenter)) continue;

        const plant = getFuncLocSecondSegment(funcLoc);
        if (!plant) continue;

        filter1Rows.push({ rowIndex: i, funcLoc, funcLocKey, desc, plant });
    }

    // Aturan 2: desc exact match nama kolom HERI (+ sterilizer alias).
    /** @type {Array<{ rowIndex: number, funcLoc: string, funcLocKey: string, desc: string, plant: string, equipmentName: string }>} */
    const heriMatches = [];
    for (const entry of filter1Rows) {
        const equipmentName = matchHeriEquipmentName(
            entry.desc,
            heriData.equipmentNamesLongestFirst,
            { allowSterilizerAlias: true }
        );
        if (!equipmentName) continue;
        heriMatches.push({ ...entry, equipmentName });
    }

    const rootAnchors = heriMatches.filter(
        anchor =>
            !heriMatches.some(
                other =>
                    other.rowIndex !== anchor.rowIndex &&
                    isFuncLocDescendant(anchor.funcLoc, other.funcLoc)
            )
    );

    /** @type {Array<{ rowIndex: number, funcLoc: string, funcLocKey: string, desc: string, plant: string, equipmentName: string, counterValue: number | null, counterIssue?: string }>} */
    const anchorsToEmit = [];

    for (const anchor of rootAnchors.sort((a, b) => a.rowIndex - b.rowIndex)) {
        const plantEquipKey = isSterilizerHeriAlias(anchor.desc, anchor.equipmentName)
            ? `${anchor.plant.toUpperCase()}\x1f${cellText(anchor.desc).toUpperCase()}`
            : `${anchor.plant.toUpperCase()}\x1f${anchor.equipmentName}`;
        if (usedPlantEquipmentKeys.has(plantEquipKey)) {
            warnings.push(
                `${fileName} baris ${layout.headerRowIndex + 2 + anchor.rowIndex}: duplikat plant+alat ${anchor.plant}/${anchor.equipmentName} — dilewati`
            );
            continue;
        }

        const counterResult = resolveHeriCounterForEquipment(
            heriData,
            anchor.plant,
            anchor.equipmentName
        );

        usedPlantEquipmentKeys.add(plantEquipKey);

        if (counterResult.skipAnchor) {
            warnings.push(
                `${fileName} baris ${layout.headerRowIndex + 2 + anchor.rowIndex}: ${anchor.plant}/${anchor.equipmentName} — ${counterResult.skipReason}, dilewati`
            );
            continue;
        }

        anchorsToEmit.push({
            ...anchor,
            counterValue: counterResult.counterValue,
            counterIssue: counterResult.counterIssue
        });
    }

    // Aturan 3+4: emit blok atomik — anchor + turunan func loc; anchor tidak emit → seluruh blok skip.
    for (const anchor of anchorsToEmit) {
        if (anchor.funcLocKey && options.seenFuncLocKeys?.has(anchor.funcLocKey)) {
            continue;
        }

        const descendants = filter1Rows
            .filter(
                entry =>
                    entry.rowIndex !== anchor.rowIndex &&
                    isFuncLocDescendant(entry.funcLoc, anchor.funcLoc)
            )
            .sort((a, b) => a.rowIndex - b.rowIndex);

        const block = [anchor, ...descendants];

        for (let bi = 0; bi < block.length; bi += 1) {
            const entry = block[bi];
            if (
                bi > 0 &&
                entry.funcLocKey &&
                options.seenFuncLocKeys?.has(entry.funcLocKey)
            ) {
                continue;
            }

            seqNo += 1;
            const built = buildOutputRow(columnOrder, {
                seqNo,
                plant: entry.plant,
                desc: entry.desc,
                counterValue: anchor.counterValue,
                equipmentByPlantDesc: options.equipmentByPlantDesc,
                pointByEquipment: options.pointByEquipment
            });

            const sourceExcelRow = layout.headerRowIndex + 2 + entry.rowIndex;
            fieldIssues.push(
                ...validateMeasuringDocumentFields(
                    {
                        counterIssue: anchor.counterIssue,
                        equnrFromMapping: built.equnrFromMapping,
                        equnr: built.equnr,
                        point: built.point,
                        counterValue: anchor.counterValue
                    },
                    {
                        regional: options.regional ?? "",
                        fileName,
                        sourceExcelRow,
                        equipmentDescription: entry.desc,
                        plant: entry.plant,
                        funcLoc: entry.funcLoc
                    }
                )
            );

            if (anchor.counterValue === null) emptyCounterCount += 1;
            if (!built.point) emptyPointCount += 1;
            if (entry.funcLocKey) options.seenFuncLocKeys?.add(entry.funcLocKey);
            outputRows.push(built.row);
        }
    }

    return {
        ok: true,
        rows: outputRows,
        warnings,
        fieldIssues,
        emptyPointCount,
        emptyCounterCount
    };
}

/**
 * @param {Array<Record<string, unknown>>} rows
 * @param {string[]} columnOrder
 * @returns {Array<Record<string, unknown>>}
 */
function renumberMeasuringDocumentSeqNo(rows, columnOrder) {
    const colNoUrut = findColumnByPrefix(columnOrder, "No. Urut");
    if (!colNoUrut || !rows.length) return rows;
    return rows.map((row, index) => ({
        ...row,
        [colNoUrut]: index + 1
    }));
}

module.exports = {
    buildOutputRow,
    buildMeasuringDocumentRows,
    isFuncLocDescendant,
    resolveHeriCounterForAnchor: resolveHeriCounterForEquipment,
    renumberMeasuringDocumentSeqNo,
    descForEquipmentOutput,
    MEASUREMENT_DATE,
    MEASUREMENT_TIME
};
