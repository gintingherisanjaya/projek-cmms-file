/**
 * Transform sumber PKS → baris nilai template equipment (untuk template-lsmw-cheking).
 * Logika mengacu pada lsmw_equipment_v0.cjs — tidak mengubah file tersebut.
 */

const path = require("path");
const fs = require("fs");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const XLSX = require("xlsx");
const {
    isEmpty,
    ensureCwcPlannerLookups,
    getEffectiveCostCenter,
    logRegionalMappingForFile,
    buildColumnIndexFirstWins,
    resolvePlannerGroup,
    resolveWorkCenter
} = require("./lsmw_lookups.cjs");
const { ensureAbcRules, resolveAbcIndicator } = require("./lsmw_abc.cjs");
const {
    collectRedEquipGroupRowIndices
} = require("./lsmw_equip_group_red.cjs");
const { findDataLayout } = require("./lsmw_cell_fill.cjs");
const {
    ensureFunclocDescAliasMap,
    applyFunclocDescAlias
} = require("./funcloc_desc_alias.cjs");
const { normalizeCapacityValue } = require("./lsmw_capacity.cjs");
const { resolveBaujjYear } = require("./lsmw_baujj.cjs");
const { findCostCenterColumnIndex } = require("./equipment_gathering_columns.cjs");

const SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly"
];

const OAUTH_PATH = path.join(__dirname, "..", "oauth.json");
const TOKEN_PATH = path.join(__dirname, "..", "token.json");
const EQUIPMENT_GROUP_MAPPING_PATH = path.join(
    __dirname,
    "..",
    "equipment-group-mapping.json"
);

let drive = null;
let sheets = null;
let equipmentGroupMappingCache = null;
let equipmentGroupFromDescMod = null;
const usedFuncLocAfter = new Set();

function resetTransformState() {
    usedFuncLocAfter.clear();
}

/** Sama seperti initDrive di lsmw_equipment_v0.cjs (Drive + Sheets). */
async function initDrive() {
    if (drive && sheets) return;

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
        const token = newAuth.credentials;
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
        auth.setCredentials(token);
        console.log("Token saved to token.json");
    }

    drive = google.drive({ version: "v3", auth });
    sheets = google.sheets({ version: "v4", auth });
}

function getDriveClient() {
    if (!drive) throw new Error("initDrive() belum dipanggil");
    return drive;
}

async function ensureEquipmentGroupMapping() {
    if (equipmentGroupMappingCache && equipmentGroupFromDescMod) {
        return {
            mod: equipmentGroupFromDescMod,
            mapping: equipmentGroupMappingCache
        };
    }
    equipmentGroupFromDescMod = await import("../equipmentGroupFromDesc.js");
    equipmentGroupMappingCache =
        equipmentGroupFromDescMod.loadEquipmentGroupMapping(
            EQUIPMENT_GROUP_MAPPING_PATH
        );
    return {
        mod: equipmentGroupFromDescMod,
        mapping: equipmentGroupMappingCache
    };
}

function normalizeHeader(name) {
    if (!name) return "";
    return name.toString().trim().toUpperCase();
}

function isValidPlantCode(code) {
    const s = String(code ?? "").trim().toUpperCase();
    return /^[A-Z0-9]{4}$/.test(s);
}

function resolveFileBegru(
    dataRows,
    equipmentRows,
    val,
    getValueWithParent,
    idxFuncLoc
) {
    for (const r of dataRows) {
        const mp = val(r, "MAINTENANCE PLAN AFTER");
        if (!isEmpty(mp)) return mp;
    }
    for (const r of equipmentRows) {
        const mp = getValueWithParent(r, "MAINTENANCE PLAN AFTER");
        if (!isEmpty(mp)) return mp;
    }
    const scanPlant = rows => {
        for (const r of rows) {
            const p = val(r, "MAINTENANCE PLANT");
            if (!isEmpty(p) && isValidPlantCode(p)) {
                return String(p).trim().toUpperCase();
            }
        }
        for (const r of rows) {
            const p = val(r, "POM");
            if (!isEmpty(p) && isValidPlantCode(p)) {
                return String(p).trim().toUpperCase();
            }
        }
        if (idxFuncLoc !== undefined) {
            for (const r of rows) {
                const raw = r[idxFuncLoc];
                if (!raw) continue;
                const segs = String(raw).trim().split("-").filter(Boolean);
                if (segs.length >= 2) {
                    const plant = segs[1].trim().toUpperCase();
                    if (isValidPlantCode(plant)) return plant;
                }
            }
        }
        return null;
    };
    return scanPlant(dataRows) || scanPlant(equipmentRows);
}

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
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

/**
 * @returns {Promise<{ sourceName: string, outputRows: Array<{ values: unknown[], functionalLocation: string, sourceExcelRow: number|null }>, redSkipped: number } | null>}
 */
async function buildEquipmentOutputRows(sourcePath, sourceName, driveFile) {
    await ensureFunclocDescAliasMap();
    await ensureCwcPlannerLookups();
    const abcRules = ensureAbcRules();
    await initDrive();

    const workbook = XLSX.readFile(sourcePath, { cellStyles: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const layout = findDataLayout(rows);
    const headerRow = rows[layout.headerRowIndex];
    if (!headerRow) return null;

    const dataRows = rows.slice(layout.headerRowIndex + 1);
    const colIndex = buildColumnIndexFirstWins(headerRow, normalizeHeader);
    const idxFuncLoc = colIndex["FUNCTIONAL LOCATION AFTER"];
    const idxFunclocDesc = findFunclocDescAfterColumnIndex(headerRow);

    if (idxFuncLoc === undefined) return null;

    function val(r, name) {
        const idx = colIndex[normalizeHeader(name)];
        return idx !== undefined ? r[idx] ?? null : null;
    }

    const idxCostCenter = findCostCenterColumnIndex(headerRow);
    const idxEquipGroupAfter =
        colIndex["EQUIPMENT GROUP AFTER"] ?? layout.idxEquipGroupAfter;

    const redEquipGroupRowIndices = await collectRedEquipGroupRowIndices({
        sourcePath,
        colIndex0: idxEquipGroupAfter,
        dataRowCount: dataRows.length,
        dataStartExcelRow: layout.dataStartExcelRow,
        driveFile: driveFile ?? null,
        sheetsApi: sheets
    });
    const redSkipCount = redEquipGroupRowIndices.size;

    const funcLocMap = {};
    for (const r of dataRows) {
        const code = normalizeFuncLocKey(r[idxFuncLoc]);
        if (code) funcLocMap[code] = r;
    }

    function getValueWithParent(r, columnName) {
        const idx = colIndex[normalizeHeader(columnName)];
        if (idx === undefined) return null;
        let value = r[idx];
        if (value !== null && value !== "" && value !== undefined) return value;
        let funcLoc = r[idxFuncLoc];
        while (funcLoc && funcLoc.includes("-")) {
            funcLoc = funcLoc.substring(0, funcLoc.lastIndexOf("-"));
            const parentRow = funcLocMap[funcLoc];
            if (!parentRow) continue;
            const parentVal = parentRow[idx];
            if (
                parentVal !== null &&
                parentVal !== "" &&
                parentVal !== undefined
            ) {
                return parentVal;
            }
        }
        return null;
    }

    function funcLocLevel(code) {
        if (!code || typeof code !== "string") return 0;
        return code.trim().split("-").filter(s => s.length > 0).length;
    }

    function funclocDescRaw(r) {
        if (idxFunclocDesc === undefined) return "";
        const v = r[idxFunclocDesc];
        return v !== null && v !== undefined && v !== "" ? String(v) : "";
    }

    function funclocDescText(r) {
        return applyFunclocDescAlias(funclocDescRaw(r));
    }

    function isEquipmentRow(r) {
        const code = r[idxFuncLoc];
        if (funcLocLevel(code) < 5) return false;
        const desc = String(funclocDescRaw(r) ?? "");
        if (/drive\s*unit/i.test(desc)) return false;
        if (/accessories\s+weigh\s*bridge/i.test(desc)) return false;
        return true;
    }

    const equipmentRows = [];
    const rowIndexByRow = new Map();
    for (let i = 0; i < dataRows.length; i++) {
        const r = dataRows[i];
        if (redEquipGroupRowIndices.has(i)) continue;
        if (!isEquipmentRow(r)) continue;
        equipmentRows.push(r);
        rowIndexByRow.set(r, i);
    }

    const { mod: egMod, mapping: egMapping } =
        await ensureEquipmentGroupMapping();
    const egByRowIndex = egMod.resolveEquipmentGroupAssignmentsForRows(
        equipmentRows.map(r => ({
            rowIndex: rowIndexByRow.get(r),
            funcLocNorm: normalizeFuncLocKey(r[idxFuncLoc]),
            desc: funclocDescRaw(r)
        })),
        egMapping
    );

    function resolveEqart(r) {
        const fromSource = val(r, "EQUIPMENT GROUP AFTER");
        if (!isEmpty(fromSource)) return fromSource;
        const rowIndex = rowIndexByRow.get(r);
        const fromMapping =
            rowIndex !== undefined
                ? egByRowIndex.get(rowIndex)?.value ?? ""
                : "";
        if (!isEmpty(fromMapping)) return fromMapping;
        return egMod.normalizeEquipmentGroupValue(funclocDescRaw(r));
    }

    const fileBegru = resolveFileBegru(
        dataRows,
        equipmentRows,
        val,
        getValueWithParent,
        idxFuncLoc
    );

    if (equipmentRows.length > 0 && isEmpty(fileBegru)) return null;

    await logRegionalMappingForFile(fileBegru, sourceName);

    const outputRows = [];
    let lastCapacity = null;
    let lastYear = null;

    for (const r of equipmentRows) {
        const funcLocKey = normalizeFuncLocKey(r[idxFuncLoc]);
        if (funcLocKey && usedFuncLocAfter.has(funcLocKey)) continue;

        const capacity = normalizeCapacityValue(
            getValueWithParent(r, "CAPACITY")
        );
        const year = resolveBaujjYear(r, getValueWithParent);
        const capOut = capacity === lastCapacity ? null : capacity;
        const yearOut = year === lastYear ? null : year;
        lastCapacity = capacity;
        lastYear = year;

        if (funcLocKey) usedFuncLocAfter.add(funcLocKey);

        const effectiveCostCenter = getEffectiveCostCenter(
            r,
            funcLocKey,
            funcLocMap,
            idxCostCenter
        );

        const sourceDataIndex = rowIndexByRow.get(r);
        outputRows.push({
            values: [
                null,
                1,
                fileBegru,
                capOut,
                resolveEqart(r),
                val(r, "MERK"),
                yearOut,
                null,
                funclocDescText(r),
                null,
                fileBegru,
                "TEK",
                resolveAbcIndicator(r, funcLocKey, val, abcRules),
                "PALM",
                idxCostCenter !== undefined ? r[idxCostCenter] ?? null : null,
                fileBegru,
                resolvePlannerGroup(
                    r,
                    funcLocKey,
                    effectiveCostCenter,
                    val,
                    fileBegru
                ),
                resolveWorkCenter(
                    r,
                    funcLocKey,
                    effectiveCostCenter,
                    val,
                    fileBegru
                ),
                fileBegru,
                "ZPM_PTPN",
                val(r, "FUNCTIONAL LOCATION AFTER"),
                null,
                null
            ],
            functionalLocation: funcLocKey,
            sourceExcelRow:
                sourceDataIndex !== undefined
                    ? layout.dataStartExcelRow + sourceDataIndex
                    : null
        });
    }

    return { sourceName, outputRows, redSkipped: redSkipCount };
}

module.exports = {
    initDrive,
    getDriveClient,
    resetTransformState,
    buildEquipmentOutputRows
};
