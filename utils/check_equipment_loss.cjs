/**
 * Logika perbandingan LPP vs Protect Y untuk check-equipment-loss.
 */
const { normalizeCostCenter } = require("./equipment_gathering_columns.cjs");

const STATUS = {
    EXIST: "EXIST",
    MISSING: "MISSING",
    ANOMALY: "ANOMALY",
    UNUSED: "UNUSED"
};

function isUnusedCostCenter(costCenterBefore) {
    return !normalizeCostCenter(costCenterBefore).includes("STAS");
}

/**
 * @param {string} equipNum
 * @param {{ eqktuBefore?: string, costCenterBefore?: string } | undefined} lppRow
 * @param {{ eqktuBefore?: string, costCenterBefore?: string } | undefined} pyRow
 * @param {string} namaPks
 */
function buildEquipmentItem(equipNum, lppRow, pyRow, namaPks) {
    const costCenterBefore =
        lppRow?.costCenterBefore ?? pyRow?.costCenterBefore ?? "";

    if (isUnusedCostCenter(costCenterBefore)) {
        return {
            equipmentNumber: equipNum,
            eqktuBefore: lppRow?.eqktuBefore || pyRow?.eqktuBefore || "",
            costCenterBefore,
            namaPks,
            status: STATUS.UNUSED
        };
    }

    if (lppRow && pyRow) {
        return {
            equipmentNumber: equipNum,
            eqktuBefore: lppRow.eqktuBefore || pyRow.eqktuBefore,
            costCenterBefore,
            namaPks,
            status: STATUS.EXIST
        };
    }

    if (lppRow) {
        return {
            equipmentNumber: equipNum,
            eqktuBefore: lppRow.eqktuBefore,
            costCenterBefore,
            namaPks,
            status: STATUS.MISSING
        };
    }

    return {
        equipmentNumber: equipNum,
        eqktuBefore: pyRow?.eqktuBefore ?? "",
        costCenterBefore,
        namaPks,
        status: STATUS.ANOMALY
    };
}

/**
 * @param {Map<string, { eqktuBefore: string, costCenterBefore: string }>} lppMap
 * @param {Map<string, { eqktuBefore: string, costCenterBefore: string }>} protectYMap
 * @param {string} namaPks
 */
function compareEquipmentMaps(lppMap, protectYMap, namaPks) {
    const items = [];

    for (const [equipNum, lppRow] of lppMap) {
        items.push(
            buildEquipmentItem(equipNum, lppRow, protectYMap.get(equipNum), namaPks)
        );
    }

    for (const [equipNum, pyRow] of protectYMap) {
        if (lppMap.has(equipNum)) continue;
        items.push(buildEquipmentItem(equipNum, undefined, pyRow, namaPks));
    }

    return items;
}

module.exports = {
    STATUS,
    isUnusedCostCenter,
    buildEquipmentItem,
    compareEquipmentMaps
};
