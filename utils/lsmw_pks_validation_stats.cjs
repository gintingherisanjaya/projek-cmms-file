const { normalizeFuncLocForCompare } = require("./equipment_number_mapping.cjs");

function createValidationStats() {
    return {
        equnrFoundCount: 0,
        equnrNotFoundCount: 0,
        plant: "",
        validFuncLocCount: 0,
        invalidFuncLocCount: 0,
        invalidFuncLocSamples: []
    };
}

/**
 * @param {ReturnType<typeof createValidationStats>} stats
 * @param {{
 *   funcLoc: string,
 *   plantSegment: string,
 *   equnrFromMapping: boolean,
 *   mappingFuncLoc: string
 * }} item
 */
function recordValidationItem(stats, item) {
    const { funcLoc, plantSegment, equnrFromMapping, mappingFuncLoc } = item;

    if (!stats.plant && plantSegment) {
        stats.plant = plantSegment;
    }

    if (equnrFromMapping) {
        stats.equnrFoundCount += 1;
        if (
            normalizeFuncLocForCompare(funcLoc) ===
            normalizeFuncLocForCompare(mappingFuncLoc)
        ) {
            stats.validFuncLocCount += 1;
        } else {
            stats.invalidFuncLocCount += 1;
            if (stats.invalidFuncLocSamples.length < 2) {
                stats.invalidFuncLocSamples.push(String(funcLoc).trim());
            }
        }
    } else {
        stats.equnrNotFoundCount += 1;
    }
}

function formatInvalidFuncLocList(samples) {
    return samples.join("; ");
}

function validationStatsToResult(stats) {
    return {
        equnrFoundCount: stats.equnrFoundCount,
        equnrNotFoundCount: stats.equnrNotFoundCount,
        plant: stats.plant,
        validFuncLocCount: stats.validFuncLocCount,
        invalidFuncLocCount: stats.invalidFuncLocCount,
        invalidFuncLocList: formatInvalidFuncLocList(stats.invalidFuncLocSamples)
    };
}

module.exports = {
    createValidationStats,
    recordValidationItem,
    formatInvalidFuncLocList,
    validationStatsToResult
};
