/**
 * Agregat output LSMW: {REGIONAL N}.xlsx + all-pks.xlsx di root run (luar folder REGIONAL).
 */

const path = require("path");

const ALL_PKS_FILE_NAME = "all-pks.xlsx";

function regionalAggregateFileName(regionalName) {
    const base = String(regionalName ?? "").trim();
    return `${base.replace(/[/\\?%*:|"<>]/g, "_")}.xlsx`;
}

/**
 * @param {{
 *   runDir: string,
 *   regionalRowsByName: Map<string, Array<unknown>>,
 *   allRows: Array<unknown>,
 *   writeRows: (outPath: string, rows: Array<unknown>) => Promise<void>,
 *   label?: string
 * }} params
 */
async function writeRegionalAndAllPksAggregates(params) {
    const { runDir, regionalRowsByName, allRows, writeRows, label = "baris" } =
        params;

    for (const [regional, rows] of regionalRowsByName) {
        if (!rows || rows.length === 0) continue;
        const outPath = path.join(runDir, regionalAggregateFileName(regional));
        await writeRows(outPath, rows);
        console.log(`Agregat regional: ${outPath} (${rows.length} ${label})`);
    }

    if (allRows.length > 0) {
        const summaryPath = path.join(runDir, ALL_PKS_FILE_NAME);
        await writeRows(summaryPath, allRows);
        console.log(`Agregat all-pks: ${summaryPath} (${allRows.length} ${label})`);
    }
}

module.exports = {
    ALL_PKS_FILE_NAME,
    regionalAggregateFileName,
    writeRegionalAndAllPksAggregates
};
