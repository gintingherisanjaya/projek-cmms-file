/**
 * lsmw_delete_equipment.cjs
 * ─────────────────────────────────────────────────────────────────────────────
 * Pipeline delete equipment PKS → template_delete_equipment.xlsx.
 *
 * Sumber: prompt URL folder Protect Y (REGIONAL*), atau --protect-y-folder-link <url>.
 * Template: template_delete_equipment.xlsx — kolom EQUNR + OLD_NAME, isi mulai baris 4.
 *
 * Mapping
 *   EQUNR ← EQUIPMENT NUMBER
 *   OLD_NAME ← EQKTU dari TRUTH_OF_DATA_30_JANUARI.xlsx (match Equipment = EQUNR)
 *
 * Filter
 *   - EQUIPMENT NUMBER wajib terisi.
 *   - Nilai unik per file dan lintas file (duplikat diabaikan).
 *   - Skip EQUIPMENT NUMBER di equipment-func-loc-belum-close-order.xlsx (order belum close).
 *
 * Agregat
 *   - {REGIONAL N}.xlsx di root folder output (gabungan per regional)
 *   - all-pks.xlsx di root folder output (gabungan semua regional)
 *
 *   node lsmw_delete_equipment.cjs
 *   node lsmw_delete_equipment.cjs --protect-y-folder-link <url>
 *   node lsmw_delete_equipment.cjs --output-drive-folder-link <url>
 *
 * Output: Output/4. LSMW DELETE EQUIPMENT SISI/***.xlsx (nama sama dengan file sumber)
 */

const path = require("path");
const { runLsmwDeleteJob } = require("./utils/lsmw_delete_runner.cjs");

function normalizeEquipmentNumber(v) {
    if (v === null || v === undefined || v === "") return "";
    return String(v).trim();
}

const LOCAL_OUTPUT_ROOT = path.join("Output", "4. LSMW DELETE EQUIPMENT SISI");

runLsmwDeleteJob({
    templatePath: "./template_delete_equipment.xlsx",
    localOutputRoot: LOCAL_OUTPUT_ROOT,
    columnJsonPath: path.join(LOCAL_OUTPUT_ROOT, "column.json"),
    sourceColumnHeader: "EQUIPMENT NUMBER",
    normalizeValue: normalizeEquipmentNumber,
    uniqueLabel: "equipment (EQUNR)",
    oldNameLookup: "eqktu-by-equipment",
    enableRegionalAggregate: true,
    enableGlobalAggregate: true,
    globalAggregateFileName: "all-pks.xlsx"
}).catch(err => {
    console.error(err);
    process.exit(1);
});
