/**
 * lsmw_delete_func_loc.cjs
 * ─────────────────────────────────────────────────────────────────────────────
 * Pipeline delete functional location PKS → template_delete_func_loc.xlsx.
 *
 * Sumber: prompt URL folder Protect Y (REGIONAL*), atau --protect-y-folder-link <url>.
 * Template: template_delete_func_loc.xlsx — kolom TPLNR + OLD_NAME, isi mulai baris 4.
 *
 * Mapping
 *   TPLNR ← FUNCTIONAL LOCATION BEFORE
 *   OLD_NAME ← FunctLocDescrip. dari TRUTH_OF_DATA_30_JANUARI.xlsx (match Functional Location = TPLNR)
 *
 * Filter
 *   - Setiap baris dengan FUNCTIONAL LOCATION BEFORE terisi ikut (tanpa syarat kolom lain).
 *   - Nilai unik FUNCTIONAL LOCATION BEFORE per file dan lintas file (duplikat diabaikan).
 *   - Skip FUNCTIONAL LOCATION BEFORE di equipment-func-loc-belum-close-order.xlsx (order belum close).
 *
 * Agregat
 *   - {REGIONAL N}.xlsx di root folder output (gabungan per regional)
 *
 *   node lsmw_delete_func_loc.cjs
 *   node lsmw_delete_func_loc.cjs --protect-y-folder-link <url>
 *   node lsmw_delete_func_loc.cjs --output-drive-folder-link <url>
 *
 * Output: Output/3. LSMW DELETE MASS FUNCTIONAL LOCATION SISI/***.xlsx
 */

const path = require("path");
const { runLsmwDeleteJob } = require("./utils/lsmw_delete_runner.cjs");

function normalizeFuncLocKey(code) {
    if (code === null || code === undefined || code === "") return "";
    return String(code).trim();
}

const LOCAL_OUTPUT_ROOT = path.join("Output", "3. LSMW DELETE MASS FUNCTIONAL LOCATION SISI");

runLsmwDeleteJob({
    templatePath: "./template_delete_func_loc.xlsx",
    localOutputRoot: LOCAL_OUTPUT_ROOT,
    columnJsonPath: path.join(LOCAL_OUTPUT_ROOT, "column.json"),
    sourceColumnHeader: "FUNCTIONAL LOCATION BEFORE",
    enableRegionalAggregate: true,
    enableGlobalAggregate: false,
    normalizeValue: normalizeFuncLocKey,
    uniqueLabel: "functional location (TPLNR)",
    oldNameLookup: "funcloc-desc-by-funcloc"
}).catch(err => {
    console.error(err);
    process.exit(1);
});
