/**
 * help.cjs — ringkasan perintah pnpm run di proyek ini.
 *
 *   pnpm run help
 */

const fs = require("fs");
const path = require("path");

const PACKAGE_JSON = path.join(__dirname, "package.json");

/** @type {Record<string, string>} */
const COMMAND_HELP = {
    help: [
        "Menampilkan daftar perintah pnpm run dan penjelasan singkatnya.",
        "Tidak memerlukan oauth.json."
    ].join("\n    "),

    test: [
        "Placeholder npm test (belum diimplementasi)."
    ].join("\n    "),

    "lsmw-equipment": [
        "Konversi data equipment PKS (Protect Y) → template LSMW upload equipment.",
        "Sumber: prompt URL folder Protect Y (subfolder REGIONAL*), atau --protect-y-folder-link <url>.",
        "Template: template_equipment.xlsx — header template baris 3, data output mulai baris 7.",
        "Filter: funcloc level ≥ 5; skip drive unit / accessories weigh bridge;",
        "  skip baris EQUIPMENT GROUP AFTER ber-background merah; dedupe FUNCTIONAL LOCATION AFTER lintas run.",
        "Skip tambahan: level 5 + cost center efektif STAS13 (Effluent/Limbah).",
        "Kolom FUNCTLOC DESC: header dinamis (FUNCTLOC + DESC + AFTER; prioritas yang mengandung LEVEL).",
        "Normalisasi typo sumber: ELECTROMOTR → ELECTROMOTOR pada FUNCTLOC DESC. AFTER.",
        "Cost center tier (lookup): COST CENTER AFTER → COST CENTER tanpa BEFORE → COST CENTER BEFORE;",
        "  cost center efektif (naik parent funcloc) dipakai lookup CWC/PG; kolom KOSTL output = nilai mentah kolom terpilih.",
        "Berhenti segera jika ada baris tanpa mapping EQART; laporkan FUNCTLOC DESC. AFTER + funcloc + nama file;",
        "  tulis Output/invalid-group-mapping/{timestamp}.json.",
        "Output per PKS: Output/1. LSMW Create Equipment V.O/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx + all-pks.xlsx di root folder output (luar subfolder REGIONAL).",
        "Audit kolom: column.json (header baris per file + missing_column) + validation.xlsx.",
        "  missing_column = daftar kolom kanonik Protect Y yang tidak ter-resolve (36 kolom, termasuk EQUIPMENT GROUP AFTER).",
        "  validation.xlsx: satu baris per PKS; kolom nama file + 36 kolom kanonik;",
        "    isi sel = teks header sumber/alternatif yang dipilih sistem (kosong jika tidak ada).",
        "  Jika EQUIPMENT GROUP AFTER missing → masuk missing_column; skip merah tidak aktif.",
        "Alternatif: header exact EQUIPMENT GROUP setelah COST CENTER AFTER dianggap EQUIPMENT GROUP AFTER (audit + skip merah).",
        "Folder to-delete/: mirror REGIONAL N/{PKS}.xlsx + agregat;",
        "  template template_delete_equipment.xlsx (EQUNR + OLD_NAME, data baris 4);",
        "  hanya equipment di-skip merah via kolom alternatif (bukan EQUIPMENT GROUP AFTER asli);",
        "  EQUNR dari new-equipment-number.xlsx (lookup sama lsmw-maintenance-item: plant + FUNCTLOC DESC. AFTER);",
        "  bukan EQUIPMENT NUMBER Protect Y; fallback {segmen2}-{desc} + [warn] jika tidak ada di mapping;",
        "  OLD_NAME = FUNCTLOC DESC. AFTER setelah typo-fix + alias (sama SHTXT create equipment);",
        "  dedupe EQUNR per file dan lintas file.",
        "Retry unduh/list Drive saat rate limit/timeout: hingga 6× dengan jeda 5 detik.",
        "Opsional: --output-drive-folder-link <url>.",
        "Pemetaan kolom output:",
        "  EQTYP* = 1 (konstanta)",
        "  BEGRU* / SWERK* / IWERK* / WERGW = per file: MAINTENANCE PLAN AFTER (pertama) → MAINTENANCE PLANT → POM → segmen-2 FUNCTIONAL LOCATION AFTER",
        "  GROES = CAPACITY (baris/parent); teks dalam (... ) dihapus; hanya jika beda dari baris sebelumnya",
        "  EQART = equipment-group-mapping.json dari FUNCTLOC DESC. AFTER (bukan salin kolom EQUIPMENT GROUP AFTER)",
        "  HERST = MERK",
        "  BAUJJ = STANDART UMUR TEKNIS (TAHUN) jika 4 digit (baris/parent), else CONSTRUCTION YEAR (baris/parent); kosong boleh; hanya jika beda dari baris sebelumnya",
        "  SHTXT* = FUNCTLOC DESC. AFTER (+ alias mapping-alias-desc-func-loc.xlsx bila cocok)",
        "  BEBER = TEK (konstanta)",
        "  ABCKZ = utils/abc.csv (pola funcloc) → default 2",
        "  BUKRS = PALM (konstanta)",
        "  KOSTL* = sel cost center Protect Y (kolom tier di atas; tanpa inherit parent di output)",
        "  INGRP = utils/planner-group.csv + cost center efektif + maintenance plant; fallback PLANNER GROUP AFTER",
        "  GEWRK = utils/cwc.csv + cost center efektif + maintenance plant; fallback WORK CENTER AFTER",
        "  RBNR = ZPM_PTPN (konstanta)",
        "  TPLNR = FUNCTIONAL LOCATION AFTER"
    ].join("\n    "),

    "lsmw-func-loc": [
        "Konversi data functional location PKS → template LSMW create func loc.",
        "Sumber: prompt URL folder Protect Y (subfolder REGIONAL*), atau --protect-y-folder-link <url>.",
        "Prompt: pilih subfolder REGIONAL yang diproses (checkbox, default semua terpilih).",
        "Template: template_create_func_loc_lsmw1.xlsx — data output mulai baris 5.",
        "Filter: semua baris dengan FUNCTIONAL LOCATION AFTER terisi; dedupe funcloc lintas file.",
        "Kolom cost center tier (lookup GEWRK/INGRP): sama lsmw-equipment (AFTER → tanpa BEFORE → BEFORE),",
        "  dengan inherit ke parent funcloc bila sel kosong.",
        "Normalisasi typo sumber: ELECTROMOTR → ELECTROMOTOR pada FUNCTLOC DESC. AFTER.",
        "Baris sintetis (jika ada desc berisi STATION): parent Area Pabrik (PLTXT = nama PKS dari nama file),",
        "  lalu AREA PABRIK; GEWRK/INGRP kosong pada baris sintetis.",
        "Output per PKS: Output/2. LSMW Create Functional Location V1/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx di root folder output (luar subfolder REGIONAL);",
        "  dibangun dari file Excel PKS lokal (body baris 5+, kolom A STRNO), bukan dari memori.",
        "Retry unduh/list Drive saat timeout/jaringan: hingga 4x dengan jeda 1 detik.",
        "Opsional: --with-cost-center (KOSTL baris non-stasiun saja), --output-drive-folder-link <url>.",
        "Pemetaan kolom output:",
        "  STRNO = FUNCTIONAL LOCATION AFTER",
        "  ALKEY = 1",
        "  TPLKZ = PTPN",
        "  FLTYP = M",
        "  PLTXT = FUNCTLOC DESC. AFTER (+ alias mapping-alias-desc-func-loc.xlsx bila cocok)",
        "  SWERK = maintenance plant per file (prioritas sama lsmw-equipment)",
        "  BUKRS = PALM",
        "  KOSTL = stasiun (4 segmen): segmen 2 STRNO + STAS + 2 digit segmen 4 (mis. 1F11STAS01);",
        "    baris lain: kosong, kecuali --with-cost-center → cost center efektif Protect Y",
        "  IWERK = sama dengan SWERK",
        "  GEWRK = utils/cwc.csv + cost center efektif + maintenance plant",
        "  INGRP = utils/planner-group.csv + cost center efektif + maintenance plant",
        "  WERGW = sama dengan SWERK",
        "  IEQUI = X",
        "  TPLMA = parent funcloc (FUNCTIONAL LOCATION AFTER tanpa segmen terakhir)"
    ].join("\n    "),

    "lsmw-changes-equipment": [
        "Generate file perubahan equipment dari Protect Y + equipment-missing-in-protect-y.xlsx.",
        "Sumber Protect Y: prompt URL folder (subfolder REGIONAL*), atau --protect-y-folder-link <url>.",
        "Sumber missing: equipment-missing-in-protect-y.xlsx (Plant, Equipment, Functional Loc., Target Func Loc, Percentage).",
        "Match missing: Plant = segmen ke-2 FUNCTIONAL LOCATION AFTER (per baris) + Target Func Loc = AFTER.",
        "Prompt: pilih subfolder REGIONAL yang diproses (checkbox, default semua).",
        "Template: template_changes_equipment.xlsx — data mulai baris 4.",
        "Mapping: OLD_NAME ← EQKTU Truth of Data (fallback Description missing) | EQUNR ← EQUIPMENT NUMBER / Equipment missing |",
        "  TPLNR ← FUNCTIONAL LOCATION AFTER | NEW_NAME ← FUNCTLOC DESC. AFTER | Percentage ← missing excel.",
        "Truth of Data: TRUTH_OF_DATA_30_JANUARI.xlsx (kolom Equipment, EQKTU).",
        "Filter: FUNCTIONAL LOCATION AFTER wajib; dedupe EQUIPMENT NUMBER per file dan lintas file.",
        "Skip: equipment atau func loc before yang ada di equipment-func-loc-belum-close-order.xlsx",
        "  (kolom Equipment, Functional Loc.) — masih punya order terbuka, tidak boleh di-change.",
        "Output per PKS: Output/5. LSMW CHANGE EQUIPMENT SISI/REGIONAL N/{nama PKS}.xlsx",
        "validation.xlsx di root output: equipment tidak ada di Truth dan/atau DESC AFTER kosong.",
        "Opsional: --protect-y-folder-link <url>, --output-drive-folder-link <url>."
    ].join("\n    "),

    "lsmw-delete-func-loc": [
        "Generate file hapus massal functional location dari Protect Y + equipment-missing-in-protect-y.xlsx.",
        "Sumber Protect Y: prompt URL folder (subfolder REGIONAL*), atau --protect-y-folder-link <url>.",
        "Sumber missing: equipment-missing-in-protect-y.xlsx — Functional Loc. sebagai BEFORE bila match",
        "  Plant (segmen ke-2 AFTER per baris) + Target Func Loc = FUNCTIONAL LOCATION AFTER.",
        "Prompt: pilih subfolder REGIONAL yang diproses (checkbox, default semua).",
        "Sumber file Protect Y: .xlsx / Google Sheets.",
        "Template: template_delete_func_loc.xlsx — kolom TPLNR + OLD_NAME, data mulai baris 4.",
        "Mapping: TPLNR ← FUNCTIONAL LOCATION BEFORE / Functional Loc. missing |",
        "  OLD_NAME ← FunctLocDescrip. Truth of Data (fallback Description Func loc old missing).",
        "Truth of Data: TRUTH_OF_DATA_30_JANUARI.xlsx (kolom Functional Location, FunctLocDescrip.).",
        "Skip: func loc before yang ada di equipment-func-loc-belum-close-order.xlsx",
        "  (kolom Functional Loc.) — masih punya order terbuka, tidak boleh di-delete.",
        "Dedupe TPLNR per file dan lintas file.",
        "Output per PKS: Output/3. LSMW DELETE MASS FUNCTIONAL LOCATION SISI/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx di root folder output (luar subfolder REGIONAL).",
        "Opsional: --protect-y-folder-link <url>, --output-drive-folder-link <url>."
    ].join("\n    "),

    "lsmw-delete-equipment": [
        "Generate file hapus equipment dari Protect Y + equipment-missing-in-protect-y.xlsx.",
        "Sumber Protect Y: prompt URL folder (subfolder REGIONAL*), atau --protect-y-folder-link <url>.",
        "Sumber missing: equipment-missing-in-protect-y.xlsx — Equipment sebagai EQUNR bila match",
        "  Plant (segmen ke-2 AFTER per baris) + Target Func Loc = FUNCTIONAL LOCATION AFTER.",
        "Prompt: pilih subfolder REGIONAL yang diproses (checkbox, default semua).",
        "Template: template_delete_equipment.xlsx — kolom EQUNR + OLD_NAME, data mulai baris 4.",
        "Mapping: EQUNR ← EQUIPMENT NUMBER / Equipment missing |",
        "  OLD_NAME ← EQKTU Truth of Data (fallback Description missing).",
        "Truth of Data: TRUTH_OF_DATA_30_JANUARI.xlsx (kolom Equipment, EQKTU).",
        "Skip: equipment yang ada di equipment-func-loc-belum-close-order.xlsx",
        "  (kolom Equipment) — masih punya order terbuka, tidak boleh di-delete.",
        "Dedupe EQUNR per file dan lintas file.",
        "Output per PKS: Output/4. LSMW DELETE EQUIPMENT SISI/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx + all-pks.xlsx di root folder output (luar subfolder REGIONAL).",
        "Opsional: --protect-y-folder-link <url>, --output-drive-folder-link <url>."
    ].join("\n    "),

    "lsmw-all": [
        "Menjalankan berurutan: lsmw-equipment → lsmw-func-loc →",
        "lsmw-changes-equipment → lsmw-delete-func-loc → lsmw-delete-equipment."
    ].join("\n    "),

    "generate-excel-link": [
        "Buat Excel berisi link publik semua file di folder Protect Y (REGIONAL*).",
        "Sumber: prompt URL folder Protect Y (subfolder REGIONAL*), atau --protect-y-folder-link <url>.",
        "Output: Output/PKS_LINK_{timestamp}.xlsx (kolom: folder | PKS | link)."
    ].join("\n    "),

    "check-number-of-line-in-same-template": [
        "Bandingkan jumlah baris data nyata semua spreadsheet di folder Drive (template homogen).",
        "Prompt URL folder Drive (subfolder REGIONAL*), atau --folder-link <url>.",
        "Prompt mode jika jumlah baris beda: terminate (exit 1) atau warning saja (exit 0).",
        "Flag non-interaktif: --on-mismatch terminate | warn (alias warning).",
        "Scan selalu selesai semua file; Excel lengkap ditulis meski ada perbedaan.",
        "Skip ~$ dan all-pks.xlsx; urut path; file pertama = baseline.",
        "Output: Output/check-number-of-line-in-same-tenplate/{timestamp WIB}.xlsx",
        "  kolom: nama file | nama sub folder | jumlah baris | baseline | status (baseline/cocok/beda)."
    ].join("\n    "),

    "check-lsmw-maintenance-plan": [
        "Cek counter maintenance plan (SZAEH) vs CDS ZPP jam jalan (CountReadng).",
        "Prompt / CLI: URL folder Drive output lsmw-maintenance-plan (REGIONAL N/),",
        "  --maintenance-plan-link <url> [--cds-path ./cds_zpp_jamjalan_eqv.xlsx] [--ik07-path ./measuring-item-ik07.xlsx].",
        "Baca POINT + SZAEH dari maintenance plan; Equipment dari reverse lookup measuring-item-ik07.xlsx.",
        "CDS: lookup CountReadng by Equipment (kolom Equipment + CountReadng + Date).",
        "  Equipment duplikat di CDS: mismatch jika ada CountReadng yang tidak cocok SZAEH (meski ada yang cocok).",
        "  Kolom tanggal = Date CDS dari baris yang tidak cocok (gabung ; jika banyak).",
        "  Kolom measuring point di output = POINT maintenance plan (informatif, bukan kunci CDS).",
        "Functional loc dari new-equipment-number.xlsx (kolom Equipment + Functional Loc.).",
        "  Parent invalid (mismatch/cds_missing) dari maintenance plan → turunan func loc ikut ditulis;",
        "  equipment + functional loc anak sendiri; SZAEH/CountReadng/tanggal/status warisan parent.",
        "Perbandingan SZAEH vs CountReadng: teks tampilan file (bukan nilai float internal),",
        "  kanonik notasi HERI/koma — tanpa toleransi; CDS 21.45 = SZAEH 21,45.",
        "Status: match | mismatch | cds_missing | no_equipment. Semua baris ditulis.",
        "Output: Output/check-lsmw-maintenance-plan/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx",
        "  + {REGIONAL N}.xlsx + all-pks.xlsx",
        "  kolom: nama regional, nama pks, measuring point, equipment, functional loc, SZAEH, CountReadng, tanggal, status."
    ].join("\n    "),

    "check-lsmw-measuring-document": [
        "Cek measuring document (Drive utama + missing) vs CDS cds_zpp_jamjalan_eqv.xlsx.",
        "Prompt / CLI: URL folder Drive output measuring document (REGIONAL N/),",
        "  URL folder Missing Measuring Document (REGIONAL N/),",
        "  --measuring-doc-link <url> --missing-link <url> [--cds-path ./cds_zpp_jamjalan_eqv.xlsx].",
        "Baca kolom Measuring point + Equipment Number dari tiap PKS; bandingkan ke CDS (Measuring point + Equipment).",
        "Status: finded (ada di CDS) | missing (belum ada di CDS). Semua baris ditulis.",
        "Gabungkan kedua sumber Drive; dedupe pasangan per PKS (first wins).",
        "Output: Output/check-lsmw-measuring-document/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx",
        "  + {REGIONAL N}.xlsx + all-pks.xlsx",
        "  kolom: nama regional, nama pks, measuring point, equipment, status."
    ].join("\n    "),

    "check-lsmw-update-change-equipment": [
        "Audit output LSMW change equipment (template_changes_equipment.xlsx) di folder Drive.",
        "Prompt URL folder Drive output change equipment (subfolder REGIONAL*), atau --folder-link <url>.",
        "Baca semua .xlsx / Google Sheets di subtree REGIONAL*; skip ~$, validation.xlsx, column.json, all-pks.xlsx.",
        "Deteksi duplikat EQUNR per file (kolom B, data mulai baris 4); tulis ulang dari template_changes_equipment.xlsx; hanya baris duplikat di-fill merah (A–D).",
        "Terminate (exit 1) jika kolom EQUNR tidak ditemukan di baris 1.",
        "Tidak menghapus baris; tidak upload ke Drive.",
        "Output lokal: Output/check-lsmw-update-change-equipment/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx",
        "  + summary.xlsx (Regional, File, BarisData, EqunrDuplikat, BarisDuplikatMerah)."
    ].join("\n    "),

    "check-all-plant-pks": [
        "Daftar stasiun per file PKS di folder Protect Y (baris funcloc 4 segmen).",
        "Prompt URL folder Protect Y (subfolder REGIONAL*), atau --protect-y-folder-link <url>.",
        "Satu baris per stasiun: FUNCTLOC DESC. AFTER + work center + planner group (lookup seperti lsmw-func-loc, cwc.csv / planner-group.csv).",
        "Maintenance plant per file: MAINTENANCE PLAN AFTER → MAINTENANCE PLANT → POM → segmen-2 FUNCTIONAL LOCATION AFTER.",
        "QTY func loc: jumlah FUNCTIONAL LOCATION AFTER unik di subtree stasiun (stasiun + semua turunan).",
        "Cost center stasiun: derivasi STRNO (segmen 2 + STAS + 2 digit segmen 4, mis. 1F11STAS01).",
        "Skip ~$ dan all-pks.xlsx.",
        "Output: Output/check-all-plant-pks/{timestamp WIB}.xlsx",
        "  kolom: nama pks | nama regional | maintenance plant | stasiun | work center | planner group | cost center | QTY func loc."
    ].join("\n    "),

    "template-lsmw-check": [
        "Cek panjang karakter isian vs template LSMW (tanpa transformasi data).",
        "Interaktif: pilih template (Protect Y / Equipment / Func Loc) + URL file Drive.",
        "Output: Output/template-lsmw-cheking/."
    ].join("\n    "),

    "equipment-group-check": [
        "Bandingkan equipment group SAP (Excel) vs mapping lokal JSON.",
        "Input: equipment-group-mapping-sap.xlsx, equipment-group-mapping.json.",
        "Output: Output/equipment-group-checking/{timestamp WIB}.xlsx",
        "STATUS: MATCH | UNUSED | NEW."
    ].join("\n    "),

    "lsmw-group-result-checking": [
        "Bandingkan EQUIPMENT GROUP AFTER (Protect Y) vs EQART (output LSMW equipment).",
        "Interaktif: URL folder output LSMW equipment.",
        "Output: Output/lsmw-group-result-checking/{timestamp WIB}/{nama PKS}.xlsx."
    ].join("\n    "),

    "check-equipment-loss": [
        "Bandingkan EQUIPMENT NUMBER + EQKTU BEFORE: LPP Standart vs Protect Y.",
        "Sumber LPP: folder Drive LPP Standart (REGIONAL*, tetap).",
        "Prompt URL folder Protect Y yang akan dicek (subfolder REGIONAL*).",
        "Output: Output/check-equipment-loss/{timestamp WIB}/",
        "  semua-pks.xlsx + per regional/{PKS}.xlsx",
        "Kolom output: EQUIPMENT NUMBER, EQKTU BEFORE, COST CENTER BEFORE, NAMA PKS, STATUS.",
        "STATUS: EXIST | MISSING | ANOMALY | UNUSED.",
        "UNUSED = COST CENTER BEFORE tidak mengandung STAS (menggantikan status lain)."
    ].join("\n    "),

    "equipment-gathering": [
        "Gathering LPP + apply kolom before ke Protect Y dalam satu perintah.",
        "Prompt: URL LPP Resource + URL Protect Y / Applied Last Result (subfolder REGIONAL*).",
        "File Protect Y dipasangkan per nama PKS dengan LPP (matching + isi kolom before).",
        "LPP: EQUIPMENT NUMBER terisi, COST CENTER BEFORE mengandung STAS.",
        "Header LPP: baris kosong di atas didukung; alias FUNCTIONAL LOCATION AFTER2 = AFTER.",
        "Baca/tulis Excel (LPP, Protect Y, gathering) memakai ExcelJS agar key FUNCTLOC konsisten.",
        "Apply Protect Y: kosongkan semua kolom before di setiap baris data PKS, lalu isi dari LPP per funcloc.",
        "REGIONAL dan POM wajib sama di semua baris Protect Y per PKS (dari LPP / nama file).",
        "Matching LPP ke Protect Y (1:1): FUNCTLOC AFTER → BEFORE → EQKTU BEFORE jika FUNCTLOC kosong.",
        "Lalu isi REGIONAL … FUNCTLOC DESC. BEFORE di Protect Y.",
        "Duplikat FUNCTLOC DESC. AFTER: [warn] saja (bukan error). Applied duplikat → first wins matching & apply.",
        "LPP duplikat + Applied unik → buang baris LPP duplikat; keduanya duplikat → pertahankan semua baris LPP.",
        "Proses berhenti jika kolom wajib hilang, tanpa pasangan LPP, atau gagal apply (bukan duplikat funcloc).",
        "Terminal: [match LPP] paired, [apply Protect Y] equipment per PKS.",
        "Output 1: Output/equipment-gathering/{timestamp WIB}/ — all-pks.xlsx + regional/{PKS}.xlsx",
        "Output 2: Output/apply-equipment-gathering-to-last-result/{timestamp WIB}/ — regional/{PKS}.xlsx",
        "Laporan duplikat: EQUIPMENT NUMBER + COST CENTER BEFORE."
    ].join("\n    "),

    "lsmw-tasklist": [
        "Generate file LSMW tasklist (preventive maintenance) dari Protect Y.",
        "Prompt 1: URL folder Drive Protect Y (subfolder REGIONAL*).",
        "Wajib unik per file selalu aktif (skip bundle duplikat HEADER/OPERATION), tanpa prompt pilihan.",
        "Sumber: .xlsx / Google Sheets.",
        "Skip ala lsmw-equipment dinonaktifkan (level, drive unit/accessories, merah, STAS13).",
        "Tetap dipertahankan: dedupe FUNCTIONAL LOCATION AFTER lintas file + mode wajib unik.",
        "Main work center & planner group mengikuti lookup yang sama dengan GEWRK/INGRP lsmw-equipment",
        "  (cwc.csv/planner-group.csv berbasis maintenance plant per file; fallback ke nilai sumber jika lookup kosong).",
        "Kolom cost center sumber: prioritas header (kiri-ke-kanan per tier) COST CENTER AFTER →",
        "  COST CENTER tanpa BEFORE → COST CENTER BEFORE.",
        "Hanya baris yang FUNCTLOC DESC. AFTER mengandung filter dari peralatan_dengan_jam_notif_overhaul.json (prioritas keyword terpanjang).",
        "Task List Group (PLNNR) = {segmen-2}{task_list_group dari JSON}; jika PRS1 dan plant (segmen-2) diawali 3 → suffix PRS2.",
        "Kolom deskripsi HEADER (PREVENTIVE …) memakai key equipment (bukan filter).",
        "Jika beberapa keyword filter cocok dalam satu deskripsi: longest keyword.",
        "Jika JSON punya beberapa entri equipment sama dengan hour_meter berbeda: 1 HEADER +",
        "  (2 × jumlah entri unik) baris OPERATION (tasklist bold + operation per interval HM).",
        "Di OPERATION per interval HM:",
        "  baris tasklist: operation no terisi, sub operation kosong, HM terkait = X.",
        "  baris operation: operation no + sub operation sesuai JSON, semua kolom HM kosong.",
        "Fail-fast: jika kandidat tasklist punya main work center atau planner group kosong,",
        "  proses langsung berhenti dan melaporkan FUNCTIONAL LOCATION AFTER + FUNCTLOC DESC. AFTER.",
        "  Pesan error juga menampilkan cost center raw, effective cost center, dan suffix lookup CSV.",
        "Template: template_lsmw_tasklist.xlsx — sheet HEADER (header baris 1–3, body baris 4+),",
        "  OPERATION (header baris 1–2, body baris 3+),",
        "  SUB_OPERATION (nama kolom baris 1, batas panjang baris 4, body baris 5+); style header dipertahankan.",
        "SUB_OPERATION (hanya baris OPERATION tasklist bold dengan operation_no = 0010):",
        "  PLNNR = Task List Group baris OPERATION terkait; STTAG = tanggal run DD.MM.YYYY (WIB);",
        "  PLNAL = Group Counter baris OPERATION terkait; ENTRY_ACT & FLG_SEL_01 kosong;",
        "  kolom HM (50–20000) = salin X dari baris OPERATION yang sama.",
        "Fail-fast SUB_OPERATION: proses berhenti jika nilai melebihi batas karakter baris 4 template.",
        "1 baris sumber → 1 baris HEADER + 2 atau lebih baris OPERATION (2 per entri JSON unik).",
        "Retry baca/unduh Drive: hingga 4x dengan jeda 1 detik.",
        "Output: Output/lsmw-tasklist/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx + all-pks.xlsx",
        "Urutan file & all-pks.xlsx: numerik REGIONAL (mis. 1…7) lalu prefix file (mis. 1. PKS …, 2. PKS …)."
    ].join("\n    "),

    "lsmw-maintenance-item": [
        "Generate file LSMW maintenance item (preventive maintenance) dari Protect Y.",
        "Prompt 1: URL folder Drive Protect Y (subfolder REGIONAL*).",
        "Prompt 2: pilih subfolder REGIONAL yang diproses (checkbox, default semua terpilih).",
        "Prompt 3: validasi panjang karakter — strict (hentikan jika melebihi batas baris 3 template) atau warning.",
        "Kolom cost center sumber: prioritas header (kiri-ke-kanan per tier) COST CENTER AFTER →",
        "  COST CENTER tanpa BEFORE → COST CENTER BEFORE.",
        "Sumber: .xlsx / Google Sheets; filter & skip mengikuti lsmw-equipment + pemilihan JSON seperti lsmw-tasklist.",
        "Hanya baris yang FUNCTLOC DESC. AFTER mengandung filter dari peralatan_dengan_jam_notif_overhaul.json (prioritas keyword terpanjang).",
        "Dedupe FUNCTIONAL LOCATION AFTER unik lintas file; tanpa prompt wajib unik.",
        "Skip: hanya level funcloc 5 (level lain di-skip), drive unit / accessories weigh bridge, EQUIPMENT GROUP AFTER merah,",
        "  level 5 + cost center station STAS13, funcloc duplikat lintas run.",
        "Template: template_lsmw_maintenance_item.xlsx — baris 1 nama kolom, baris 2 deskripsi, baris 3 batas karakter; body baris 4+ (tanpa kolom TPLNR).",
        "Pemetaan kolom output:",
        "  MPTYP=PM | WSTRA=PALMPB | AUART=PM02 | ILART=PRE | PRIOK=3 | PLNTY=A | PLNAL=1",
        "  PSTXT = {segmen-2 funcloc} + spasi + json.text",
        "  EQUNR = Equipment dari new-equipment-number.xlsx",
        "    key: Planning Plant = segmen2 FUNCTIONAL LOCATION AFTER, Description = FUNCTLOC DESC. AFTER",
        "      setelah typo-fix dan alias mapping-alias-desc-func-loc.xlsx (sama lsmw-equipment, case-insensitive).",
        "    Jika mapping tidak ditemukan maka fallback ke {segmen2}-{FUNCTLOC DESC. AFTER asli dari Protect Y}.",
        "  IWERK/WERGW = segmen-2 funcloc | GEWERK = main work center (cwc.csv) | WPGRP = planner group",
        "  PLNNR = {segmen-2}{task_list_group dari JSON}; jika PRS1 dan plant (segmen-2) diawali 3 → suffix PRS2",
        "Kolom cost center sumber: prioritas header yang sama dengan lsmw-equipment (AFTER → tanpa BEFORE → BEFORE).",
        "Retry baca/unduh Drive: hingga 4x dengan jeda 1 detik.",
        "Output per PKS: Output/lsmw-maintenance-item/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx + all-pks.xlsx di root folder run (luar subfolder REGIONAL).",
        "validation.xlsx di root folder run: Nama Sub Folder, Nama File Excel,",
        "  Jumlah EQUNR Ditemukan, Jumlah EQUNR Tidak Ditemukan, Plant, Valid Func Loc,",
        "  Invalid Func Loc, Invalid Func Loc List (hanya file berhasil diproses).",
        "  EQUNR ditemukan = ada di new-equipment-number.xlsx; tidak ditemukan = fallback {segmen2}-{desc}.",
        "  Valid/Invalid Func Loc hanya baris EQUNR ditemukan: bandingkan FUNCTIONAL LOCATION AFTER vs kolom Functional Loc."
    ].join("\n    "),

    "lsmw-maintenance-plan": [
        "Generate file LSMW maintenance plan (preventive maintenance) dari Protect Y.",
        "Prompt 1: URL folder Drive Protect Y (subfolder REGIONAL*).",
        "Prompt 2: pilih subfolder REGIONAL yang diproses (checkbox, default semua terpilih).",
        "Prompt 3: URL folder Drive output lsmw-measuring-document (folder run berisi REGIONAL N/{PKS}.xlsx).",
        "Prompt 4: validasi kolom wajib — strict atau warning (+ validation.xlsx).",
        "Prompt 5 (hanya strict): checkbox kolom wajib yang tetap terminate (default semua tercentang).",
        "Uncheck kolom = hanya warning + masuk validation.xlsx, tidak terminate.",
        "Batas karakter baris 4 template: selalu terminate di mode strict (di luar checkbox).",
        "Kolom cost center sumber: prioritas header (kiri-ke-kanan per tier) COST CENTER AFTER →",
        "  COST CENTER tanpa BEFORE → COST CENTER BEFORE.",
        "Sumber: .xlsx / Google Sheets; filter & skip mengikuti lsmw-maintenance-item + JSON seperti lsmw-tasklist.",
        "Jumlah baris output = sama lsmw-maintenance-item (hanya filter overhaul JSON); HERI tidak mengurangi jumlah baris.",
        "  Berbeda dengan lsmw-measuring-document: di sana HERI invalid/0/no-match melewati seluruh blok anchor.",
        "Hanya baris yang FUNCTLOC DESC. AFTER mengandung filter dari peralatan_dengan_jam_notif_overhaul.json (prioritas keyword terpanjang).",
        "  (Tidak menulis PLNNR/task_list_group; aturan PRS1→PRS2 hanya di lsmw-tasklist & lsmw-maintenance-item.)",
        "Dedupe FUNCTIONAL LOCATION AFTER unik lintas file; tanpa prompt wajib unik.",
        "Skip: hanya level funcloc 5 (level lain di-skip), drive unit / accessories weigh bridge, EQUIPMENT GROUP AFTER merah,",
        "  level 5 + cost center station STAS13, funcloc duplikat lintas run (sama maintenance-item, tanpa skip HERI).",
        "Template: template_lsmw_maintenance_plan.xlsx — header baris 1–4, body baris 5+; batas karakter di baris 4.",
        "Sumber HERI: sheet HERI pada spreadsheet Data Jam Jalan Peralatan (dibaca otomatis via Sheets API).",
        "Pemetaan kolom output:",
        "  MPTYP=PM | WSTRA=PALMPB",
        "  EQUNR (lookup internal, tidak kolom output) = Equipment dari new-equipment-number.xlsx",
        "    key: Planning Plant = segmen-2 FUNCTIONAL LOCATION AFTER, Description = FUNCTLOC DESC. AFTER",
        "      setelah typo-fix dan alias mapping-alias-desc-func-loc.xlsx (sama lsmw-measuring-document, case-insensitive).",
        "    Jika mapping tidak ditemukan maka fallback ke {segmen2}-{FUNCTLOC DESC. AFTER asli dari Protect Y}.",
        "  MAINTENANCE_ITEM = MaintItem pertama dari maintenance-item-ip18.xlsx (lookup Equipment; 1 item per plan; warning jika >1)",
        "    key utama: EQUNR mapped; fallback {segmen2}-{FUNCTLOC DESC. AFTER asli} jika tidak ada di ip18.",
        "    Strict MAINTENANCE_ITEM: hanya lulus jika ditemukan via key utama (bukan fallback).",
        "  WPTXT = {segmen-2 funcloc} + spasi + json.text",
        "  POINT = Measuring point dari measuring-item-ik07.xlsx (lookup Equipment; first wins)",
        "    key utama: EQUNR mapped; fallback {segmen2}-{FUNCTLOC DESC. AFTER asli} jika tidak ada di ik07.",
        "    Strict POINT: hanya lulus jika ditemukan via key utama (bukan fallback).",
        "  HORIZ=90 | HORIZ_QUALIFIER=% | ABRHO=5 | HUNIT=YR",
        "  SZAEH = counter HERI (plant + FUNCTLOC DESC. AFTER match kolom HERI, case-insensitive); hanya baris filter JSON.",
        "  HERI tidak ditemukan / invalid / kosong / plant tidak ada / nilai 0 → SZAEH = 0 (baris tetap emit, tidak skip).",
        "  Notasi HERI: titik = pemisah ribuan (dihapus), koma = desimal (mis. 1.234,56 → 1234.56).",
        "  Output SZAEH ditulis dengan koma sebagai pemisah desimal (mis. 1234,56).",
        "  Pengecualian sterilizer: HORIZONTAL/VERTICAL STERILIZER NO. N → kolom HERI STERILIZER NO. N.",
        "  Counter Reading (strict): jika EQUNR ada di output measuring-document, SZAEH harus sama persis",
        "    dengan kolom Counter Reading (template_measuring_document.xlsx); EQUNR tidak ada di measuring-document → lolos.",
        "Kolom wajib: EQUNR, MAINTENANCE_ITEM, SZAEH, WPTXT, POINT, Counter Reading (cross-check MD).",
        "validation.xlsx di root folder run: Regional, Nama File PKS, Maintenance Plan, Baris Sumber,",
        "  FUNCLOC, nilai EQUNR/MAINTENANCE_ITEM/POINT/SZAEH, Kolom Bermasalah, Masalah (satu baris per masalah).",
        "Pelanggaran batas karakter dilaporkan dengan FUNCLOC (FUNCTIONAL LOCATION AFTER).",
        "Retry baca/unduh Drive: hingga 4x dengan jeda 1 detik.",
        "Output per PKS: Output/lsmw-maintenance-plan/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx + all-pks.xlsx di root folder run (luar subfolder REGIONAL)."
    ].join("\n    "),

    "lsmw-measuring-point": [
        "Generate file LSMW measuring point (hourmeter) dari Protect Y.",
        "Prompt 1: URL folder Drive Protect Y (subfolder REGIONAL*).",
        "Prompt 2: pilih subfolder REGIONAL yang diproses (checkbox, default semua terpilih).",
        "Prompt 3: validasi panjang karakter — strict (hentikan jika melebihi batas) atau warning.",
        "Kolom cost center sumber: prioritas header (kiri-ke-kanan per tier) COST CENTER AFTER →",
        "  COST CENTER tanpa BEFORE → COST CENTER BEFORE.",
        "Sumber: .xlsx / Google Sheets; header via findDataLayout (FUNCTIONAL LOCATION AFTER).",
        "Filter baris: level funcloc ≥ 5; kecuali deskripsi drive unit / accessories weigh bridge.",
        "Baris dengan EQUIPMENT GROUP AFTER ber-background merah di-skip (sama lsmw-equipment).",
        "Skip tambahan mengikuti lsmw-equipment: dedupe FUNCTIONAL LOCATION AFTER lintas file + item level 5 dengan cost center station STAS13.",
        "Template: template_measuring_point.xlsx — baris 1 nama kolom, 2 deskripsi, 3 batas karakter; body baris 4+.",
        "Pemetaan kolom output:",
        "  MPOTY = IEQ | MPTYP = M | PSORT = HMKM_FRONTEND | ATNAM = ZPM_HOURMETER | CJUMC = 999999 | PYEAC = 8400 (konstanta)",
        "  BEGRU = segmen ke-2 FUNCTIONAL LOCATION AFTER (split \"-\")",
        "  EQUNR = Equipment dari new-equipment-number.xlsx",
        "    key: Planning Plant = segmen2 FUNCTIONAL LOCATION AFTER, Description = FUNCTLOC DESC. AFTER",
        "      setelah typo-fix dan alias mapping-alias-desc-func-loc.xlsx (sama lsmw-equipment, case-insensitive).",
        "    Jika mapping tidak ditemukan maka fallback ke {segmen2}-{FUNCTLOC DESC. AFTER asli dari Protect Y}.",
        "  PTTXT = \"Pencatatan HM\" (konstanta, sama di semua baris)",
        "Output per PKS: Output/lsmw-measuring-point/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx + all-pks.xlsx di root folder run (luar subfolder REGIONAL).",
        "validation.xlsx di root folder run: Nama Sub Folder, Nama File Excel,",
        "  Jumlah EQUNR Ditemukan, Jumlah EQUNR Tidak Ditemukan, Plant, Valid Func Loc,",
        "  Invalid Func Loc, Invalid Func Loc List (hanya file berhasil diproses).",
        "  EQUNR ditemukan = ada di new-equipment-number.xlsx; tidak ditemukan = fallback {segmen2}-{desc}.",
        "  Valid/Invalid Func Loc hanya baris EQUNR ditemukan: bandingkan FUNCTIONAL LOCATION AFTER vs kolom Functional Loc.",
        "Header baris 1–3 disalin dari template (style/warna tetap); data ditulis mulai baris 4.",
        "EQUNR sebenranynya hasil export dari IH08 SAP "
    ].join("\n    "),

    "lsmw-measuring-document": [
        "Generate file LSMW measuring document dari Protect Y + sheet HERI (Google Spreadsheet).",
        "Sumber HERI: sheet HERI pada spreadsheet Data Jam Jalan Peralatan (dibaca otomatis via Sheets API).",
        "Layout HERI: baris 2 kolom D+ = nama alat; kolom C baris 4+ = plant; sel = counter reading.",
        "Notasi angka HERI: titik = pemisah ribuan (dihapus), koma = desimal (mis. 1.234.567 → 1234567; 1.234,56 → 1234.56).",
        "Output Counter Reading ditulis dengan koma sebagai pemisah desimal (mis. 1234,56).",
        "Prompt 1: URL folder Drive Protect Y (subfolder REGIONAL*).",
        "Prompt 2: pilih subfolder REGIONAL yang diproses (checkbox, default semua terpilih).",
        "Prompt 3: validasi kolom wajib — strict (hentikan sesuai kolom tercentang) atau warning (+ validation.xlsx).",
        "Prompt 4 (hanya strict): checkbox kolom terminate — Counter Reading, Equipment Number, Measuring point (default semua).",
        "Uncheck kolom = hanya warning + masuk validation.xlsx, tidak terminate.",
        "Kolom wajib: Counter Reading (HERI root anchor), Equipment Number (mapping EQUNR), Measuring point (ik07).",
        "Counter Reading: anchor dengan counter HERI invalid atau bernilai 0 (termasuk HM Rusak) dilewati seluruh bloknya",
        "  (warning, tidak terminate, tidak masuk output); issue Counter Reading hanya jika plant tidak ditemukan di HERI.",
        "Aturan masuk output (semua harus terpenuhi):",
        "  1) Filter Protect Y umum: level funcloc ≥ 5; skip drive unit / accessories weigh bridge,",
        "     EQUIPMENT GROUP AFTER merah, level 5 + STAS13, dedupe FUNCTIONAL LOCATION AFTER lintas run.",
        "  2) Root anchor: FUNCTLOC DESC. AFTER sama persis dengan nama kolom HERI (case-insensitive)",
        "     untuk plant yang sama (segmen-2 FUNCTIONAL LOCATION AFTER), bukan turunan func loc",
        "     dari root anchor HERI lain. Pengecualian sterilizer: HORIZONTAL/VERTICAL STERILIZER NO. N",
        "     → kolom HERI STERILIZER NO. N (dedupe per deskripsi).",
        "  3) Counter HERI anchor valid dan > 0; invalid/0 → seluruh blok tidak emit.",
        "  4) Turunan: baris lolos filter (1) yang func loc-nya turunan langsung/tidak langsung dari anchor (2+3);",
        "     deskripsi tidak perlu match HERI. Blok atomik: anchor tidak ditulis → seluruh turunan tidak ditulis.",
        "Prompt tambahan: URL folder Drive output measuring-document sebelumnya (untuk Missing Measuring Document).",
        "Missing Measuring Document/ di root folder run: baris output run ini yang EQUNR-nya tidak ada",
        "  di folder Drive tersebut — per PKS REGIONAL N/{nama PKS}.xlsx;",
        "  agregat {REGIONAL N}.xlsx + all-pks.xlsx di root folder Missing (bukan di subfolder REGIONAL).",
        "Template: template_measuring_document.xlsx — header baris 1; body baris 2+.",
        "Pemetaan kolom output:",
        "  No. Urut = auto increment per file PKS mulai 1; agregat {REGIONAL N}.xlsx di-renumber 1..N",
        "  Equipment Number = new-equipment-number.xlsx (lookup sama lsmw-maintenance-item)",
        "  Equipment Description = FUNCTLOC DESC. AFTER (+ alias mapping-alias-desc-func-loc.xlsx bila cocok, sama lsmw-equipment)",
        "  Measuring point = measuring-item-ik07.xlsx (lookup by EQUNR, sama lsmw-maintenance-plan)",
        "  Measurement Date = 09.06.2026 (konstanta) | Measurement Time = 10:00:00 (konstanta)",
        "  Counter Reading = nilai HERI dari root anchor (sama untuk seluruh turunan di bloknya);",
        "    hanya nilai > 0 yang diproses; invalid/0 skip blok anchor",
        "  Unit of Measure = kosong",
        "  Difference = x | Read By = regional | Short Text = IB sd 31 Mei 2026",
        "Output per PKS: Output/lsmw-measuring-document/{timestamp WIB}/REGIONAL N/{nama PKS}.xlsx",
        "Agregat: {REGIONAL N}.xlsx di root folder run (tanpa all-pks.xlsx).",
        "Missing Measuring Document/: REGIONAL N/{nama PKS}.xlsx + {REGIONAL N}.xlsx + all-pks.xlsx",
        "  (EQUNR baru vs folder Drive sebelumnya).",
        "validation.xlsx di root folder run: Regional, Nama PKS, Plant (segmen-2 FUNCTIONAL LOCATION AFTER),",
        "  Counter Reading, Equipment Number, Measuring point, Alasan tidak valid",
        "  (mode strict maupun warning; satu baris per masalah kolom).",
        "File dengan 0 baris match tidak ditulis."
    ].join("\n    "),

    "apply-equipment-gathering-to-last-result": [
        "Apply ulang saja (tanpa gathering) dari hasil equipment-gathering yang sudah ada.",
        "Prompt 1: sumber gathering — Google Drive atau folder lokal (run REGIONAL*).",
        "Prompt 2: URL folder Drive Protect Y Last Result (REGIONAL*).",
        "Kunci baris: FUNCTLOC DESC. AFTER APPLIED (gathering) = FUNCTLOC DESC. AFTER (Protect Y).",
        "Fase 1: kosongkan semua kolom before di seluruh baris data; fase 2: isi dari gathering per funcloc key.",
        "Kolom after Protect Y tidak diubah.",
        "Terminal: jumlah equipment ter-apply vs gagal apply per PKS.",
        "Output: Output/apply-equipment-gathering-to-last-result/{timestamp WIB}/",
        "  per regional/{PKS}.xlsx + all-failed-apply-equipment.xlsx (baris gathering gagal apply)."
    ].join("\n    ")
};

function loadScripts() {
    const pkg = JSON.parse(fs.readFileSync(PACKAGE_JSON, "utf8"));
    return Object.entries(pkg.scripts ?? {}).sort(([a], [b]) =>
        a.localeCompare(b)
    );
}

function printHelp() {
    const scripts = loadScripts();

    console.log("\nPerintah pnpm run — projek-cmms-file\n");
    console.log("Prasyarat umum (kecuali help / equipment-group-check):");
    console.log("  oauth.json + token.json (akses Google Drive).\n");

    const maxName = Math.max(...scripts.map(([name]) => name.length), 4);

    for (const [name, command] of scripts) {
        const desc =
            COMMAND_HELP[name] ??
            `(belum ada deskripsi — jalankan: ${command})`;
        console.log(`  pnpm run ${name.padEnd(maxName)}`);
        console.log(`    ${desc.split("\n").join("\n    ")}\n`);
    }

    console.log("Contoh:");
    console.log("  pnpm run equipment-gathering");
    console.log("  pnpm run help\n");
}

printHelp();
