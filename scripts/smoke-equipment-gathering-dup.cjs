/**
 * Smoke: duplikat FUNCTLOC (tanpa Drive).
 *   node scripts/smoke-equipment-gathering-dup.cjs
 */
const {
    extractAppliedMatchCandidates,
    findLppGatheringLayout,
    resolveSourceColumnIndices,
    listMissingSourceColumns,
    PROTECT_Y_APPLY_HEADERS,
    OUTPUT_HEADERS,
    FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
} = require("../utils/equipment_gathering_columns.cjs");
const { applyGatheringToWorksheet } = require("../utils/apply_equipment_gathering_protecty.cjs");
const {
    filterOutputRowsByDuplicateAppliedKey,
    buildGatheringIndexFromOutputRows
} = require("../utils/apply_equipment_gathering_io.cjs");
const { normalizeFunclocDesc } = require("../utils/string_similarity.cjs");
const {
    matchLppRowsToApplied,
    resolveLppMatchFields
} = require("../utils/equipment_gathering_match.cjs");
const { readSpreadsheetRows } = require("../utils/equipment_excel_io.cjs");
const ExcelJS = require("exceljs");
const fs = require("fs");
const os = require("os");
const path = require("path");

const COL_APPLIED = OUTPUT_HEADERS.indexOf("FUNCTLOC DESC. AFTER APPLIED");
const COL_CC_AFTER = OUTPUT_HEADERS.indexOf("COST CENTER AFTER");
const COL_CC_BEFORE = OUTPUT_HEADERS.indexOf("COST CENTER BEFORE");
const COL_EQKTU_BEFORE = OUTPUT_HEADERS.indexOf("EQKTU BEFORE");
const COL_FUNCLOC_AFTER = OUTPUT_HEADERS.indexOf(
    FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
);
const COL_FUNCLOC_BEFORE = OUTPUT_HEADERS.indexOf("FUNCTLOC DESC. BEFORE");

function assert(cond, msg) {
    if (!cond) {
        console.error(`FAIL: ${msg}`);
        process.exit(1);
    }
    console.log(`ok: ${msg}`);
}

// 1) Applied: 2 baris same funcloc → 1 kandidat
{
    const header = [
        "COST CENTER AFTER",
        FUNCTLOC_DESC_AFTER_LEVEL123_HEADER,
        "EQUIPMENT NUMBER"
    ];
    const data = [
        ["CC1", "PUMP A", "EQ1"],
        ["CC2", "PUMP A", "EQ2"]
    ];
    const r = extractAppliedMatchCandidates(header, data);
    assert(r.ok && r.candidates.length === 1, "Applied dedupe → 1 kandidat");
}

// 2) LPP duplikat, Applied count=1 → 1 baris output
{
    const appliedCounts = new Map([[normalizeFunclocDesc("PUMP A"), 1]]);
    const outputRows = [
        Object.fromEntries(OUTPUT_HEADERS.map(h => [h, ""])),
        Object.fromEntries(OUTPUT_HEADERS.map(h => [h, ""]))
    ].map((_, i) => {
        const row = OUTPUT_HEADERS.map(() => "");
        row[COL_APPLIED] = "PUMP A";
        row[0] = `eq${i}`;
        return row;
    });
    const { rows, removedLppDuplicates } = filterOutputRowsByDuplicateAppliedKey(
        outputRows,
        appliedCounts
    );
    assert(rows.length === 1 && removedLppDuplicates === 1, "LPP dup + Applied unik → 1 baris");
}

// 3) LPP duplikat, Applied count=2 → 2 baris, index first wins
{
    const keyB = normalizeFunclocDesc("PUMP B");
    const appliedCounts = new Map([[keyB, 2]]);
    const outputRows = [0, 1].map(i => {
        const row = OUTPUT_HEADERS.map(() => "");
        row[COL_APPLIED] = "PUMP B";
        row[0] = `row${i}`;
        return row;
    });
    const { rows, removedLppDuplicates } = filterOutputRowsByDuplicateAppliedKey(
        outputRows,
        appliedCounts
    );
    assert(
        rows.length === 2 && removedLppDuplicates === 0,
        "LPP dup + Applied dup → keep all"
    );
    const gathered = buildGatheringIndexFromOutputRows(rows);
    assert(
        gathered.duplicateKeys.includes(keyB) && gathered.index.size === 1,
        "Index first wins dengan duplicateKeys warn"
    );
}

// 5) LPP SEI GALUH: baris kosong + FUNCTIONAL LOCATION AFTER2
{
    const seiGaluhHeader = [
        "REGIONAL",
        "POM",
        "EQUIPMENT NUMBER",
        "MAINTENANCE PLANT",
        "COMPANY CODE",
        "COST CENTER BEFORE",
        "EQKTU BEFORE",
        "OBJECT NUMBER",
        "EQUIPMENT CATEGORY",
        "EQUIPMENT GROUP BEFORE",
        "CONSTRUCTION YEAR",
        "GROES",
        "MANUFACTURE",
        "MEASURING POINT",
        "MeasurmntRangeUnit",
        "MAINTENANCE PLANNER GROUP",
        "WORK CENTER",
        "FUNCTIONAL LOCATION BEFORE",
        "FUNCTLOC DESC. BEFORE",
        "COST CENTER AFTER",
        "FUNCTIONAL LOCATION AFTER2",
        FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
    ];
    const rawRows = [[], [], seiGaluhHeader];
    const layout = findLppGatheringLayout(rawRows);
    assert(layout.headerRowIndex === 2, "LPP header baris 3 (index 2, AFTER2)");
    const indices = resolveSourceColumnIndices(seiGaluhHeader);
    const missing = listMissingSourceColumns(indices);
    assert(
        missing.length === 0,
        `LPP SEI GALUH kolom lengkap (missing: ${missing.join(", ")})`
    );
}

(async () => {
    // 4) readSpreadsheetRows: dua kali baca file sama → rawRows identik
    {
        const tmp = path.join(
            os.tmpdir(),
            `smoke-excel-io-${Date.now()}.xlsx`
        );
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Data");
        ws.addRow(["COST CENTER AFTER", FUNCTLOC_DESC_AFTER_LEVEL123_HEADER]);
        ws.addRow(["CC1", "PUMP TEST"]);
        await wb.xlsx.writeFile(tmp);

        const a = await readSpreadsheetRows(tmp);
        const b = await readSpreadsheetRows(tmp);
        fs.unlinkSync(tmp);

        assert(
            a.ok &&
                b.ok &&
                JSON.stringify(a.rawRows) === JSON.stringify(b.rawRows),
            "readSpreadsheetRows konsisten antar baca"
        );
    }

    // 6) Apply: clear semua before, lalu fill hanya baris yang match
    {
        const pyHeader = [
            ...PROTECT_Y_APPLY_HEADERS,
            "FUNCTIONAL LOCATION AFTER",
            FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
        ];
        const idxRegional = pyHeader.indexOf("REGIONAL");
        const idxFuncloc = pyHeader.indexOf(
            FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
        );

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Data");
        ws.addRow(pyHeader);
        ws.addRow(
            pyHeader.map((_, i) => {
                if (i === idxRegional) return "OLD_MATCH";
                if (i === idxFuncloc) return "PUMP A";
                return "";
            })
        );
        ws.addRow(
            pyHeader.map((_, i) => {
                if (i === idxRegional) return "OLD_UNMATCH";
                if (i === idxFuncloc) return "PUMP B";
                return "";
            })
        );
        ws.addRow(
            pyHeader.map((_, i) => {
                if (i === idxRegional) return "OLD_NOKEY";
                if (i === idxFuncloc) return "";
                return "";
            })
        );

        /** @type {Record<string, unknown>} */
        const gatherRow = {};
        for (const h of PROTECT_Y_APPLY_HEADERS) {
            gatherRow[h] = h === "REGIONAL" ? "MATCHED_REGIONAL" : "X";
        }

        const idxPom = pyHeader.indexOf("POM");
        const stats = applyGatheringToWorksheet(
            ws,
            new Map([["PUMP A", gatherRow]]),
            { regional: "REG III", pom: "PKS TEST" }
        );
        assert(stats.clearedBeforeRows === 3, "clear before: 3 baris data");

        const plain = (rowNum, colIdx) => {
            const v = ws.getRow(rowNum).getCell(colIdx + 1).value;
            return v === null || v === undefined ? "" : String(v).trim();
        };

        assert(
            plain(2, idxRegional) === "REG III" &&
                plain(3, idxRegional) === "REG III" &&
                plain(4, idxRegional) === "REG III",
            "REGIONAL sama di semua baris"
        );
        assert(
            plain(2, idxPom) === "PKS TEST" &&
                plain(3, idxPom) === "PKS TEST" &&
                plain(4, idxPom) === "PKS TEST",
            "POM sama di semua baris"
        );
        assert(stats.regionalPomFilledRows === 3, "REGIONAL/POM diisi 3 baris");
        assert(plain(3, idxRegional) === "REG III", "baris unmatched tetap REGIONAL");
    }

    // 7) Apply baris > 5000: clear before + REGIONAL tetap diisi
    {
        const pyHeader = [
            ...PROTECT_Y_APPLY_HEADERS,
            "FUNCTIONAL LOCATION AFTER",
            FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
        ];
        const idxRegional = pyHeader.indexOf("REGIONAL");
        const idxEq = pyHeader.indexOf("EQUIPMENT NUMBER");
        const idxFuncloc = pyHeader.indexOf(
            FUNCTLOC_DESC_AFTER_LEVEL123_HEADER
        );

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Data");
        ws.addRow(pyHeader);
        const farRow = 5100;
        ws.getRow(farRow).getCell(idxRegional + 1).value = "OLD_FAR";
        ws.getRow(farRow).getCell(idxEq + 1).value = "EQ_FAR";
        ws.getRow(farRow).getCell(idxFuncloc + 1).value = "PUMP FAR";

        const stats = applyGatheringToWorksheet(ws, new Map(), {
            regional: "REG FAR",
            pom: "PKS FAR"
        });
        assert(stats.clearedBeforeRows >= farRow - 1, "clear before mencakup baris > 5000");

        const plain = (rowNum, colIdx) => {
            const v = ws.getRow(rowNum).getCell(colIdx + 1).value;
            return v === null || v === undefined ? "" : String(v).trim();
        };

        assert(plain(farRow, idxEq) === "", "baris >5000 tanpa match: equipment dikosongkan");
        assert(plain(farRow, idxRegional) === "REG FAR", "baris >5000: REGIONAL diisi");
    }

    // 8) Matching: FUNCTLOC kosong → fallback EQKTU BEFORE
    {
        const lppRow = OUTPUT_HEADERS.map(() => "");
        lppRow[COL_CC_BEFORE] = "1F02STAS01";
        lppRow[COL_EQKTU_BEFORE] = "ELECTRIC PUMP NO.3";

        const fields = resolveLppMatchFields(
            lppRow,
            COL_FUNCLOC_AFTER,
            COL_FUNCLOC_BEFORE,
            COL_CC_AFTER,
            COL_CC_BEFORE,
            COL_EQKTU_BEFORE
        );
        assert(
            fields && String(fields.desc).includes("ELECTRIC PUMP"),
            "resolveLppMatchFields pakai EQKTU BEFORE"
        );

        const candidates = [
            {
                costCenterAfter: "1F02STAS01",
                funclocDescAfter: "ELECTRIC PUMP NO.3"
            }
        ];
        const matches = matchLppRowsToApplied(
            [lppRow],
            candidates,
            COL_CC_AFTER,
            COL_FUNCLOC_AFTER,
            COL_CC_BEFORE,
            COL_FUNCLOC_BEFORE,
            COL_EQKTU_BEFORE
        );
        assert(matches.has(0), "matchLppRowsToApplied via EQKTU BEFORE");
    }

    console.log("\nSemua smoke duplikat FUNCTLOC lulus.");
})().catch(err => {
    console.error(err);
    process.exit(1);
});
