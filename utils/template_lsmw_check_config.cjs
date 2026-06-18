const path = require("path");

const CHECK_MODES = {
    "protect-y": {
        id: "protect-y",
        label: "Template Protect Y",
        description: "Semua kolom sumber yang dipetakan (hanya file sumber)",
        templatePath: path.join(__dirname, "..", "template_equipment.xlsx"),
        templateHeaderRow: 3,
        templateMaxLenRow: 6,
        scope: "all-source-columns",
        outputRoot: null,
        outputHeaderRow: null,
        outputMaxLenRow: null,
        outputDataStartRow: null,
        outputKeyColumn: null,
        outputDescColumn: null
    },
    equipment: {
        id: "equipment",
        label: "Template Equipment",
        description: "FUNCTLOC DESC (sumber + output LSMW equipment)",
        templatePath: path.join(__dirname, "..", "template_equipment.xlsx"),
        templateHeaderRow: 3,
        templateMaxLenRow: 6,
        scope: "funcloc-desc-only",
        sourceDescTemplateCol: "SHTXT*",
        outputRoot: path.join(
            __dirname,
            "..",
            "Output",
            "1. LSMW Create Equipment V.O"
        ),
        outputHeaderRow: 3,
        outputMaxLenRow: 6,
        outputDataStartRow: 7,
        outputKeyColumn: "TPLNR",
        outputDescColumn: "SHTXT*"
    },
    "func-loc": {
        id: "func-loc",
        label: "Template Func Loc",
        description: "FUNCTLOC DESC (sumber + output LSMW func loc)",
        templatePath: path.join(
            __dirname,
            "..",
            "template_create_func_loc_lsmw1.xlsx"
        ),
        templateHeaderRow: 1,
        templateMaxLenRow: 4,
        scope: "funcloc-desc-only",
        sourceDescTemplateCol: "PLTXT",
        outputRoot: path.join(
            __dirname,
            "..",
            "Output",
            "2. LSMW Create Functional Location V1"
        ),
        outputHeaderRow: 1,
        outputMaxLenRow: 4,
        outputDataStartRow: 5,
        outputKeyColumn: "STRNO",
        outputDescColumn: "PLTXT"
    }
};

module.exports = { CHECK_MODES };
