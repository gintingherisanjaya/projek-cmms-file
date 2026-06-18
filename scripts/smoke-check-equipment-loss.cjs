/**
 * Smoke test logika compare check-equipment-loss (tanpa Drive).
 */
const {
    STATUS,
    isUnusedCostCenter,
    compareEquipmentMaps
} = require("../utils/check_equipment_loss.cjs");

function assert(cond, msg) {
    if (!cond) throw new Error(msg);
}

function row(cc, eqktu = "") {
    return { eqktuBefore: eqktu, costCenterBefore: cc };
}

function main() {
    assert(isUnusedCostCenter(""), "kosong = UNUSED");
    assert(isUnusedCostCenter("1F02XXXX01"), "tanpa STAS = UNUSED");
    assert(!isUnusedCostCenter("1F02STAS01"), "dengan STAS bukan UNUSED");

    const lpp = new Map([
        ["E1", row("1F02STAS01", "eq1")],
        ["E2", row("1F02XXXX01", "eq2")],
        ["E3", row("1F02STAS03", "eq3")]
    ]);
    const py = new Map([
        ["E1", row("1F02STAS01", "eq1")],
        ["E4", row("", "eq4")],
        ["E5", row("1F02STAS05", "eq5")]
    ]);

    const items = compareEquipmentMaps(lpp, py, "PKS TEST");
    const byEquip = new Map(items.map(i => [i.equipmentNumber, i]));

    assert(byEquip.get("E1").status === STATUS.EXIST, "E1 EXIST");
    assert(byEquip.get("E1").costCenterBefore === "1F02STAS01", "E1 CC");

    assert(byEquip.get("E2").status === STATUS.UNUSED, "E2 UNUSED bukan MISSING");
    assert(byEquip.get("E3").status === STATUS.MISSING, "E3 MISSING");

    assert(byEquip.get("E4").status === STATUS.UNUSED, "E4 UNUSED bukan ANOMALY");
    assert(byEquip.get("E5").status === STATUS.ANOMALY, "E5 ANOMALY");

    console.log("ok: isUnusedCostCenter");
    console.log("ok: EXIST dengan STAS");
    console.log("ok: UNUSED (LPP tanpa STAS)");
    console.log("ok: MISSING dengan STAS");
    console.log("ok: UNUSED (Protect Y tanpa STAS)");
    console.log("ok: ANOMALY dengan STAS");
    console.log("\nSemua smoke check-equipment-loss lulus.");
}

main();
