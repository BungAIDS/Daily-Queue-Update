"use strict";

const assert = require("assert");
const sim = require("./order_similarity.js");

const CORE = {
  Design: "BC-220",
  Size: "220",
  Arrangement: "A/4",
  Class: "II",
  Wheel: "Backward Inclined",
  "% Width": "100",
  Rotation: "CW",
  Discharge: "UB",
  "Motor Pos": "W",
  "Design Temp": "70 F",
  "Max Temp": "180 F",
  "Special Temp": "250 F",
};

function component(name, attributes = {}) {
  return { n: name, k: 1, p: 0, a: attributes, r: [], i: [1], s: [] };
}

function order(extra = {}) {
  const entry = {
    c: "Customer A",
    sp: Object.entries(CORE),
    it: [[1, "Wheel", "", "", "", "WHEEL", ["WHEEL"]]],
    cp: [component("WHEEL", {
      material: "CARBON STEEL", wheel_feature: "BACKWARD INCLINED",
    })],
  };
  return Object.assign(entry, extra);
}

function copy(value) { return JSON.parse(JSON.stringify(value)); }
function close(actual, expected, message) {
  assert.ok(Math.abs(actual - expected) < 1e-9,
    `${message}: expected ${expected}, got ${actual}`);
}

function testIdenticalConstructionIsOne() {
  const a = order(), b = copy(a);
  b.c = "Completely Different Customer";
  b.bd = { pr: "$999,999", no: "ship by a different carrier" };
  const got = sim.orderSimilarity(a, b);
  close(got.score, 1, "fully evidenced identical construction");
  close(got.coverage, 1, "identical evidence coverage");
}

function testPublishedWeightBudgetIsExactlyOne() {
  const core = sim.CORE_FIELDS.reduce((total, field) => total + field.weight, 0);
  const slots = sim.COMPONENT_SLOTS.reduce((total, slot) => total + slot.weight, 0);
  close(core, 0.50, "core budget");
  close(slots, 0.50, "component budgets");
  close(core + slots, 1, "whole construction budget");
  for (const [group, expected] of Object.entries(sim.GROUP_TOTALS)) {
    const got = group === "core" ? core : sim.COMPONENT_SLOTS
      .filter(slot => slot.group === group)
      .reduce((total, slot) => total + slot.weight, 0);
    close(got, expected, `${group} group budget`);
  }
}

function testOneExtraDamperCostsExactlyItsCap() {
  const a = order(), b = copy(a);
  b.cp.push(component("DAMPER", {
    quantity: "1", operation: "AUTOMATIC", damper_type: "OUTLET",
    manufacturer: "EXAMPLE", model: "D-100", fail_position_upon_loss_of_power: "OPEN",
    fail_position_upon_loss_of_signal: "CLOSED", micro_switch_qty: "2",
  }));
  const ab = sim.orderSimilarity(a, b), ba = sim.orderSimilarity(b, a);
  close(ab.score, 0.98, "one additional damper");
  close(ba.score, 0.98, "damper comparison symmetry");
}

function testCommercialAndUnclassifiedLinesDoNotMoveWholeScore() {
  const a = order(), b = copy(a);
  b.cp.push(component("SHIP TO", { shipping_instruction: "CALL CUSTOMER" }));
  b.cp.push(component("Miscellaneous prose not yet classified", { note: "different" }));
  close(sim.orderSimilarity(a, b).score, 1,
    "shipping and unclassified prose must not change whole construction");
}

function testPreviewAttributeScopeMatchesConstructionScore() {
  const attrs = sim.scoredAttrs(component("MOTOR", {
    quantity: "2", motor_hp: "40", enclosure: "TEFC",
    shipping_instruction: "CALL CUSTOMER", note: "commercial note",
    vendor_quote: "VQ-123",
  }));
  assert.deepStrictEqual(attrs, { motor_hp: "40", enclosure: "TEFC" },
    "preview highlighting must expose construction attributes only");
  close(sim.valueSimilarity("TEFC | 460V", "460V, TEFC"), 1,
    "attribute comparison should ignore list order and separators");
}

function testIndependentCoreWeights() {
  const a = order();
  for (const [label, expected] of [["Design", 0.90], ["Size", 0.90], ["% Width", 0.96]]) {
    const b = copy(a);
    const pair = b.sp.find(x => x[0] === label);
    pair[1] = "DIFFERENT";
    close(sim.orderSimilarity(a, b).score, expected, `${label} independent weight`);
  }
}

function testDifferentDesignDoesNotTreatRawSizeCodeAsComparable() {
  const a = order(), b = copy(a);
  b.sp.find(x => x[0] === "Design")[1] = "DIFFERENT DESIGN";
  b.sp.find(x => x[0] === "Size")[1] = "DIFFERENT FORMAT";
  const got = sim.orderSimilarity(a, b);
  close(got.coverage, 0.90, "cross-design raw size evidence");
  assert.ok(got.score > 0.80,
    "different size notation must not count as a second definite mismatch");
  assert.ok(got.differences.includes("Size not compared across different designs"));
}

function testRequiredAttributesStayOnSelectedComponent() {
  const target = component("DAMPER", { operation: "AUTOMATIC", damper_type: "OUTLET" });
  const source = order({ cp: [component("WHEEL"), target] });
  const wrong = order({ cp: [
    component("WHEEL"),
    component("DAMPER", { operation: "MANUAL", damper_type: "OUTLET" }),
    component("MOTOR", { operation: "AUTOMATIC" }),
  ] });
  const wholeWrong = sim.orderSimilarity(source, wrong);
  assert.strictEqual(
    sim.focusedSimilarity(wholeWrong, target, wrong, new Set(["operation=AUTOMATIC"])),
    null,
    "an attribute on MOTOR must not satisfy a required DAMPER attribute",
  );

  const right = copy(wrong);
  right.cp[1].a.operation = "AUTOMATIC";
  const focused = sim.focusedSimilarity(
    sim.orderSimilarity(source, right), target, right, new Set(["operation=AUTOMATIC"]));
  assert.ok(focused && focused.score >= 0 && focused.score <= 1,
    "same-component required attribute should yield a bounded focused score");
}

function testFocusedMatchIdentifiesTheExactCandidateComponent() {
  const target = component("DAMPER", {
    operation: "AUTOMATIC", damper_type: "OUTLET",
  });
  const source = order({ cp: [component("WHEEL"), target] });
  const manual = component("DAMPER", {
    operation: "MANUAL", damper_type: "OUTLET",
  });
  const automatic = component("DAMPER", {
    operation: "AUTOMATIC", damper_type: "OUTLET",
  });
  const candidate = order({ cp: [component("WHEEL"), manual, automatic] });
  const focused = sim.focusedSimilarity(
    sim.orderSimilarity(source, candidate), target, candidate,
    new Set(["operation=AUTOMATIC"]),
  );
  assert.strictEqual(focused.candidateComponent, automatic,
    "preview green highlight must follow the candidate chosen by focused ranking");
}

function testCombinedFocusRequiresEverySelectedComponent() {
  const wheel = component("WHEEL", { material: "CARBON STEEL" });
  const motor = component("MOTOR", { motor_hp: "40", enclosure: "TEFC" });
  const source = order({ cp: [wheel, motor] });
  const requirements = [
    { component: wheel, pins: new Set() },
    { component: motor, pins: new Set() },
  ];
  const missingMotor = order({ cp: [copy(wheel)] });
  assert.strictEqual(sim.combinedFocusedSimilarity(
    sim.orderSimilarity(source, missingMotor), requirements, missingMotor), null,
  "every selected component must exist on the candidate order");

  const complete = order({ cp: [copy(wheel), copy(motor)] });
  const focused = sim.combinedFocusedSimilarity(
    sim.orderSimilarity(source, complete), requirements, complete);
  assert.ok(focused && focused.matches.length === 2,
    "all selected components should contribute to one combination score");
}

function testCombinedAttributesStayTiedToTheirComponents() {
  const damper = component("DAMPER", {
    operation: "AUTOMATIC", damper_type: "OUTLET",
  });
  const motor = component("MOTOR", { enclosure: "TEFC" });
  const source = order({ cp: [damper, motor] });
  const requirements = [
    { component: damper, pins: new Set([
      "operation=AUTOMATIC", "damper_type=OUTLET",
    ]) },
    { component: motor, pins: new Set(["enclosure=TEFC"]) },
  ];
  const wrong = order({ cp: [
    component("DAMPER", {
      operation: "MANUAL", damper_type: "OUTLET", enclosure: "TEFC",
    }),
    component("MOTOR", { operation: "AUTOMATIC", enclosure: "ODP" }),
  ] });
  assert.strictEqual(sim.combinedFocusedSimilarity(
    sim.orderSimilarity(source, wrong), requirements, wrong), null,
    "attributes on different components must not satisfy the combination");

  const partial = order({ cp: [
    component("DAMPER", { operation: "AUTOMATIC", damper_type: "INLET" }),
    component("MOTOR", { enclosure: "TEFC" }),
  ] });
  assert.strictEqual(sim.combinedFocusedSimilarity(
    sim.orderSimilarity(source, partial), requirements, partial), null,
    "every selected attribute on one component must match together");
}

function testCombinedDuplicateSelectionsNeedDistinctCandidates() {
  const first = component("DAMPER", { location: "INLET" });
  const second = component("DAMPER", { location: "OUTLET" });
  const source = order({ cp: [first, second] });
  const requirements = [
    { component: first, pins: new Set(["location=INLET"]) },
    { component: second, pins: new Set(["location=OUTLET"]) },
  ];
  const oneDamper = order({ cp: [component("DAMPER", {
    location: "INLET | OUTLET",
  })] });
  assert.strictEqual(sim.combinedFocusedSimilarity(
    sim.orderSimilarity(source, oneDamper), requirements, oneDamper), null,
  "one candidate component cannot satisfy two selected component instances");
}

function testSparseEvidenceCannotLookIdentical() {
  const a = { sp: [["Design", "BC-220"]], it: [], cp: [] };
  const b = copy(a);
  const got = sim.orderSimilarity(a, b);
  assert.ok(got.score < 1, "unknown construction must prevent a 1.000 score");
  assert.ok(got.coverage < 0.2, "sparse construction must report low evidence");
}

function testAlwaysBounded() {
  const values = [
    sim.orderSimilarity(order(), order()).score,
    sim.orderSimilarity(order(), { sp: [], it: [], cp: [] }).score,
    sim.componentSimilarity(component("DAMPER"), component("MOTOR")).score,
  ];
  for (const value of values)
    assert.ok(value >= 0 && value <= 1, `score outside [0,1]: ${value}`);
}

function main() {
  testPublishedWeightBudgetIsExactlyOne();
  testIdenticalConstructionIsOne();
  testOneExtraDamperCostsExactlyItsCap();
  testCommercialAndUnclassifiedLinesDoNotMoveWholeScore();
  testPreviewAttributeScopeMatchesConstructionScore();
  testIndependentCoreWeights();
  testDifferentDesignDoesNotTreatRawSizeCodeAsComparable();
  testRequiredAttributesStayOnSelectedComponent();
  testFocusedMatchIdentifiesTheExactCandidateComponent();
  testCombinedFocusRequiresEverySelectedComponent();
  testCombinedAttributesStayTiedToTheirComponents();
  testCombinedDuplicateSelectionsNeedDistinctCandidates();
  testSparseEvidenceCannotLookIdentical();
  testAlwaysBounded();
  console.log("All order similarity tests passed.");
}

main();
