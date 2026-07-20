/* GL Queue construction similarity.
 *
 * The score is deliberately semantic and bounded: 0.000..1.000 describes
 * physical construction likeness.  SolidWorks 3D availability is a separate
 * filter.
 * Customer, price, shipping, documents, and other commercial metadata never
 * participate.  Keep this file standalone: order_explorer.py embeds this exact
 * source into the self-contained HTML, while test_order_similarity.js requires
 * it directly under Node.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.GLQSimilarity = api;
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const UNKNOWN_SCORE = 0.5;
  const FOCUS_WHOLE_WEIGHT = 0.65;
  const GROUP_TOTALS = {
    core: 0.50,
    construction: 0.25,
    motor: 0.15,
    accessories: 0.10,
  };

  /* Design, Size, and % Width are independently weighted CBC values.  Until
   * quote-run integration supplies a physical size, raw Size codes are only
   * comparable within the same Design. */
  const CORE_FIELDS = [
    { label: "Design", weight: 0.10 },
    { label: "Size", weight: 0.10 },
    { label: "Arrangement", weight: 0.06 },
    { label: "Class", weight: 0.04 },
    { label: "Wheel", weight: 0.05 },
    { label: "% Width", weight: 0.04 },
    { label: "Rotation", weight: 0.025 },
    { label: "Discharge", weight: 0.025 },
    { label: "Motor Pos", weight: 0.02 },
    { label: "Design Temp", weight: 0.013333 },
    { label: "Max Temp", weight: 0.013333 },
    { label: "Special Temp", weight: 0.013334 },
  ];

  /* Each slot has a hard cap.  A damper can therefore change at most 0.020 of
   * a whole-order score, regardless of how many damper attributes were parsed.
   * The slot weights sum to the three non-core group budgets above. */
  const COMPONENT_SLOTS = [
    { id: "damper", group: "accessories", weight: 0.020,
      terms: ["DAMPER", "IVC", "INLET VANE", "VOLUME CONTROL"],
      attrs: ["damper_", "ivc_"] },
    { id: "actuator", group: "accessories", weight: 0.015,
      terms: ["ACTUATOR", "POSITIONER"], attrs: ["actuator_", "positioner_"] },
    { id: "silencer", group: "accessories", weight: 0.010,
      terms: ["SILENCER", "MUFFLER"], attrs: ["silencer_"] },
    { id: "door-drain", group: "accessories", weight: 0.010,
      terms: ["ACCESS DOOR", "INSPECTION DOOR", "DRAIN", "PRESSURE TAP"],
      attrs: ["door_"] },
    { id: "base-isolation", group: "accessories", weight: 0.020,
      terms: ["UNITARY BASE", "VIBRATION BASE", "STRUCTURAL BASE", "BASE FRAME",
              "T-RAIL", "ISOLATOR", "ISOLATION"],
      attrs: ["unitary_base_", "isolation_"] },
    { id: "cover-screen-guard", group: "accessories", weight: 0.020,
      terms: ["WEATHER COVER", "SCREEN", "GUARD"],
      attrs: ["weather_cover_", "screen_", "guard_"] },
    { id: "flex-connector", group: "accessories", weight: 0.005,
      terms: ["FLEX CONNECTOR", "FLEXIBLE CONNECTOR", "EXPANSION JOINT"],
      attrs: ["flex_connector_"] },

    { id: "motor", group: "motor", weight: 0.080,
      terms: ["MOTOR"], attrs: ["motor_"] },
    { id: "drive-vfd", group: "motor", weight: 0.040,
      terms: ["DRIVE", "VFD", "VARIABLE FREQUENCY"],
      attrs: ["drive_", "vfd_"] },
    { id: "coupling-belt", group: "motor", weight: 0.030,
      terms: ["COUPLING", "BELT", "SHEAVE", "BUSHING"],
      attrs: ["coupling_", "belt_"] },

    { id: "wheel", group: "construction", weight: 0.045,
      terms: ["WHEEL"], attrs: ["wheel_"] },
    { id: "housing", group: "construction", weight: 0.035,
      terms: ["HOUSING"], attrs: ["housing_"] },
    { id: "inlet", group: "construction", weight: 0.025,
      terms: ["INLET"], attrs: ["inlet_"] },
    { id: "outlet", group: "construction", weight: 0.025,
      terms: ["OUTLET", "DISCHARGE"], attrs: ["outlet_"] },
    { id: "shaft-bearing", group: "construction", weight: 0.035,
      terms: ["SHAFT", "BEARING", "SEAL", "SLEEVE", "COOLER"],
      attrs: ["shaft_", "bearing_"] },
    { id: "materials", group: "construction", weight: 0.035,
      terms: ["MATERIAL", "STAINLESS", "ALUMINUM"], attrs: ["material_"] },
    { id: "finish", group: "construction", weight: 0.030,
      terms: ["COATING", "PAINT", "PRIMER", "EPOXY", "LINING", "INSULATION"],
      attrs: ["coating_", "lining_", "insulation_"] },
    { id: "special", group: "construction", weight: 0.020,
      terms: ["SPECIAL CONSTRUCTION", "SPARK", "SPLIT", "FLANGE"],
      attrs: ["special_construction_", "spark_", "split_", "flange_"] },

    /* Unclassified lines remain searchable when explicitly selected, but do
     * not silently change a whole-order score until they are classified. */
    { id: "other-accessory", group: "accessories", weight: 0,
      terms: [], attrs: [], fallback: true },
  ];

  const IGNORED_ATTR = /^(?:applies_to|component|description|drawing_|inquiry_num|job_number|label_|note(?:_|$)|quote_number|referenced_po|related_nameplate_|ship_|shipping_|used_on|vendor_quote|warranty_)/i;
  const IGNORED_COMPONENT = /(?:\bBASE FAN\b|\bPERCENT WIDTH\b|\bRUN TEST\b|\bINSPECTION\b|\bDRAWINGS?\b|\bTRANSMITTAL\b|\bNAMEPLATE\b|\bLABEL\b|\bMARK\b|\bSHIP(?:PING)?\b|\bFREIGHT\b|\bCRAT(?:E|ING)\b|\bPACKING\b|\bTRUCKING\b|\bSKID\b|\bPAYMODE\b|\bSLOW PAY\b|\bINVOICE\b|\bWARRANTY\b)/i;

  function clamp01(n) { return Math.max(0, Math.min(1, Number(n) || 0)); }
  function norm(v) {
    const s = String(v == null ? "" : v).trim().toUpperCase()
      .replace(/[^A-Z0-9.%+/-]+/g, " ").replace(/\s+/g, " ").trim();
    return ["", "N/A", "NA", "NONE", "UNKNOWN"].includes(s) ? "" : s;
  }
  function specMap(entry) {
    const out = {};
    for (const pair of entry.sp || []) out[pair[0]] = pair[1];
    return out;
  }
  function flattenComponents(entry) {
    const out = [];
    const walk = list => {
      for (const c of list || []) { out.push(c); walk(c.s); }
    };
    walk(entry.cp || []);
    return out;
  }
  const PROFILE_CACHE = new WeakMap();
  function componentSlot(component) {
    const name = norm(component && component.n);
    if (!name || IGNORED_COMPONENT.test(name)) return null;
    const keys = Object.keys((component && component.a) || {});
    for (const slot of COMPONENT_SLOTS) {
      if (slot.fallback) continue;
      if (slot.terms.some(t => name.includes(t))) return slot;
      if (slot.attrs.some(prefix => keys.some(k => k.toLowerCase().startsWith(prefix))))
        return slot;
    }
    return COMPONENT_SLOTS[COMPONENT_SLOTS.length - 1];
  }
  function slotsForComponents(components) {
    const out = {};
    for (const slot of COMPONENT_SLOTS) out[slot.id] = [];
    for (const c of components) {
      const slot = componentSlot(c);
      if (slot) out[slot.id].push(c);
    }
    return out;
  }
  function orderProfile(entry) {
    if (entry && PROFILE_CACHE.has(entry)) return PROFILE_CACHE.get(entry);
    const components = flattenComponents(entry || {});
    const profile = {
      specs: specMap(entry || {}),
      components,
      slots: slotsForComponents(components),
      observed: !!(((entry && entry.it) || []).length || components.length),
    };
    if (entry && typeof entry === "object") PROFILE_CACHE.set(entry, profile);
    return profile;
  }
  function componentsBySlot(entry) { return orderProfile(entry).slots; }
  function componentsNamed(entry, name) {
    const wanted = norm(name);
    return orderProfile(entry).components.filter(c => norm(c.n) === wanted);
  }
  function attrParts(value) {
    const parts = String(value == null ? "" : value).split(/\s*(?:\||,)\s*/)
      .map(norm).filter(Boolean);
    return [...new Set(parts)];
  }
  function valueSimilarity(a, b) {
    const aa = attrParts(a), bb = attrParts(b);
    if (!aa.length || !bb.length) return UNKNOWN_SCORE;
    const bs = new Set(bb);
    let shared = 0;
    for (const x of aa) if (bs.has(x)) shared++;
    return (2 * shared) / (aa.length + bb.length);
  }
  function quantityOf(component) {
    const raw = component && component.a ? component.a.quantity : "";
    const match = String(raw || "").match(/\d+(?:\.\d+)?/);
    return match ? Math.max(1, Number(match[0])) : 1;
  }
  function scoredAttrs(component) {
    const out = {};
    for (const [key, value] of Object.entries((component && component.a) || {})) {
      if (key === "quantity" || IGNORED_ATTR.test(key)
          || key.toLowerCase().includes("vendor_quote") || !norm(value)) continue;
      out[key] = value;
    }
    return out;
  }
  function componentSimilarity(a, b) {
    if (!a || !b) return { score: 0, coverage: 0, differences: ["component missing"] };
    const sameName = norm(a.n) === norm(b.n);
    const aSlot = componentSlot(a), bSlot = componentSlot(b);
    const nameScore = sameName ? 1 : (aSlot && bSlot && aSlot.id === bSlot.id ? 0.5 : 0);
    const qa = quantityOf(a), qb = quantityOf(b);
    const quantityScore = Math.min(qa, qb) / Math.max(qa, qb);
    const aa = scoredAttrs(a), bb = scoredAttrs(b);
    const keys = [...new Set([...Object.keys(aa), ...Object.keys(bb)])].sort();
    let attrPoints = 0, attrKnown = 0;
    const differences = [];
    for (const key of keys) {
      if (key in aa && key in bb) {
        const sim = valueSimilarity(aa[key], bb[key]);
        attrPoints += sim; attrKnown++;
        if (sim < 1) differences.push(key.replace(/_/g, " "));
      } else {
        attrPoints += UNKNOWN_SCORE;
        differences.push(key.replace(/_/g, " ") + " unavailable");
      }
    }
    const attrScore = keys.length ? attrPoints / keys.length : (sameName ? 1 : UNKNOWN_SCORE);
    const attrCoverage = keys.length ? attrKnown / keys.length : 1;
    return {
      score: clamp01(0.30 * nameScore + 0.20 * quantityScore + 0.50 * attrScore),
      coverage: clamp01(0.50 + 0.50 * attrCoverage),
      differences,
    };
  }
  function componentKey(c) {
    return norm(c.n) + "|" + JSON.stringify(scoredAttrs(c));
  }
  function collectionSimilarity(first, second) {
    if (!first.length && !second.length)
      return { score: 1, coverage: 1, differences: [] };
    if (!first.length || !second.length)
      return { score: 0, coverage: 1, differences: ["component present on one order only"] };
    let left = [...first].sort((a, b) => componentKey(a).localeCompare(componentKey(b)));
    let right = [...second].sort((a, b) => componentKey(a).localeCompare(componentKey(b)));
    const leftSig = left.map(componentKey).join("\n"), rightSig = right.map(componentKey).join("\n");
    if (left.length > right.length || (left.length === right.length && leftSig > rightSig))
      [left, right] = [right, left];
    const used = new Set();
    let points = 0, coverage = 0;
    const differences = [];
    for (const c of left) {
      let best = null, bestIndex = -1;
      for (let i = 0; i < right.length; i++) {
        if (used.has(i)) continue;
        const got = componentSimilarity(c, right[i]);
        if (!best || got.score > best.score) { best = got; bestIndex = i; }
      }
      if (best) {
        used.add(bestIndex); points += best.score; coverage += best.coverage;
        differences.push(...best.differences);
      }
    }
    const denominator = Math.max(first.length, second.length);
    return {
      score: clamp01(points / denominator),
      coverage: clamp01(coverage / denominator),
      differences,
    };
  }
  function orderSimilarity(a, b) {
    const points = { core: 0, construction: 0, motor: 0, accessories: 0 };
    const known = { core: 0, construction: 0, motor: 0, accessories: 0 };
    const differences = [];
    let sharedEvidence = 0;
    const pa = orderProfile(a), pb = orderProfile(b);
    const sa = pa.specs, sb = pb.specs;
    const designA = norm(sa.Design), designB = norm(sb.Design);
    for (const field of CORE_FIELDS) {
      const av = norm(sa[field.label]), bv = norm(sb[field.label]);
      if (field.label === "Size" && av && bv && av !== bv
          && designA && designB && designA !== designB) {
        points.core += field.weight * UNKNOWN_SCORE;
        if (av || bv) differences.push("Size not compared across different designs");
        continue;
      }
      if (av && bv) {
        known.core += field.weight;
        if (av === bv) { points.core += field.weight; sharedEvidence++; }
        else differences.push(field.label + ": " + av + " vs " + bv);
      } else {
        points.core += field.weight * UNKNOWN_SCORE;
        if (av || bv) differences.push(field.label + " unavailable on one order");
      }
    }

    const ca = pa.slots, cb = pb.slots;
    const aObserved = pa.observed, bObserved = pb.observed;
    for (const slot of COMPONENT_SLOTS) {
      if (slot.weight === 0) continue;
      const ac = ca[slot.id], bc = cb[slot.id];
      let result;
      if (!ac.length && !bc.length) {
        result = aObserved && bObserved
          ? { score: 1, coverage: 1, differences: [] }
          : { score: UNKNOWN_SCORE, coverage: 0, differences: [] };
      } else if (!ac.length || !bc.length) {
        result = { score: 0, coverage: aObserved && bObserved ? 1 : 0.5,
                   differences: [slot.id.replace(/-/g, " ") + " present on one order only"] };
      } else {
        result = collectionSimilarity(ac, bc);
        const candidateNames = new Set(bc.map(c => norm(c.n)));
        if (slot.weight > 0 && ac.some(c => candidateNames.has(norm(c.n))))
          sharedEvidence++;
      }
      points[slot.group] += slot.weight * result.score;
      known[slot.group] += slot.weight * result.coverage;
      differences.push(...result.differences);
    }

    const rawScore = clamp01(Object.values(points).reduce((x, y) => x + y, 0));
    const coverage = clamp01(Object.values(known).reduce((x, y) => x + y, 0));
    /* Coverage is evidence, not construction.  It can only reduce confidence
     * in a raw match; full evidence leaves the semantic score unchanged. */
    const score = clamp01(rawScore * (0.50 + 0.50 * coverage));
    const groups = {};
    for (const group of Object.keys(GROUP_TOTALS))
      groups[group] = clamp01(points[group] / GROUP_TOTALS[group]);
    return {
      score,
      rawScore,
      coverage,
      groups,
      sharedEvidence,
      differences: [...new Set(differences)].slice(0, 6),
    };
  }
  function attrMatches(have, wanted) {
    const h = attrParts(have), w = norm(wanted);
    return !!w && h.includes(w);
  }
  function componentHasRequired(component, pins) {
    return [...(pins || [])].every(pin => {
      const ix = String(pin).indexOf("=");
      if (ix < 1) return false;
      const key = String(pin).slice(0, ix), wanted = String(pin).slice(ix + 1);
      return component && component.a && key in component.a
        && attrMatches(component.a[key], wanted);
    });
  }
  function focusedSimilarity(whole, targetComponent, candidate, pins) {
    const eligible = componentsNamed(candidate, targetComponent.n)
      .filter(c => componentHasRequired(c, pins));
    if (!eligible.length) return null;
    let best = null, bestComponent = null;
    for (const c of eligible) {
      const got = componentSimilarity(targetComponent, c);
      if (!best || got.score > best.score) { best = got; bestComponent = c; }
    }
    const componentWeight = 1 - FOCUS_WHOLE_WEIGHT;
    return {
      score: clamp01(FOCUS_WHOLE_WEIGHT * whole.score + componentWeight * best.score),
      coverage: clamp01(FOCUS_WHOLE_WEIGHT * whole.coverage
        + componentWeight * best.coverage),
      componentScore: best.score,
      componentCoverage: best.coverage,
      candidateComponent: bestComponent,
      differences: best.differences,
    };
  }

  return {
    CORE_FIELDS,
    COMPONENT_SLOTS,
    GROUP_TOTALS,
    FOCUS_WHOLE_WEIGHT,
    norm,
    specMap,
    flattenComponents,
    componentSlot,
    componentsBySlot,
    componentsNamed,
    scoredAttrs,
    valueSimilarity,
    componentSimilarity,
    componentHasRequired,
    orderSimilarity,
    focusedSimilarity,
  };
}));
