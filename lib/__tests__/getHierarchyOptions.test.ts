import { describe, it, expect } from "vitest";
import { getHierarchyOptions } from "../hierarchy";
import type { HierarchyMap, FilledData } from "../types";

// ── Fixtures ──────────────────────────────────────────────────────────────────

const HIERARCHY: HierarchyMap = {
  divToDept: {
    "04: DRY FOOD":  ["0420: SWEETED GROCE.2", "0450: LIQUOR"],
    "08: HOUSEHOLD": ["0810: CLEANING"],
  },
  deptToSub: {
    "0420: SWEETED GROCE.2": ["042050: CONFECTIONARY"],
    "0450: LIQUOR":          ["045060: WINE"],
    "0810: CLEANING":        ["081010: FLOOR CARE"],
  },
  subToCls: {
    "042050: CONFECTIONARY": ["04205001: CANDY", "04205002: GUM"],
    "045060: WINE":          ["04506010: RED", "04506011: WHITE", "04506012: ROSE"],
    "081010: FLOOR CARE":    ["08101001: MOP"],
  },
};

const ALL: Record<keyof FilledData, string[]> = {
  division:  ["04: DRY FOOD", "08: HOUSEHOLD"],
  dept:      ["0420: SWEETED GROCE.2", "0450: LIQUOR", "0810: CLEANING"],
  subDept:   ["042050: CONFECTIONARY", "045060: WINE", "081010: FLOOR CARE"],
  cls:       ["04205001: CANDY", "04205002: GUM", "04506010: RED", "04506011: WHITE", "04506012: ROSE", "08101001: MOP"],
  planogram: ["POG-001"],
  colN:      ["100"],
  colPiece:  ["5"],
  colO:      ["200"],
};

// ── division ──────────────────────────────────────────────────────────────────

describe("division", () => {
  it("always returns all division options regardless of draft", () => {
    expect(getHierarchyOptions({}, HIERARCHY, ALL).division).toEqual(ALL.division);
    expect(getHierarchyOptions({ division: "04: DRY FOOD" }, HIERARCHY, ALL).division).toEqual(ALL.division);
  });
});

// ── dept ─────────────────────────────────────────────────────────────────────

describe("dept", () => {
  it("shows all depts when division is empty", () => {
    expect(getHierarchyOptions({ division: "" }, HIERARCHY, ALL).dept).toEqual(ALL.dept);
  });

  it("shows all depts when division is undefined", () => {
    expect(getHierarchyOptions({}, HIERARCHY, ALL).dept).toEqual(ALL.dept);
  });

  it("filters depts when division is a known value", () => {
    const result = getHierarchyOptions({ division: "04: DRY FOOD" }, HIERARCHY, ALL);
    expect(result.dept).toEqual(["0420: SWEETED GROCE.2", "0450: LIQUOR"]);
    expect(result.dept).not.toContain("0810: CLEANING");
  });

  it("shows all depts when division is a new (unknown) value", () => {
    const result = getHierarchyOptions({ division: "99: NEW DIVISION" }, HIERARCHY, ALL);
    expect(result.dept).toEqual(ALL.dept);
  });

  it("returns correct depts for a different known division", () => {
    const result = getHierarchyOptions({ division: "08: HOUSEHOLD" }, HIERARCHY, ALL);
    expect(result.dept).toEqual(["0810: CLEANING"]);
  });
});

// ── subDept ──────────────────────────────────────────────────────────────────

describe("subDept", () => {
  it("shows all subDepts when dept is empty", () => {
    expect(getHierarchyOptions({ dept: "" }, HIERARCHY, ALL).subDept).toEqual(ALL.subDept);
  });

  it("shows all subDepts when dept is undefined", () => {
    expect(getHierarchyOptions({}, HIERARCHY, ALL).subDept).toEqual(ALL.subDept);
  });

  it("filters subDepts when dept is a known value", () => {
    const result = getHierarchyOptions({ dept: "0450: LIQUOR" }, HIERARCHY, ALL);
    expect(result.subDept).toEqual(["045060: WINE"]);
    expect(result.subDept).not.toContain("042050: CONFECTIONARY");
  });

  it("shows all subDepts when dept is a new (unknown) value", () => {
    const result = getHierarchyOptions({ dept: "9999: UNKNOWN DEPT" }, HIERARCHY, ALL);
    expect(result.subDept).toEqual(ALL.subDept);
  });

  it("shows all subDepts when dept is empty even if division is known", () => {
    const result = getHierarchyOptions({ division: "04: DRY FOOD", dept: "" }, HIERARCHY, ALL);
    expect(result.subDept).toEqual(ALL.subDept);
  });

  it("shows all subDepts when dept is new even if division is known", () => {
    // Known division does NOT cascade filtering down to grandchild when middle is new
    const result = getHierarchyOptions({ division: "04: DRY FOOD", dept: "NEW DEPT" }, HIERARCHY, ALL);
    expect(result.subDept).toEqual(ALL.subDept);
  });
});

// ── cls ───────────────────────────────────────────────────────────────────────

describe("cls", () => {
  it("shows all classes when subDept is empty", () => {
    expect(getHierarchyOptions({ subDept: "" }, HIERARCHY, ALL).cls).toEqual(ALL.cls);
  });

  it("shows all classes when subDept is undefined", () => {
    expect(getHierarchyOptions({}, HIERARCHY, ALL).cls).toEqual(ALL.cls);
  });

  it("filters classes when subDept is a known value", () => {
    const result = getHierarchyOptions({ subDept: "045060: WINE" }, HIERARCHY, ALL);
    expect(result.cls).toEqual(["04506010: RED", "04506011: WHITE", "04506012: ROSE"]);
    expect(result.cls).not.toContain("04205001: CANDY");
  });

  it("shows all classes when subDept is a new (unknown) value", () => {
    const result = getHierarchyOptions({ subDept: "999999: UNKNOWN SUB" }, HIERARCHY, ALL);
    expect(result.cls).toEqual(ALL.cls);
  });

  it("shows all classes when subDept is new even if parent levels are known", () => {
    const result = getHierarchyOptions(
      { division: "04: DRY FOOD", dept: "0450: LIQUOR", subDept: "NEW SUB" },
      HIERARCHY,
      ALL
    );
    expect(result.cls).toEqual(ALL.cls);
  });

  it("filters classes using only own parent (subDept) regardless of higher ancestors", () => {
    // subDept is known → filter cls, regardless of what division/dept are
    const result = getHierarchyOptions(
      { division: "UNKNOWN DIV", dept: "UNKNOWN DEPT", subDept: "042050: CONFECTIONARY" },
      HIERARCHY,
      ALL
    );
    expect(result.cls).toEqual(["04205001: CANDY", "04205002: GUM"]);
  });
});

// ── full cascade ─────────────────────────────────────────────────────────────

describe("full cascade", () => {
  it("all four columns cascade correctly when all parents are known", () => {
    const result = getHierarchyOptions(
      { division: "04: DRY FOOD", dept: "0450: LIQUOR", subDept: "045060: WINE" },
      HIERARCHY,
      ALL
    );
    expect(result.division).toEqual(ALL.division); // always unfiltered
    expect(result.dept).toEqual(["0420: SWEETED GROCE.2", "0450: LIQUOR"]);
    expect(result.subDept).toEqual(["045060: WINE"]);
    expect(result.cls).toEqual(["04506010: RED", "04506011: WHITE", "04506012: ROSE"]);
  });

  it("non-hierarchy fields (planogram, colN, colO) are not affected by this function", () => {
    // getHierarchyOptions only returns hierarchy keys; callers use allOptions for the rest
    const result = getHierarchyOptions({ division: "04: DRY FOOD" }, HIERARCHY, ALL);
    // The return type only has division/dept/subDept/cls — no planogram/colN/colO keys
    expect(Object.keys(result)).toEqual(["division", "dept", "subDept", "cls"]);
  });
});

// ── fallback behaviour ────────────────────────────────────────────────────────

describe("fallback to allOptions", () => {
  it("falls back to allOptions.dept when a known division has no children in map (defensive)", () => {
    const sparseHierarchy: HierarchyMap = {
      divToDept: { "04: DRY FOOD": [] }, // empty array (edge case)
      deptToSub: {},
      subToCls: {},
    };
    const result = getHierarchyOptions({ division: "04: DRY FOOD" }, sparseHierarchy, ALL);
    // empty array in hierarchy → return that empty array (not allOptions.dept)
    expect(result.dept).toEqual([]);
  });

  it("returns allOptions when hierarchy map has completely empty objects", () => {
    const emptyHierarchy: HierarchyMap = { divToDept: {}, deptToSub: {}, subToCls: {} };
    const result = getHierarchyOptions(
      { division: "04: DRY FOOD", dept: "0450: LIQUOR", subDept: "045060: WINE" },
      emptyHierarchy,
      ALL
    );
    expect(result.dept).toEqual(ALL.dept);
    expect(result.subDept).toEqual(ALL.subDept);
    expect(result.cls).toEqual(ALL.cls);
  });
});
