import { describe, it, expect } from "vitest";
import { buildHierarchyFromRows, mergeHierarchies } from "../hierarchy";
import type { ProcessedRow, HierarchyMap } from "../types";

// ── Fixtures ─────────────────────────────────────────────────────────────────

function row(
  division: string,
  dept: string,
  subDept: string,
  cls: string,
  override?: Partial<{ division: string; dept: string; subDept: string; cls: string }>
): ProcessedRow {
  return {
    rowIndex: 0,
    barcode: "000",
    name: "test",
    confidence: "confirmed",
    note: "",
    filled: { division, dept, subDept, cls, planogram: "", colN: "", colO: "" },
    override,
  };
}

const EMPTY: HierarchyMap = { divToDept: {}, deptToSub: {}, subToCls: {} };

// ── buildHierarchyFromRows ────────────────────────────────────────────────────

describe("buildHierarchyFromRows — basic", () => {
  it("returns empty maps for empty row array", () => {
    expect(buildHierarchyFromRows([])).toEqual(EMPTY);
  });

  it("maps division → dept from a single row", () => {
    const result = buildHierarchyFromRows([row("04: DRY FOOD", "0450: LIQUOR", "045060: WINE", "04506010: RED")]);
    expect(result.divToDept).toEqual({ "04: DRY FOOD": ["0450: LIQUOR"] });
  });

  it("maps dept → subDept from a single row", () => {
    const result = buildHierarchyFromRows([row("04: DRY FOOD", "0450: LIQUOR", "045060: WINE", "04506010: RED")]);
    expect(result.deptToSub).toEqual({ "0450: LIQUOR": ["045060: WINE"] });
  });

  it("maps subDept → cls from a single row", () => {
    const result = buildHierarchyFromRows([row("04: DRY FOOD", "0450: LIQUOR", "045060: WINE", "04506010: RED")]);
    expect(result.subToCls).toEqual({ "045060: WINE": ["04506010: RED"] });
  });
});

describe("buildHierarchyFromRows — deduplication + sorting", () => {
  it("deduplicates dept under the same division", () => {
    const rows = [
      row("04: DRY FOOD", "0450: LIQUOR", "", ""),
      row("04: DRY FOOD", "0450: LIQUOR", "", ""),
    ];
    expect(buildHierarchyFromRows(rows).divToDept["04: DRY FOOD"]).toHaveLength(1);
  });

  it("collects multiple unique depts under same division (sorted)", () => {
    const rows = [
      row("04: DRY FOOD", "0450: LIQUOR", "", ""),
      row("04: DRY FOOD", "0420: SWEETED GROCE.2", "", ""),
    ];
    expect(buildHierarchyFromRows(rows).divToDept["04: DRY FOOD"]).toEqual([
      "0420: SWEETED GROCE.2",
      "0450: LIQUOR",
    ]);
  });

  it("collects multiple unique classes under same subDept (sorted)", () => {
    const rows = [
      row("", "", "045060: WINE", "04506012: ROSE"),
      row("", "", "045060: WINE", "04506010: RED"),
    ];
    expect(buildHierarchyFromRows(rows).subToCls["045060: WINE"]).toEqual([
      "04506010: RED",
      "04506012: ROSE",
    ]);
  });
});

describe("buildHierarchyFromRows — empty/missing cells", () => {
  it("skips divToDept when division is empty", () => {
    const rows = [row("", "0450: LIQUOR", "", "")];
    expect(Object.keys(buildHierarchyFromRows(rows).divToDept)).toHaveLength(0);
  });

  it("skips divToDept when dept is empty", () => {
    const rows = [row("04: DRY FOOD", "", "", "")];
    expect(Object.keys(buildHierarchyFromRows(rows).divToDept)).toHaveLength(0);
  });

  it("skips deptToSub when subDept is empty", () => {
    const rows = [row("04: DRY FOOD", "0450: LIQUOR", "", "")];
    expect(Object.keys(buildHierarchyFromRows(rows).deptToSub)).toHaveLength(0);
  });

  it("skips subToCls when cls is empty", () => {
    const rows = [row("04: DRY FOOD", "0450: LIQUOR", "045060: WINE", "")];
    expect(Object.keys(buildHierarchyFromRows(rows).subToCls)).toHaveLength(0);
  });
});

describe("buildHierarchyFromRows — override takes precedence over filled", () => {
  it("uses override value for division when present", () => {
    const r: ProcessedRow = {
      ...row("04: DRY FOOD", "0450: LIQUOR", "045060: WINE", "04506010: RED"),
      override: { division: "08: HOUSEHOLD" },
    };
    const result = buildHierarchyFromRows([r]);
    expect(result.divToDept).toHaveProperty("08: HOUSEHOLD");
    expect(result.divToDept).not.toHaveProperty("04: DRY FOOD");
  });

  it("uses override value for dept when present", () => {
    const r: ProcessedRow = {
      ...row("04: DRY FOOD", "0450: LIQUOR", "045060: WINE", "04506010: RED"),
      override: { dept: "9999: CUSTOM DEPT" },
    };
    const result = buildHierarchyFromRows([r]);
    expect(result.divToDept["04: DRY FOOD"]).toContain("9999: CUSTOM DEPT");
    expect(result.divToDept["04: DRY FOOD"]).not.toContain("0450: LIQUOR");
  });

  it("handles row where filled is null", () => {
    const r: ProcessedRow = {
      rowIndex: 0, barcode: "000", name: "test", confidence: "not_found",
      note: "",
      filled: null,
    };
    expect(buildHierarchyFromRows([r])).toEqual(EMPTY);
  });
});

// ── mergeHierarchies ─────────────────────────────────────────────────────────

describe("mergeHierarchies — basic", () => {
  it("merging two empty maps stays empty", () => {
    expect(mergeHierarchies(EMPTY, EMPTY)).toEqual(EMPTY);
  });

  it("merging base with empty extra returns base unchanged", () => {
    const base: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0410: SWEETED GROCE.1"] },
      deptToSub: {},
      subToCls: {},
    };
    expect(mergeHierarchies(base, EMPTY)).toEqual(base);
  });

  it("merging empty base with extra returns extra", () => {
    const extra: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0450: LIQUOR"] },
      deptToSub: {},
      subToCls: {},
    };
    expect(mergeHierarchies(EMPTY, extra)).toEqual(extra);
  });
});

describe("mergeHierarchies — union under same key", () => {
  it("unions dept lists under same division (sorted, deduplicated)", () => {
    const base: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0410: SWEETED GROCE.1", "0430: SALTED GROCERY"] },
      deptToSub: {},
      subToCls: {},
    };
    const extra: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0420: SWEETED GROCE.2", "0450: LIQUOR"] },
      deptToSub: {},
      subToCls: {},
    };
    const merged = mergeHierarchies(base, extra);
    expect(merged.divToDept["04: DRY FOOD"]).toEqual([
      "0410: SWEETED GROCE.1",
      "0420: SWEETED GROCE.2",
      "0430: SALTED GROCERY",
      "0450: LIQUOR",
    ]);
  });

  it("deduplicates values that exist in both base and extra", () => {
    const base: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0450: LIQUOR"] },
      deptToSub: {},
      subToCls: {},
    };
    const extra: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0450: LIQUOR"] },
      deptToSub: {},
      subToCls: {},
    };
    expect(mergeHierarchies(base, extra).divToDept["04: DRY FOOD"]).toHaveLength(1);
  });

  it("adds entirely new division from extra that base doesn't have", () => {
    const base: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0450: LIQUOR"] },
      deptToSub: {},
      subToCls: {},
    };
    const extra: HierarchyMap = {
      divToDept: { "08: HOUSEHOLD": ["0810: CLEANING"] },
      deptToSub: {},
      subToCls: {},
    };
    const merged = mergeHierarchies(base, extra);
    expect(Object.keys(merged.divToDept)).toEqual(["04: DRY FOOD", "08: HOUSEHOLD"]);
    expect(merged.divToDept["08: HOUSEHOLD"]).toEqual(["0810: CLEANING"]);
  });

  it("unions all three map levels independently", () => {
    const base: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0450: LIQUOR"] },
      deptToSub: { "0450: LIQUOR": ["045060: WINE"] },
      subToCls:  { "045060: WINE": ["04506010: RED"] },
    };
    const extra: HierarchyMap = {
      divToDept: { "04: DRY FOOD": ["0420: SWEETED GROCE.2"] },
      deptToSub: { "0420: SWEETED GROCE.2": ["042050: CONFECTIONARY"] },
      subToCls:  { "042050: CONFECTIONARY": ["04205001: CANDY"] },
    };
    const merged = mergeHierarchies(base, extra);
    expect(merged.divToDept["04: DRY FOOD"]).toEqual(["0420: SWEETED GROCE.2", "0450: LIQUOR"]);
    expect(merged.deptToSub).toHaveProperty("0420: SWEETED GROCE.2");
    expect(merged.subToCls).toHaveProperty("042050: CONFECTIONARY");
  });
});

describe("mergeHierarchies — the original bug scenario", () => {
  it("fills in missing children that only exist in result rows, not in RECAP", () => {
    // RECAP hierarchy: "04: DRY FOOD" has only existing-in-file depts
    const recapHierarchy: HierarchyMap = {
      divToDept: {
        "04: DRY FOOD": ["0410: SWEETED GROCE.1", "0430: SALTED GROCERY", "0440: BEVERAGE"],
      },
      deptToSub: {},
      subToCls: {},
    };

    // Result rows: processor matched new depts not previously in RECAP
    const resultRows = [
      row("04: DRY FOOD", "0450: LIQUOR", "045060: WINE", "04506010: RED"),
      row("04: DRY FOOD", "0420: SWEETED GROCE.2", "042050: CONFECTIONARY", "04205001: CANDY"),
    ];
    const rowsHierarchy = buildHierarchyFromRows(resultRows);
    const merged = mergeHierarchies(recapHierarchy, rowsHierarchy);

    // After merge, G dropdown for "04: DRY FOOD" must include BOTH old and new depts
    expect(merged.divToDept["04: DRY FOOD"]).toContain("0410: SWEETED GROCE.1");
    expect(merged.divToDept["04: DRY FOOD"]).toContain("0420: SWEETED GROCE.2");
    expect(merged.divToDept["04: DRY FOOD"]).toContain("0430: SALTED GROCERY");
    expect(merged.divToDept["04: DRY FOOD"]).toContain("0440: BEVERAGE");
    expect(merged.divToDept["04: DRY FOOD"]).toContain("0450: LIQUOR");
  });
});
