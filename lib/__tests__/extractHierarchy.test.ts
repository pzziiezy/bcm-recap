import { describe, it, expect } from "vitest";
import * as XLSX from "xlsx";
import { extractHierarchy } from "../processor";

// ── Helper: build a minimal workbook with a "NEW SCM" sheet ──────────────────

/**
 * Each element in `dataRows` maps column-index → cell-value.
 * Data rows start at Excel row 4 (0-based), matching the parser's `r = 4` start.
 */
function makeWorkbook(dataRows: Record<number, string>[]): XLSX.WorkBook {
  const ws: XLSX.WorkSheet = {};
  let maxRow = 3;

  dataRows.forEach((row, i) => {
    const r = 4 + i;
    maxRow = Math.max(maxRow, r);
    Object.entries(row).forEach(([col, val]) => {
      ws[XLSX.utils.encode_cell({ r, c: Number(col) })] = { t: "s", v: val };
    });
  });

  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: 20 } });

  return { SheetNames: ["NEW SCM"], Sheets: { "NEW SCM": ws } };
}

// col indices: 5=F(div), 6=G(dept), 7=H(subDept), 8=I(cls)
const F = 5, G = 6, H = 7, I = 8;

// ── Tests ─────────────────────────────────────────────────────────────────────

describe("extractHierarchy — missing sheet", () => {
  it("returns empty maps when NEW SCM sheet is absent", () => {
    const wb: XLSX.WorkBook = { SheetNames: [], Sheets: {} };
    expect(extractHierarchy(wb)).toEqual({ divToDept: {}, deptToSub: {}, subToCls: {} });
  });
});

describe("extractHierarchy — divToDept", () => {
  it("maps a division to its dept", () => {
    const wb = makeWorkbook([
      { [F]: "04: DRY FOOD", [G]: "0450: LIQUOR", [H]: "045060: WINE", [I]: "04506010: RED" },
    ]);
    expect(extractHierarchy(wb).divToDept).toEqual({ "04: DRY FOOD": ["0450: LIQUOR"] });
  });

  it("maps one division to multiple depts (sorted)", () => {
    const wb = makeWorkbook([
      { [F]: "04: DRY FOOD", [G]: "0450: LIQUOR" },
      { [F]: "04: DRY FOOD", [G]: "0420: SWEETED GROCE.2" },
    ]);
    const { divToDept } = extractHierarchy(wb);
    expect(divToDept["04: DRY FOOD"]).toEqual(["0420: SWEETED GROCE.2", "0450: LIQUOR"]);
  });

  it("handles multiple divisions independently", () => {
    const wb = makeWorkbook([
      { [F]: "04: DRY FOOD",  [G]: "0450: LIQUOR"   },
      { [F]: "08: HOUSEHOLD", [G]: "0810: CLEANING"  },
    ]);
    const { divToDept } = extractHierarchy(wb);
    expect(divToDept["04: DRY FOOD"]).toEqual(["0450: LIQUOR"]);
    expect(divToDept["08: HOUSEHOLD"]).toEqual(["0810: CLEANING"]);
  });

  it("deduplicates depts within the same division", () => {
    const wb = makeWorkbook([
      { [F]: "04: DRY FOOD", [G]: "0450: LIQUOR" },
      { [F]: "04: DRY FOOD", [G]: "0450: LIQUOR" }, // duplicate
      { [F]: "04: DRY FOOD", [G]: "0450: LIQUOR" }, // duplicate
    ]);
    expect(extractHierarchy(wb).divToDept["04: DRY FOOD"]).toHaveLength(1);
  });

  it("skips row when division cell is empty", () => {
    const wb = makeWorkbook([
      { [F]: "", [G]: "0450: LIQUOR" },
    ]);
    expect(Object.keys(extractHierarchy(wb).divToDept)).toHaveLength(0);
  });

  it("skips row when dept cell is empty", () => {
    const wb = makeWorkbook([
      { [F]: "04: DRY FOOD", [G]: "" },
    ]);
    expect(Object.keys(extractHierarchy(wb).divToDept)).toHaveLength(0);
  });
});

describe("extractHierarchy — deptToSub", () => {
  it("maps a dept to its subDept", () => {
    const wb = makeWorkbook([
      { [F]: "04: DRY FOOD", [G]: "0450: LIQUOR", [H]: "045060: WINE", [I]: "04506010: RED" },
    ]);
    expect(extractHierarchy(wb).deptToSub).toEqual({ "0450: LIQUOR": ["045060: WINE"] });
  });

  it("deduplicates subDepts within the same dept", () => {
    const wb = makeWorkbook([
      { [G]: "0450: LIQUOR", [H]: "045060: WINE" },
      { [G]: "0450: LIQUOR", [H]: "045060: WINE" },
    ]);
    expect(extractHierarchy(wb).deptToSub["0450: LIQUOR"]).toHaveLength(1);
  });

  it("skips row when subDept cell is empty", () => {
    const wb = makeWorkbook([
      { [G]: "0450: LIQUOR", [H]: "" },
    ]);
    expect(Object.keys(extractHierarchy(wb).deptToSub)).toHaveLength(0);
  });

  it("skips row when dept cell is empty even if subDept is present", () => {
    const wb = makeWorkbook([
      { [G]: "", [H]: "045060: WINE" },
    ]);
    expect(Object.keys(extractHierarchy(wb).deptToSub)).toHaveLength(0);
  });
});

describe("extractHierarchy — subToCls", () => {
  it("maps a subDept to its classes (sorted)", () => {
    const wb = makeWorkbook([
      { [H]: "045060: WINE", [I]: "04506012: ROSE"  },
      { [H]: "045060: WINE", [I]: "04506010: RED"   },
      { [H]: "045060: WINE", [I]: "04506011: WHITE" },
    ]);
    expect(extractHierarchy(wb).subToCls["045060: WINE"]).toEqual([
      "04506010: RED",
      "04506011: WHITE",
      "04506012: ROSE",
    ]);
  });

  it("deduplicates classes within the same subDept", () => {
    const wb = makeWorkbook([
      { [H]: "045060: WINE", [I]: "04506010: RED" },
      { [H]: "045060: WINE", [I]: "04506010: RED" },
    ]);
    expect(extractHierarchy(wb).subToCls["045060: WINE"]).toHaveLength(1);
  });

  it("skips row when cls cell is empty", () => {
    const wb = makeWorkbook([
      { [H]: "045060: WINE", [I]: "" },
    ]);
    expect(Object.keys(extractHierarchy(wb).subToCls)).toHaveLength(0);
  });

  it("skips row when subDept is empty even if cls is present", () => {
    const wb = makeWorkbook([
      { [H]: "", [I]: "04506010: RED" },
    ]);
    expect(Object.keys(extractHierarchy(wb).subToCls)).toHaveLength(0);
  });
});

describe("extractHierarchy — full integration", () => {
  it("builds all three maps from a realistic multi-row dataset", () => {
    const wb = makeWorkbook([
      { [F]: "04: DRY FOOD",  [G]: "0450: LIQUOR",           [H]: "045060: WINE",         [I]: "04506010: RED"   },
      { [F]: "04: DRY FOOD",  [G]: "0450: LIQUOR",           [H]: "045060: WINE",         [I]: "04506011: WHITE" },
      { [F]: "04: DRY FOOD",  [G]: "0450: LIQUOR",           [H]: "045060: WINE",         [I]: "04506012: ROSE"  },
      { [F]: "04: DRY FOOD",  [G]: "0420: SWEETED GROCE.2",  [H]: "042050: CONFECTIONARY",[I]: "04205001: CANDY" },
      { [F]: "08: HOUSEHOLD", [G]: "0810: CLEANING",          [H]: "081010: FLOOR CARE",   [I]: "08101001: MOP"   },
    ]);

    const { divToDept, deptToSub, subToCls } = extractHierarchy(wb);

    expect(divToDept).toEqual({
      "04: DRY FOOD":  ["0420: SWEETED GROCE.2", "0450: LIQUOR"],
      "08: HOUSEHOLD": ["0810: CLEANING"],
    });
    expect(deptToSub).toEqual({
      "0450: LIQUOR":          ["045060: WINE"],
      "0420: SWEETED GROCE.2": ["042050: CONFECTIONARY"],
      "0810: CLEANING":        ["081010: FLOOR CARE"],
    });
    expect(subToCls).toEqual({
      "045060: WINE":          ["04506010: RED", "04506011: WHITE", "04506012: ROSE"],
      "042050: CONFECTIONARY": ["04205001: CANDY"],
      "081010: FLOOR CARE":    ["08101001: MOP"],
    });
  });

  it("ignores rows before row 4 (header rows)", () => {
    // Put data in rows 0-3 — should be ignored
    const ws: XLSX.WorkSheet = {};
    ws[XLSX.utils.encode_cell({ r: 0, c: F })] = { t: "s", v: "04: DRY FOOD" };
    ws[XLSX.utils.encode_cell({ r: 0, c: G })] = { t: "s", v: "0450: LIQUOR" };
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 3, c: 20 } });
    const wb: XLSX.WorkBook = { SheetNames: ["NEW SCM"], Sheets: { "NEW SCM": ws } };

    expect(extractHierarchy(wb).divToDept).toEqual({});
  });
});
