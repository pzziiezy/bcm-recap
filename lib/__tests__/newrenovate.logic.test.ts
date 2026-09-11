/**
 * Unit tests for critical logic in lib/newrenovate.worker.ts
 *
 * These tests inline the pure algorithms from the worker so they can be
 * verified without spinning up the full worker environment.  When changing
 * any of the covered functions in the worker, update the corresponding
 * inline copy here to keep them in sync.
 */

import { describe, it, expect } from "vitest";

// ─── Helpers inlined from worker ─────────────────────────────────────────────

/** Per-store classification (mirrors cross-tab INDEX parsing). */
function classifyStores(asIsStores: Set<string>, toBeStores: Set<string>) {
  return {
    existingStores: [...asIsStores].filter(s =>  toBeStores.has(s)),
    newStores:      [...toBeStores].filter(s => !asIsStores.has(s)),
    deleteStores:   [...asIsStores].filter(s => !toBeStores.has(s)),
  };
}

/** Dominant planogram status (mirrors cross-tab INDEX parsing). */
function deriveStatus(
  hasExisting: boolean,
  hasAsIs:     boolean,
  hasToBe:     boolean,
): string {
  if (hasExisting) return "EXISTING";
  if (hasAsIs)     return "DELETE";
  if (hasToBe)     return "NEW EXPAND";
  return "";
}

/** Sheet 2 trigger (mirrors Sheet 2 check in main QRY loop). */
function triggersSheet2(status: string): boolean {
  const up = status.toUpperCase();
  return up === "NEW EXPAND" || up === "NEW"; // "NEW" for flat-table INDEX files
}

/** % Ordering config-rule match (mirrors getOrderingPct in the worker). */
function getOrderingPct(
  cfg: Array<{ category: string; subcategory: string; descC: string; percentage: string; status: string }>,
  category: string,
  subcategory: string,
  descC: string,
): number {
  for (const rule of cfg) {
    if (rule.status === "inactive" || rule.status === "deleted") continue;
    const catOk = rule.category    === "ทั้งหมด" || rule.category    === category;
    const subOk = rule.subcategory === "ทั้งหมด" || rule.subcategory === subcategory;
    const dscOk = rule.descC       === "ทั้งหมด" || rule.descC       === descC;
    if (catOk && subOk && dscOk) return Number(rule.percentage) / 100;
  }
  return 1.0;
}

// ─── 1. Per-store classification ─────────────────────────────────────────────

describe("classifyStores — per-store status from AS IS / TO BE", () => {
  it("all stores EXISTING (both AS IS and TO BE) → existingStores only", () => {
    const { existingStores, newStores, deleteStores } = classifyStores(
      new Set(["1001", "1002"]),
      new Set(["1001", "1002"]),
    );
    expect(existingStores).toEqual(["1001", "1002"]);
    expect(newStores).toEqual([]);
    expect(deleteStores).toEqual([]);
  });

  it("all stores NEW (TO BE only, no AS IS) → newStores only", () => {
    const { existingStores, newStores, deleteStores } = classifyStores(
      new Set([]),
      new Set(["1001", "1002"]),
    );
    expect(existingStores).toEqual([]);
    expect(newStores).toEqual(["1001", "1002"]);
    expect(deleteStores).toEqual([]);
  });

  it("all stores DELETE (AS IS only, no TO BE) → deleteStores only", () => {
    const { existingStores, newStores, deleteStores } = classifyStores(
      new Set(["1001", "1002"]),
      new Set([]),
    );
    expect(existingStores).toEqual([]);
    expect(newStores).toEqual([]);
    expect(deleteStores).toEqual(["1001", "1002"]);
  });

  it("mixed EXISTING + DELETE: 1001 both, 1002 AS IS only", () => {
    const { existingStores, newStores, deleteStores } = classifyStores(
      new Set(["1001", "1002"]),
      new Set(["1001"]),
    );
    expect(existingStores).toEqual(["1001"]);
    expect(newStores).toEqual([]);
    expect(deleteStores).toEqual(["1002"]);
  });

  it("mixed EXISTING + NEW: 1001 both, 1002 TO BE only", () => {
    const { existingStores, newStores, deleteStores } = classifyStores(
      new Set(["1001"]),
      new Set(["1001", "1002"]),
    );
    expect(existingStores).toEqual(["1001"]);
    expect(newStores).toEqual(["1002"]);
    expect(deleteStores).toEqual([]);
  });

  it("all three types in one planogram: 1001=EXISTING, 1002=NEW, 1003=DELETE", () => {
    const { existingStores, newStores, deleteStores } = classifyStores(
      new Set(["1001", "1003"]),
      new Set(["1001", "1002"]),
    );
    expect(existingStores).toEqual(["1001"]);
    expect(newStores).toEqual(["1002"]);
    expect(deleteStores).toEqual(["1003"]);
  });

  it("no stores at all → all empty", () => {
    const { existingStores, newStores, deleteStores } = classifyStores(
      new Set([]),
      new Set([]),
    );
    expect(existingStores).toEqual([]);
    expect(newStores).toEqual([]);
    expect(deleteStores).toEqual([]);
  });
});

// ─── 2. Derived status (dominant) ────────────────────────────────────────────

describe("deriveStatus — dominant planogram status", () => {
  it("hasExisting → EXISTING", () => {
    expect(deriveStatus(true, false, false)).toBe("EXISTING");
  });

  it("EXISTING takes priority over DELETE", () => {
    expect(deriveStatus(true, true, false)).toBe("EXISTING");
  });

  it("EXISTING takes priority over NEW EXPAND", () => {
    expect(deriveStatus(true, false, true)).toBe("EXISTING");
  });

  it("only AS IS stores → DELETE", () => {
    expect(deriveStatus(false, true, false)).toBe("DELETE");
  });

  it("only TO BE stores → NEW EXPAND (not NEW)", () => {
    expect(deriveStatus(false, false, true)).toBe("NEW EXPAND");
  });

  it("no stores → empty string", () => {
    expect(deriveStatus(false, false, false)).toBe("");
  });
});

// ─── 3. Sheet routing ─────────────────────────────────────────────────────────

describe("sheet routing — each store goes to the correct sheet", () => {
  function routeStores(
    existingStores: string[],
    newStores:      string[],
    deleteStores:   string[],
  ) {
    // Sheet 1 = EXISTING + NEW (isDelete logic kept from original worker)
    const sheet1 = [...existingStores, ...newStores];
    const sheet2 = newStores;          // NEW EXPAND only
    const sheet3 = deleteStores;       // DELETE only
    return { sheet1, sheet2, sheet3 };
  }

  it("EXISTING store: Sheet 1 only", () => {
    const { sheet1, sheet2, sheet3 } = routeStores(["1001"], [], []);
    expect(sheet1).toContain("1001");
    expect(sheet2).not.toContain("1001");
    expect(sheet3).not.toContain("1001");
  });

  it("NEW store: Sheet 1 AND Sheet 2", () => {
    const { sheet1, sheet2, sheet3 } = routeStores([], ["1002"], []);
    expect(sheet1).toContain("1002");
    expect(sheet2).toContain("1002");
    expect(sheet3).not.toContain("1002");
  });

  it("DELETE store: Sheet 3 only", () => {
    const { sheet1, sheet2, sheet3 } = routeStores([], [], ["1003"]);
    expect(sheet1).not.toContain("1003");
    expect(sheet2).not.toContain("1003");
    expect(sheet3).toContain("1003");
  });

  it("mixed planogram — each store routed independently", () => {
    // 1001=EXISTING, 1002=NEW, 1003=DELETE
    const { sheet1, sheet2, sheet3 } = routeStores(["1001"], ["1002"], ["1003"]);
    expect(sheet1).toContain("1001");
    expect(sheet1).toContain("1002");
    expect(sheet1).not.toContain("1003");

    expect(sheet2).not.toContain("1001");
    expect(sheet2).toContain("1002");
    expect(sheet2).not.toContain("1003");

    expect(sheet3).not.toContain("1001");
    expect(sheet3).not.toContain("1002");
    expect(sheet3).toContain("1003");
  });

  it("DELETE store does not leak into Sheet 1 even when planogram is EXISTING (mixed)", () => {
    // Critical regression: before the fix, DELETE stores from mixed planograms
    // were silently dropped (not in Sheet 1, not in Sheet 3).
    const { sheet1, sheet3 } = routeStores(["1001"], [], ["1002"]);
    expect(sheet1).not.toContain("1002");  // DELETE must NOT be in Sheet 1
    expect(sheet3).toContain("1002");       // DELETE MUST be in Sheet 3
  });
});

// ─── 4. Planofolder data mapping (Sheet 1) ───────────────────────────────────

describe("planofolder mapping — Sheet 1 column data sources", () => {
  const sm = {
    descA: "DESC_A_val",
    descB: "DESC_B_val",
  };

  it("DIVISION uses DESC_A", () => {
    const divisionVal = sm.descA || "";
    expect(divisionVal).toBe("DESC_A_val");
  });

  it("DEPARTMENT column (PF03_COL) uses DESC_B", () => {
    const deptVal = sm.descB ?? "";
    expect(deptVal).toBe("DESC_B_val");
  });

  it("POG CATE is keyed by planogram name, not by barcode — same barcode on two planograms can get two different POG CATE values", () => {
    const planogramToCate = new Map<string, string>([
      ["POG A", "CATE_A"],
      ["POG B", "CATE_B"],
    ]);
    const pogCateForRowOnPogA = planogramToCate.get("POG A") ?? "";
    const pogCateForRowOnPogB = planogramToCate.get("POG B") ?? "";
    expect(pogCateForRowOnPogA).toBe("CATE_A");
    expect(pogCateForRowOnPogB).toBe("CATE_B");
    expect(pogCateForRowOnPogA).not.toBe(pogCateForRowOnPogB);
  });

  it("POG CATE falls back to blank (not row exclusion) when the planogram isn't in DATA_SPACEMAN", () => {
    const planogramToCate = new Map<string, string>([["POG A", "CATE_A"]]);
    const pogCateForUnknownPog = planogramToCate.get("POG UNKNOWN") ?? "";
    expect(pogCateForUnknownPog).toBe("");
  });
});

// ─── 4b. % Ordering Config Rule matching — must key on CATEGORY/SUBCATEGORY/DESC_C ──

describe("% Ordering — Config Rule matching keys", () => {
  const rule = { category: "04", subcategory: "20", descC: "60", percentage: "40", status: "active" };

  it("matches when barcode's CATEGORY/SUBCATEGORY/DESC_C equal the rule's (not PLANOFOLDER01/03/04)", () => {
    const pct = getOrderingPct([rule], "04", "20", "60");
    expect(pct).toBe(0.4);
  });

  it("regression: PLANOFOLDER01/03/04 values must NOT be compared against the rule — they live in a different value space and would never match a real rule", () => {
    // Before the fix, planofolder01/03/04 (e.g. "04 DRY FOOD" hierarchy text) were passed
    // here instead of category/subcategory/descC, so a rule like the one above practically
    // never matched and % Ordering silently fell back to the 100% default.
    const pctWithWrongKeys = getOrderingPct([rule], "04 DRY FOOD", "20 SWEETED GROCE.2", "60 BISCUITS");
    expect(pctWithWrongKeys).toBe(1.0); // falls through to default — proves these are NOT what a rule matches on
  });

  it("wildcard ทั้งหมด on any of the 3 fields still matches", () => {
    const wildcardRule = { ...rule, category: "ทั้งหมด" };
    expect(getOrderingPct([wildcardRule], "99", "20", "60")).toBe(0.4);
  });

  it("no matching rule falls back to 100%", () => {
    expect(getOrderingPct([rule], "99", "99", "99")).toBe(1.0);
  });

  it("deleted/inactive rules are skipped during matching", () => {
    const deletedRule = { ...rule, status: "deleted" };
    expect(getOrderingPct([deletedRule], "04", "20", "60")).toBe(1.0);
  });
});

// ─── 5. Status rename: NEW → NEW EXPAND ──────────────────────────────────────

describe("status naming — NEW must be stored as NEW EXPAND", () => {
  it("cross-tab INDEX derives NEW EXPAND (not NEW) for TO BE-only planograms", () => {
    expect(deriveStatus(false, false, true)).toBe("NEW EXPAND");
    expect(deriveStatus(false, false, true)).not.toBe("NEW");
  });

  it("Sheet 2 trigger accepts NEW EXPAND", () => {
    expect(triggersSheet2("NEW EXPAND")).toBe(true);
  });

  it("Sheet 2 trigger also accepts NEW (flat-table INDEX backward compat)", () => {
    expect(triggersSheet2("NEW")).toBe(true);
  });

  it("Sheet 2 trigger rejects EXISTING", () => {
    expect(triggersSheet2("EXISTING")).toBe(false);
  });

  it("Sheet 2 trigger rejects DELETE", () => {
    expect(triggersSheet2("DELETE")).toBe(false);
  });
});
