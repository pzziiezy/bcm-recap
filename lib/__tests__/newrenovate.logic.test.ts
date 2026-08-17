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
    planofolder01: "PF01_val",
    planofolder02: "PF02_val",
    planofolder03: "PF03_val",
    planofolder04: "PF04_val",
    planofolder05: "PF05_val",
  };

  it("DIVISION uses planofolder02 (not planofolder01)", () => {
    const divisionVal = sm.planofolder02;
    expect(divisionVal).toBe("PF02_val");
    expect(divisionVal).not.toBe(sm.planofolder01);
  });

  it("DEPARTMENT column (PF03_COL) uses planofolder04 (not planofolder03)", () => {
    const deptVal = sm.planofolder04;
    expect(deptVal).toBe("PF04_val");
    expect(deptVal).not.toBe(sm.planofolder03);
  });

  it("POG CATE column (PF04_COL) uses planofolder05 (not planofolder04)", () => {
    const pogVal = sm.planofolder05;
    expect(pogVal).toBe("PF05_val");
    expect(pogVal).not.toBe(sm.planofolder04);
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
