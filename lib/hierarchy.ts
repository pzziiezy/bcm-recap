import type { FilledData, HierarchyMap, ProcessedRow } from "./types";

// ─── Helpers ──────────────────────────────────────────────────────────────────

function unionMaps(
  a: Record<string, string[]>,
  b: Record<string, string[]>
): Record<string, string[]> {
  const result: Record<string, Set<string>> = {};
  for (const [k, vs] of Object.entries(a)) result[k] = new Set(vs);
  for (const [k, vs] of Object.entries(b)) {
    if (!result[k]) result[k] = new Set();
    for (const v of vs) result[k].add(v);
  }
  return Object.fromEntries(
    Object.entries(result).map(([k, s]) => [k, [...s].sort()])
  );
}

// ─── Public API ───────────────────────────────────────────────────────────────

export const HIERARCHY_KEYS = ["division", "dept", "subDept", "cls"] as const;
export type HierarchyKey = (typeof HIERARCHY_KEYS)[number];

export function isHierarchyKey(key: string): key is HierarchyKey {
  return (HIERARCHY_KEYS as readonly string[]).includes(key);
}

/**
 * Returns filtered option lists for the 4 cascade columns (division → dept → subDept → cls).
 *
 * A column is filtered by its parent ONLY when the parent's current draft value is:
 *   - non-empty, AND
 *   - a *known* value (exists as a key in the hierarchy map).
 *
 * If the parent value is empty or new (not in the map), the column falls back to
 * showing all options from `allOptions`, so the user is never stuck.
 */
export function getHierarchyOptions(
  draft: Partial<FilledData>,
  hierarchy: HierarchyMap,
  allOptions: Record<keyof FilledData, string[]>
): Record<HierarchyKey, string[]> {
  const knownDivs  = new Set(Object.keys(hierarchy.divToDept));
  const knownDepts = new Set(Object.keys(hierarchy.deptToSub));
  const knownSubs  = new Set(Object.keys(hierarchy.subToCls));

  const div    = draft.division ?? "";
  const dept   = draft.dept     ?? "";
  const subDpt = draft.subDept  ?? "";

  return {
    division: allOptions.division,

    dept: div && knownDivs.has(div)
      ? (hierarchy.divToDept[div] ?? allOptions.dept)
      : allOptions.dept,

    subDept: dept && knownDepts.has(dept)
      ? (hierarchy.deptToSub[dept] ?? allOptions.subDept)
      : allOptions.subDept,

    cls: subDpt && knownSubs.has(subDpt)
      ? (hierarchy.subToCls[subDpt] ?? allOptions.cls)
      : allOptions.cls,
  };
}

/**
 * Derives a HierarchyMap from processed result rows.
 * Uses each row's effective data (override takes precedence over filled) to
 * build the division→dept→subDept→cls parent-child relationships.
 */
export function buildHierarchyFromRows(rows: ProcessedRow[]): HierarchyMap {
  const divToDept: Record<string, Set<string>> = {};
  const deptToSub: Record<string, Set<string>> = {};
  const subToCls:  Record<string, Set<string>> = {};

  for (const row of rows) {
    const data = row.override
      ? { ...row.filled, ...row.override }
      : row.filled;
    if (!data) continue;

    const { division, dept, subDept, cls } = data;
    if (division && dept)   (divToDept[division] ??= new Set()).add(dept);
    if (dept     && subDept)(deptToSub[dept]     ??= new Set()).add(subDept);
    if (subDept  && cls)    (subToCls[subDept]   ??= new Set()).add(cls);
  }

  const toSorted = (m: Record<string, Set<string>>) =>
    Object.fromEntries(Object.entries(m).map(([k, s]) => [k, [...s].sort()]));

  return {
    divToDept: toSorted(divToDept),
    deptToSub: toSorted(deptToSub),
    subToCls:  toSorted(subToCls),
  };
}

/**
 * Merges two HierarchyMaps by taking the union of children at every level.
 * Values are deduplicated and sorted. Typically called with the RECAP-file
 * hierarchy as `base` and the result-rows hierarchy as `extra`.
 */
export function mergeHierarchies(base: HierarchyMap, extra: HierarchyMap): HierarchyMap {
  return {
    divToDept: unionMaps(base.divToDept, extra.divToDept),
    deptToSub: unionMaps(base.deptToSub, extra.deptToSub),
    subToCls:  unionMaps(base.subToCls,  extra.subToCls),
  };
}
