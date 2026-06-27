import type { FilledData, HierarchyMap } from "./types";

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
