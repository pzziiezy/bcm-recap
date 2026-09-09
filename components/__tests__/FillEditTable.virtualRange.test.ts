import { describe, it, expect } from "vitest";
import { computeVirtualRange } from "../FillEditTable";

describe("computeVirtualRange", () => {
  it("renders everything when viewport/rowHeight aren't known yet (safe fallback)", () => {
    expect(computeVirtualRange(0, 0, 29, 4236, 10)).toEqual({ start: 0, end: 4236 });
    expect(computeVirtualRange(0, 500, 0, 4236, 10)).toEqual({ start: 0, end: 4236 });
  });

  it("returns an empty range for an empty list", () => {
    expect(computeVirtualRange(0, 500, 29, 0, 10)).toEqual({ start: 0, end: 0 });
  });

  it("windows around the top of the list with overscan, never going negative", () => {
    // viewport fits ~17 rows (500/29), scrolled to the very top
    const { start, end } = computeVirtualRange(0, 500, 29, 4236, 10);
    expect(start).toBe(0); // overscan can't push it below 0
    expect(end).toBeGreaterThan(17);
    expect(end).toBeLessThan(40);
  });

  it("windows around the middle of the list, offset by scrollTop", () => {
    const scrollTop = 29 * 1000; // scrolled to row 1000
    const { start, end } = computeVirtualRange(scrollTop, 500, 29, 4236, 10);
    expect(start).toBe(1000 - 10);
    expect(end).toBeGreaterThan(1000);
    expect(end).toBeLessThan(1040);
  });

  it("clamps the end of the range to totalRows at the very bottom of the list", () => {
    const scrollTop = 29 * 4236; // scrolled all the way to (or past) the last row
    const { start, end } = computeVirtualRange(scrollTop, 500, 29, 4236, 10);
    expect(end).toBe(4236); // never overshoots the real row count
    expect(start).toBeLessThan(4236);
  });

  it("below the virtualize threshold the caller just uses the full range directly (not this function)", () => {
    // sanity: a tiny list still behaves correctly if ever passed through
    expect(computeVirtualRange(0, 500, 29, 5, 10)).toEqual({ start: 0, end: 5 });
  });
});
