import type {
  MinorReportSheets,
  MinorReportNewItemRow,
  MinorReportNewNotLinkRow,
  MinorReportDeleteItemRow,
} from "./types";
import type { FillRow } from "./download";

export const MINOR_REPORT_SHEET_NAMES = {
  newItem: "Recap_New_item",
  newNotLink: "Recap_New_not_link",
  deleteItem: "Recap_Delete_item",
} as const;

/**
 * Marker cell used to auto-detect each sheet's header row at build time (see
 * xlsxPatch.ts's findHeaderRowNum()). Counting legend/filler rows from a screenshot
 * has proven unreliable — real templates can have extra blank/filter-only rows that
 * aren't obvious from a picture — so the header row is located by content, not by a
 * hardcoded row count, and works regardless of how many rows precede it.
 *
 * Recap_New_item: column A is NOT a data column (holds "Pure"/"Mini" row labels for
 * the legend rows only), so its marker is DIVISION in column B, not column A.
 */
export const MINOR_REPORT_HEADER_MARKER: Record<string, { col: number; text: string }> = {
  [MINOR_REPORT_SHEET_NAMES.newItem]: { col: 1, text: "DIVISION" },
  [MINOR_REPORT_SHEET_NAMES.newNotLink]: { col: 0, text: "UPC" },
  [MINOR_REPORT_SHEET_NAMES.deleteItem]: { col: 0, text: "UPC" },
};

/**
 * Column order for each Minor Report sheet — verified against real screenshots of the
 * template. These are used only to know WHERE to write each field; the header/legend
 * text itself already lives in the uploaded template and is never touched.
 *
 * Recap_New_item: column A is NOT a data column — it only holds the "Pure"/"Mini" row
 * labels for the legend rows, so real data starts at column B. Column J ("Pure Non POG")
 * is a hidden column in the template and intentionally always left blank (doc §4.1/§6).
 */
const NEW_ITEM_COL_GETTERS: Array<(r: MinorReportNewItemRow) => string> = [
  () => "",                             // A — unused (holds the "Pure"/"Mini" row labels only)
  r => r.division,                      // B DIVISION
  r => r.department,                    // C DEPARTMENT
  r => r.upc,                           // D UPC
  r => r.name,                          // E NAME
  r => r.salepack,                      // F SALEPACK
  r => r.recipe,                        // G RECIPE
  r => r.packSize,                      // H PACK SIZE
  r => r.totalUnits,                    // I TOTAL_UNITS
  () => "",                             // J Pure Non POG — hidden column, always blank
  r => r.purShelfStockPiece,            // K PUR Shelf stock ON POG (Piece) First Order (NEWNEW)
  r => r.pctOrdering,                   // L % Ordering
  r => r.netCapacity,                   // M Net Capacity for odering
  r => r.attClass,                      // N ATT_CLASS
  r => r.attCode,                       // O ATT_CODE
  r => r.storeNumber,                   // P STORE NUMBER
  r => r.link,                          // Q LINK
  r => r.forecastSalesPerMonthStore,    // R Forecast Sales/Month/Store
];

const NEW_NOT_LINK_COL_GETTERS: Array<(r: MinorReportNewNotLinkRow) => string> = [
  r => r.upc,          // A UPC
  r => r.name,         // B NAME
  r => r.division,     // C DIVISION
  r => r.department,   // D DEPARTMENT
  r => r.attClass,     // E ATT_CLASS
  r => r.attCode,      // F ATT_CODE
  r => r.storeNumber,  // G STORENUMBER
  r => r.link,         // H LINK
];

const DELETE_ITEM_COL_GETTERS: Array<(r: MinorReportDeleteItemRow) => string> = [
  r => r.upc,          // A UPC
  r => r.name,         // B NAME
  r => r.division,     // C DIVISION
  r => r.department,   // D DEPARTMENT
  r => r.attClass,     // E ATT_CLASS
  r => r.attCode,      // F ATT_CODE
  r => r.storeNumber,  // G STORENUMBER
  r => r.link,         // H LINK
  r => r.remark,       // I REMARK
];

/** rowIndex here is RELATIVE (0, 1, 2, ...) — the real header row is only known once the
 *  template is unzipped in the worker, so absolute row numbers are assigned there via
 *  shiftFillRows(), not here. */
function toFillRows<T>(rows: T[], getters: Array<(r: T) => string>): FillRow[] {
  return rows.map((row, i) => ({
    rowIndex: i,
    cells: getters
      .map((get, col) => ({ col, value: get(row) }))
      .filter(c => c.value !== ""),
  }));
}

/** Shifts relative FillRow.rowIndex values so they land right after the (now known)
 *  detected header row. headerRowNum is 1-based (Excel row number); FillRow.rowIndex is
 *  0-based (Excel row − 1), so the first data row's rowIndex == headerRowNum. */
export function shiftFillRows(rows: FillRow[], headerRowNum: number): FillRow[] {
  return rows.map(r => ({ ...r, rowIndex: r.rowIndex + headerRowNum }));
}

export interface MinorReportFillPlan {
  [MINOR_REPORT_SHEET_NAMES.newItem]: FillRow[];
  [MINOR_REPORT_SHEET_NAMES.newNotLink]: FillRow[];
  [MINOR_REPORT_SHEET_NAMES.deleteItem]: FillRow[];
}

/** Converts the reshaped Minor Report data into the FillRow[] shape the byte-patch
 *  worker (minorReportDownload.worker.ts) understands — one array per sheet, keyed
 *  by the exact sheet name expected in the uploaded template. Row numbers are still
 *  relative at this point; the worker shifts them after detecting each sheet's real
 *  header row (see shiftFillRows()). */
export function buildMinorReportFillPlan(sheets: MinorReportSheets): MinorReportFillPlan {
  return {
    [MINOR_REPORT_SHEET_NAMES.newItem]: toFillRows(sheets.newItem, NEW_ITEM_COL_GETTERS),
    [MINOR_REPORT_SHEET_NAMES.newNotLink]: toFillRows(sheets.newNotLink, NEW_NOT_LINK_COL_GETTERS),
    [MINOR_REPORT_SHEET_NAMES.deleteItem]: toFillRows(sheets.deleteItem, DELETE_ITEM_COL_GETTERS),
  };
}
