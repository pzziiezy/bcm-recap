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
 * Header-row count per sheet — verified against real screenshots of the actual
 * `TO BE_Minor Report.xlsx` template. NOT uniform across sheets, despite the original
 * context doc describing "4 header rows the same in all 3 sheets":
 *   - Recap_New_item:     5 rows — blank / "Pure" legend / "Mini" legend / col-number / header
 *   - Recap_New_not_link: 3 rows — ONE legend row (no separate Pure/Mini, no leading blank) / col-number / header
 *   - Recap_Delete_item:  4 rows — "Pure" legend / "Mini" legend / col-number / header (no leading blank)
 * Data starts immediately after, so dataStartRowIndex (0-based Excel row) == headerRowCount.
 */
export const MINOR_REPORT_HEADER_ROW_COUNT: Record<string, number> = {
  [MINOR_REPORT_SHEET_NAMES.newItem]: 5,
  [MINOR_REPORT_SHEET_NAMES.newNotLink]: 3,
  [MINOR_REPORT_SHEET_NAMES.deleteItem]: 4,
};

/**
 * Column order for each Minor Report sheet — verified against real screenshots of the
 * template. These are used only to know WHERE to write each field; the header/legend
 * text itself already lives in the uploaded template and is never touched.
 *
 * Recap_New_item: column A is NOT a data column — it only holds the "Pure"/"Mini" row
 * labels for rows 2/3, so real data starts at column B. Column J ("Pure Non POG") is a
 * hidden column in the template and intentionally always left blank (doc §4.1/§6).
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

function toFillRows<T>(rows: T[], getters: Array<(r: T) => string>, dataStartRowIndex: number): FillRow[] {
  return rows.map((row, i) => ({
    rowIndex: dataStartRowIndex + i,
    cells: getters
      .map((get, col) => ({ col, value: get(row) }))
      .filter(c => c.value !== ""),
  }));
}

export interface MinorReportFillPlan {
  [MINOR_REPORT_SHEET_NAMES.newItem]: FillRow[];
  [MINOR_REPORT_SHEET_NAMES.newNotLink]: FillRow[];
  [MINOR_REPORT_SHEET_NAMES.deleteItem]: FillRow[];
}

/** Converts the reshaped Minor Report data into the FillRow[] shape the byte-patch
 *  worker (minorReportDownload.worker.ts) understands — one array per sheet, keyed
 *  by the exact sheet name expected in the uploaded template. */
export function buildMinorReportFillPlan(sheets: MinorReportSheets): MinorReportFillPlan {
  return {
    [MINOR_REPORT_SHEET_NAMES.newItem]: toFillRows(
      sheets.newItem, NEW_ITEM_COL_GETTERS, MINOR_REPORT_HEADER_ROW_COUNT[MINOR_REPORT_SHEET_NAMES.newItem]
    ),
    [MINOR_REPORT_SHEET_NAMES.newNotLink]: toFillRows(
      sheets.newNotLink, NEW_NOT_LINK_COL_GETTERS, MINOR_REPORT_HEADER_ROW_COUNT[MINOR_REPORT_SHEET_NAMES.newNotLink]
    ),
    [MINOR_REPORT_SHEET_NAMES.deleteItem]: toFillRows(
      sheets.deleteItem, DELETE_ITEM_COL_GETTERS, MINOR_REPORT_HEADER_ROW_COUNT[MINOR_REPORT_SHEET_NAMES.deleteItem]
    ),
  };
}
