import type {
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

/** Marker text used to auto-detect each sheet's header ROW (xlsxPatch.findHeaderRowNum
 *  scans every cell in a row for this, not a fixed column, so it survives reordering too). */
export const MINOR_REPORT_HEADER_MARKER_TEXT: Record<string, string> = {
  [MINOR_REPORT_SHEET_NAMES.newItem]: "UPC",
  [MINOR_REPORT_SHEET_NAMES.newNotLink]: "UPC",
  [MINOR_REPORT_SHEET_NAMES.deleteItem]: "UPC",
};

/**
 * field → expected header text. The actual column for each field is only resolved at
 * build time from the REAL uploaded template (xlsxPatch.mapHeaderColumns), never a
 * hardcoded position — confirmed with the user that template columns get reordered
 * between revisions, so position-based mapping would silently break again otherwise.
 * If a header text below isn't found in the template, that field is just skipped
 * (see buildFillRowsForSheet) rather than guessing a column.
 */
const NEW_ITEM_HEADER_TEXT: Record<keyof MinorReportNewItemRow, string> = {
  upc: "UPC",
  name: "NAME",
  division: "DIVISION",
  department: "DEPARTMENT",
  salepack: "SALEPACK",
  recipe: "RECIPE",
  packSize: "PACK SIZE",
  totalUnits: "TOTAL_UNITS",
  purShelfStockPiece: "PUR Shelf stock ON POG (Piece) First Order (NEWNEW)",
  pctOrdering: "% Ordering",
  netCapacity: "Net Capacity for odering",
  attClass: "ATT_CLASS",
  attCode: "ATT_CODE",
  storeNumber: "STORE NUMBER",
  link: "LINK",
  forecastSalesPerMonthStore: "Forecast Sales/Month/Store",
  remark: "REMARK",
};

const NEW_NOT_LINK_HEADER_TEXT: Record<keyof MinorReportNewNotLinkRow, string> = {
  upc: "UPC",
  name: "NAME",
  division: "DIVISION",
  department: "DEPARTMENT",
  attClass: "ATT_CLASS",
  attCode: "ATT_CODE",
  storeNumber: "STORENUMBER",
  link: "LINK",
  remark: "REMARK",
};

const DELETE_ITEM_HEADER_TEXT: Record<keyof MinorReportDeleteItemRow, string> = {
  upc: "UPC",
  name: "NAME",
  division: "DIVISION",
  department: "DEPARTMENT",
  attClass: "ATT_CLASS",
  attCode: "ATT_CODE",
  storeNumber: "STORENUMBER",
  link: "LINK",
  remark: "REMARK",
};

export const MINOR_REPORT_HEADER_TEXT_MAP: Record<string, Record<string, string>> = {
  [MINOR_REPORT_SHEET_NAMES.newItem]: NEW_ITEM_HEADER_TEXT,
  [MINOR_REPORT_SHEET_NAMES.newNotLink]: NEW_NOT_LINK_HEADER_TEXT,
  [MINOR_REPORT_SHEET_NAMES.deleteItem]: DELETE_ITEM_HEADER_TEXT,
};

/**
 * Builds FillRow[] for one sheet using a header-text → column map discovered from the
 * REAL uploaded template, not a hardcoded position. rowIndex is RELATIVE (0,1,2,...) —
 * shiftFillRows() places it after the real header row once that's known too.
 *
 * A field whose expected header text isn't found in the template is silently dropped
 * from the row data (never guessed into some other column) and reported back in
 * `missingHeaders` so the caller can warn the user — a blank column is a safe failure,
 * writing into the wrong column is not.
 */
export function buildFillRowsForSheet<T>(
  rows: T[],
  headerTextMap: Record<string, string>,
  discoveredColumns: Map<string, number>
): { fillRows: FillRow[]; missingHeaders: string[] } {
  const fieldToCol: Array<[string, number]> = [];
  const missingHeaders: string[] = [];
  for (const [field, headerText] of Object.entries(headerTextMap)) {
    const col = discoveredColumns.get(headerText.trim().toLowerCase());
    if (col === undefined) {
      missingHeaders.push(headerText);
      continue;
    }
    fieldToCol.push([field, col]);
  }

  const fillRows: FillRow[] = rows.map((row, i) => ({
    rowIndex: i,
    cells: fieldToCol
      .map(([field, col]) => ({ col, value: (row as unknown as Record<string, string>)[field] ?? "" }))
      .filter(c => c.value !== ""),
  }));

  return { fillRows, missingHeaders };
}

/** Shifts relative FillRow.rowIndex values so they land right after the (now known)
 *  detected header row. headerRowNum is 1-based (Excel row number); FillRow.rowIndex is
 *  0-based (Excel row − 1), so the first data row's rowIndex == headerRowNum. */
export function shiftFillRows(rows: FillRow[], headerRowNum: number): FillRow[] {
  return rows.map(r => ({ ...r, rowIndex: r.rowIndex + headerRowNum }));
}
