import type { FilledData, ProcessedRow } from "./types";

export interface DownloadRow {
  rowIndex: number;
  filled: FilledData | null;
  override?: Partial<FilledData>;
}

export function toDownloadRows(rows: ProcessedRow[]): DownloadRow[] {
  return rows.map(({ rowIndex, filled, override }) => ({
    rowIndex,
    filled,
    override,
  }));
}

// ─── Check Space fill plan — passed to download worker for ZIP-patch ──────

export interface FillCell {
  col: number;    // 0-based column index
  value: string;  // all check-space values are strings
}

export interface FillRow {
  rowIndex: number;  // 0-based (Excel row − 1)
  cells: FillCell[];
}

export interface SheetFill {
  sheetName: string;
  rows: FillRow[];
}

/** Serialisable plan passed to the download worker instead of XLSX.write output */
export interface CheckSpaceFillPlan {
  newScmRows: FillRow[];      // rows to INSERT into NEW SCM before F-J patching
  extraSheets: SheetFill[];   // NEW_DELETE_IM and DEL SCM
}
