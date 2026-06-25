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
