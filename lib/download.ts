import * as XLSX from "xlsx";
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

export function applyRowsToSheet(
  ws: XLSX.WorkSheet,
  rows: DownloadRow[]
): void {
  for (const row of rows) {
    const data = row.override ? { ...row.filled, ...row.override } : row.filled;
    if (!data) continue;

    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 5 })] = {
      t: "s",
      v: data.division || "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 6 })] = {
      t: "s",
      v: data.dept || "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 7 })] = {
      t: "s",
      v: data.subDept || "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 8 })] = {
      t: "s",
      v: data.cls || "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 9 })] = {
      t: "s",
      v: data.planogram || "",
    };
  }
}
