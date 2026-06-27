export interface MissingRow {
  rowIndex: number;    // 0-based Excel row index
  barcode: string;
  name: string;
}

export interface SubclassInfo {
  subclassCode: string;
  subclassName: string;
  sourceFile: string;
  colDF: string;       // col DF from 100 ช่อง → RECAP col N (index 13)
}

export type ConfidenceLevel = "confirmed" | "inferred" | "not_found";

export interface FilledData {
  division: string;    // F  (index 5)
  dept: string;        // G  (index 6)
  subDept: string;     // H  (index 7)
  cls: string;         // I  (index 8)
  planogram: string;   // J  (index 9)  ← DATA_SPACEMAN col D
  colN: string;        // N  (index 13) ← 100 ช่อง col DF
  colO: string;        // O  (index 14) ← DATA_SPACEMAN col AL
}

export interface ProcessedRow extends MissingRow {
  filled: FilledData | null;
  confidence: ConfidenceLevel;
  note: string;
  subclassCode?: string;
  // allow user override
  override?: Partial<FilledData>;
}

export interface HierarchyNames {
  divFull: string;     // "04 DRY FOOD"
  deptFull: string;    // "20 SWEETED GROCE.2"
  subdeptFull: string; // "60 BISCUITS"
  clsFull: string;     // "01 BISCUITS/ WAFERS"
}
