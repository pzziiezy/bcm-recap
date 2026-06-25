export interface MissingRow {
  rowIndex: number;    // 0-based Excel row index
  barcode: string;
  name: string;
}

export interface SubclassInfo {
  subclassCode: string;
  subclassName: string;
  sourceFile: string;
}

export type ConfidenceLevel = "confirmed" | "inferred" | "not_found";

export interface FilledData {
  division: string;    // F
  dept: string;        // G
  subDept: string;     // H
  cls: string;         // I
  planogram: string;   // J
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
