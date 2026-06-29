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
  planogram: string;   // J  (index 9)  ← DATA_SPACEMAN col D (PLANOGRAM)
  colN: string;        // N  (index 13) ← 100 ช่อง col DF (MBC Forecast)
  colPiece: string;    // O  (index 14) ← DATA_SPACEMAN TOTAL_UNITS (Shelf stock ON POG Piece 100%)
  colO: string;        // P  (index 15) ← config % string e.g. "100%" or "40%"
  // Q (index 16) = colO% × colPiece — computed on write, not stored
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

/** Cascading parent→children relationships extracted from an existing RECAP file */
export interface HierarchyMap {
  divToDept: Record<string, string[]>; // DIVISION value → sorted DEPT values
  deptToSub: Record<string, string[]>; // DEPT value → sorted SUB-DEPT values
  subToCls:  Record<string, string[]>; // SUB-DEPT value → sorted CLASS values
}

/** Metadata for a single product looked up from DATA_SPACEMAN */
export interface SpacemanRowMeta {
  category: string;
  subcategory: string;
  descC: string;
  totalUnits: string;  // TOTAL_UNITS column from DATA_SPACEMAN
}

/** One exception rule in the O% config */
export interface ExceptionConfig {
  id: string;
  category: string;     // exact CATEGORY value, or "ทั้งหมด" (wildcard)
  subcategory: string;  // exact SUBCATEGORY value, or "ทั้งหมด"
  descC: string;        // exact DESC_C value, or "ทั้งหมด"
  percentage: string;   // numeric string without %, e.g. "50"
  status: "active" | "inactive";
  createdAt: string;    // ISO timestamp when rule was first added
  updatedAt: string;    // ISO timestamp of last write to Google Sheets
}

/** Return value of parsePlanogramLookup */
export interface PlanogramLookupResult {
  byPrefix: Map<string, { planogram: string; colAL: string }>;
  byUpc: Map<string, SpacemanRowMeta>;
  categories: string[];
  subcategories: string[];
  descCList: string[];
}
