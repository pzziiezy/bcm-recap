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
  packSize?: string;   // "LABL size of UOM : ขนาด หรือ น้ำหนักสินค้า" — used by Minor Report
}

export type ConfidenceLevel = "confirmed" | "inferred" | "not_found" | "from_spaceman";

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
  descA: string;       // DESC_A → DIVISION fallback
  descB: string;       // DESC_B → DEPT fallback
  descC: string;
  totalUnits: string;  // TOTAL_UNITS column from DATA_SPACEMAN
  salepack?: string;                 // SALEPACK column — used by Minor Report
  purchaseItemForSalepack?: string;  // PURCHASE_ITEM_FOR_SALEPACK column — used by Minor Report
  // Every DISTINCT PLANOGRAM (col D) value seen across all of this UPC's rows in
  // QRY_Product_by_POG — confirmed with the user that one barcode can legitimately sell
  // on multiple planograms at once, so this must be the full set, not just one row's value.
  planograms?: string[];
}

/** One exception rule in the O% config */
export interface ExceptionConfig {
  id: string;
  category: string;     // exact CATEGORY value, or "ทั้งหมด" (wildcard)
  subcategory: string;  // exact SUBCATEGORY value, or "ทั้งหมด"
  descC: string;        // exact DESC_C value, or "ทั้งหมด"
  percentage: string;   // numeric string without %, e.g. "50"
  status: "active" | "inactive" | "deleted";
  createdAt: string;    // ISO timestamp when rule was first added
  updatedAt: string;    // ISO timestamp of last write to Google Sheets
  deletedAt?: string;   // ISO timestamp when soft-deleted (undefined if not deleted)
}

/** Return value of parsePlanogramLookup */
export interface PlanogramLookupResult {
  byPrefix: Map<string, { planogram: string; colAL: string }>;
  byUpc: Map<string, SpacemanRowMeta>;
  categories: string[];
  subcategories: string[];
  descCList: string[];
}

/** One item row from Check Space.xlsx (Sheet2) */
export interface CheckSpaceItem {
  barcode: string;
  name: string;
  status: string;   // Check Space col C (e.g. "NEW ADD SOME STORE", "DELETE ALL STORE")
  remark: string;   // Check Space col D
  pogs: string[];   // POG names from unpivot of matrix col E+
}

/** Lookup tables built from FILE_INDEX_1.xlsx (INDX_BCM) */
export interface IndexLookup {
  pogToByCode: Map<string, string>;        // POG name → BY_CODE (Attribute Code)
  pogToStores: Map<string, Set<string>>;   // POG name → Set of store codes that have the POG
  storeList: string[];                     // ordered store codes from row 13
}

// ═══════════════════════════════════════════════════════════════════════════════
// MINOR REPORT — now the wizard's only output (RECAP retired, see git history for
// the earlier RECAP-round-trip version if ever needed for reference).
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Everything buildMinorReportSheets() needs — computed directly from the wizard's own
 * upload steps (Check Space, FILE_INDEX, 100 ช่อง, DATA_SPACEMAN). No RECAP file, no
 * intermediate NEW SCM/DEL SCM sheets: Check Space's ticked POGs give "stores newly
 * targeted this round", DATA_SPACEMAN's per-UPC planogram gives "stores already
 * selling this item", and FILE_INDEX resolves either into an actual store set.
 */
export interface MinorReportInput {
  checkSpaceItems: CheckSpaceItem[];
  indexLookup: IndexLookup;
  barcodeMap: Map<string, SubclassInfo>;      // parseXlsbFiles() — 100 ช่อง
  structureMap: Map<string, HierarchyNames>;  // buildStructureLookup() — 100 ช่อง
  byUpc: Map<string, SpacemanRowMeta>;        // parsePlanogramLookup() — DATA_SPACEMAN
  exceptionConfig: ExceptionConfig[];
}

// ─── Minor Report output rows — one shape per output sheet ─────────────────

export interface MinorReportNewItemRow {
  division: string;
  department: string;
  upc: string;
  name: string;
  salepack: string;
  recipe: string;
  packSize: string;
  totalUnits: string;
  purShelfStockPiece: string;          // = filled.colPiece
  pctOrdering: string;                 // = filled.colO
  netCapacity: string;                 // = computeNetCapacity(filled.colO, filled.colPiece)
  attClass: string;
  attCode: string;
  storeNumber: string;
  link: "LINK";
  forecastSalesPerMonthStore: string;  // = filled.colN
  remark: string;                      // Check Space status, pass-through verbatim
}

export interface MinorReportNewNotLinkRow {
  upc: string;
  name: string;
  division: string;
  department: string;
  attClass: string;
  attCode: string;
  storeNumber: string;
  link: "New not link";
  remark: string; // Check Space status, pass-through verbatim
}

export interface MinorReportDeleteItemRow {
  upc: string;
  name: string;
  division: string;
  department: string;
  attClass: string;
  attCode: string;
  storeNumber: string;
  link: "NOT LINK";
  remark: string; // pass-through from DEL SCM verbatim — no normalization
}

export interface MinorReportSheets {
  newItem: MinorReportNewItemRow[];
  newNotLink: MinorReportNewNotLinkRow[];
  deleteItem: MinorReportDeleteItemRow[];
}
