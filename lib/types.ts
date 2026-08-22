// Type-only imports — erased at compile time, so these do not create a runtime circular
// dependency with processor.ts / download.ts (which both import type declarations from here).
import type { NewScmRowInfo, DelScmRowInfo, AttributeInfo } from "./processor";
import type { CheckSpaceFillPlan } from "./download";

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
// MINOR REPORT SUB-TAB
// Additive only — nothing above this line is read by Minor Report code directly;
// it consumes the wizard's finished output through PipelineSnapshot below.
// ═══════════════════════════════════════════════════════════════════════════════

/** Barcode → set of store codes with a non-empty value in a sheet's store columns. */
export type StoreFlagMap = Map<string, Set<string>>;

/**
 * Frozen snapshot of everything RECAP Auto-Filler (Step 1-6) produced for one run of
 * handleProcess(). Read-only for downstream consumers (Minor Report today, possibly
 * others later) — nobody outside app/page.tsx should ever mutate this.
 */
export interface PipelineSnapshot {
  results: ProcessedRow[];
  checkSpacePlan: CheckSpaceFillPlan | null;
  indexLookup: IndexLookup;
  barcodeMap: Map<string, SubclassInfo>;      // parseXlsbFiles() — 100 ช่อง
  structureMap: Map<string, HierarchyNames>;  // buildStructureLookup() — 100 ช่อง
  byUpc: Map<string, SpacemanRowMeta>;        // parsePlanogramLookup() — DATA_SPACEMAN
  newScmStoreFlags: StoreFlagMap;             // covers pre-existing rows + Check Space additions
  delScmStoreFlags: StoreFlagMap;
  // Discovered while implementing Minor Report: results/checkSpacePlan alone don't cover
  // (a) DEL SCM at all — parseMissingRows() only ever scans NEW SCM, and
  // (b) NEW SCM rows whose col F was already filled before this run (processRows() skips them).
  // These three fill that gap by reading the enriched in-memory workbook directly.
  newScmRowInfo: Map<string, NewScmRowInfo>;  // extractNewScmRowInfo() — NEW SCM, all barcodes
  delScmRowInfo: Map<string, DelScmRowInfo>;  // extractDelScmRowInfo() — DEL SCM, all barcodes
  attributeMap: Map<string, AttributeInfo>;   // extractAttributeMap() — NEW_DELETE_IM, both blocks
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
