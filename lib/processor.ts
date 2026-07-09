import * as XLSX from "xlsx";
import type {
  MissingRow,
  SubclassInfo,
  ProcessedRow,
  HierarchyNames,
  HierarchyMap,
  FilledData,
  ExceptionConfig,
  SpacemanRowMeta,
  PlanogramLookupResult,
  CheckSpaceItem,
  IndexLookup,
} from "./types";
import type { FillCell, FillRow } from "./download";

// ─── Helpers ───────────────────────────────────────────────────────────────

/** Find a worksheet by name — tries exact match first, then case-insensitive trim. */
function findWbSheet(wb: XLSX.WorkBook, name: string): XLSX.WorkSheet | null {
  if (wb.Sheets[name]) return wb.Sheets[name];
  const lower = name.trim().toLowerCase();
  const actual = wb.SheetNames.find(n => n.trim().toLowerCase() === lower);
  return actual ? wb.Sheets[actual] : null;
}

/** Resolve the actual sheet name in wb (for passing to the download worker). */
export function resolveSheetName(wb: XLSX.WorkBook, name: string): string {
  if (wb.Sheets[name]) return name;
  const lower = name.trim().toLowerCase();
  return wb.SheetNames.find(n => n.trim().toLowerCase() === lower) ?? name;
}

function cellVal(ws: XLSX.WorkSheet, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell) return "";
  const v = cell.v;
  if (v == null) return "";
  // strip trailing .0 from numbers stored as barcodes
  const s = String(v);
  return s.includes(".") && !isNaN(Number(s)) ? s.split(".")[0] : s;
}

/**
 * Normalize a barcode for reliable comparison across sources (RECAP / 100 ช่อง).
 * Strips control chars (C0/C1), soft-hyphen, zero-width chars, BOM, then trims.
 */
function normalizeBarcode(s: string): string {
  let result = "";
  for (const ch of s) {
    const cp = ch.codePointAt(0) ?? 0;
    const invisible =
      cp <= 0x1f ||
      (cp >= 0x7f && cp <= 0x9f) ||
      cp === 0x00ad ||
      (cp >= 0x200b && cp <= 0x200f) ||
      cp === 0x2028 || cp === 0x2029 ||
      cp === 0xfeff;
    if (!invisible) result += ch;
  }
  return result.trim();
}

/** Build hierarchical RECAP-format codes from the 4 level strings */
function buildRecapCodes(h: HierarchyNames): FilledData {
  const div = h.divFull.trim().split(" ")[0]; // "04"
  const dep = h.deptFull.trim().split(" ")[0]; // "20"
  const sub = h.subdeptFull.trim().split(" ")[0]; // "60"
  const cls = h.clsFull.trim().split(" ")[0]; // "01"

  const divName = h.divFull.trim().slice(div.length).trim();
  const depName = h.deptFull.trim().slice(dep.length).trim();
  const subName = h.subdeptFull.trim().slice(sub.length).trim();
  const clsName = h.clsFull.trim().slice(cls.length).trim();

  return {
    division: `${div}: ${divName}`,
    dept: `${div}${dep}: ${depName}`,
    subDept: `${div}${dep}${sub}: ${subName}`,
    cls: `${div}${dep}${sub}${cls}: ${clsName}`,
    planogram: "", // filled later by parsePlanogramLookup col D
    colN: "",      // filled later by parseXlsbFiles col DF
    colPiece: "",  // filled later by processRows (DATA_SPACEMAN TOTAL_UNITS)
    colO: "",      // filled later by processRows (exception config → percentage%)
  };
}

export interface ParseMissingResult {
  rows: MissingRow[];
  totalScanned: number;   // rows that have a barcode in col D
  alreadyFilled: number;  // barcoded rows where col F was already non-empty
}

// ─── Step 1: Parse RECAP ───────────────────────────────────────────────────

export function parseMissingRows(wb: XLSX.WorkBook): ParseMissingResult {
  const ws = wb.Sheets["NEW SCM"];
  if (!ws) throw new Error('ไม่พบชีต "NEW SCM" ในไฟล์ RECAP');

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const rows: MissingRow[] = [];
  let totalScanned = 0;
  let alreadyFilled = 0;

  for (let r = 4; r <= range.e.r; r++) {
    const barcode = normalizeBarcode(cellVal(ws, r, 3)); // col D
    if (!barcode) continue;
    totalScanned++;
    const fVal = cellVal(ws, r, 5); // col F (DIVISION)
    if (fVal) {
      alreadyFilled++;
    } else {
      rows.push({ rowIndex: r, barcode, name: cellVal(ws, r, 4) });
    }
  }
  return { rows, totalScanned, alreadyFilled };
}

// ─── Step 2: Parse xlsb/xlsx source files (100 ช่อง) ──────────────────────

const COL_DF_HEADER = "MBC Forecast sale / Month / Store (pcs)";

export async function parseXlsbFiles(
  files: File[]
): Promise<Map<string, SubclassInfo>> {
  const map = new Map<string, SubclassInfo>();

  for (const file of files) {
    const buf = await file.arrayBuffer();
    let wb: XLSX.WorkBook;
    try {
      wb = XLSX.read(buf, { type: "array" });
    } catch {
      continue;
    }

    // Try "Base" sheet first, then "Input"
    for (const sheetName of ["Base", "Input"]) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;

      const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
      let barcodeCol = -1;
      let subclassCodeCol = -1;
      let subclassNameCol = -1;
      let colDFCol = -1;

      // Find header row (search first 70 rows, up to 250 cols)
      let headerRow = -1;
      for (let r = 0; r <= Math.min(range.e.r, 70); r++) {
        for (let c = 0; c <= Math.min(range.e.c, 250); c++) {
          const v = cellVal(ws, r, c);
          if (v.includes("Barcode / PLU")) barcodeCol = c;
          if (v.includes("Sub-Class") && v.includes("รหัส")) subclassCodeCol = c;
          if (v === "Sub-Class Name ชื่อโครงสร้างสินค้า") subclassNameCol = c;
          if (v === COL_DF_HEADER) colDFCol = c;
        }
        if (barcodeCol >= 0 && subclassCodeCol >= 0) {
          headerRow = r;
          break;
        }
      }

      if (barcodeCol < 0 || subclassCodeCol < 0) continue;

      // Data rows start right after the header row
      for (let r = headerRow + 1; r <= range.e.r; r++) {
        const rawBc = cellVal(ws, r, barcodeCol);
        const bc = normalizeBarcode(rawBc);
        const code = cellVal(ws, r, subclassCodeCol);
        if (bc && code && bc.length >= 7 && code.length === 10) {
          if (!map.has(bc)) {
            map.set(bc, {
              subclassCode: code,
              subclassName: subclassNameCol >= 0 ? cellVal(ws, r, subclassNameCol) : "",
              sourceFile: file.name,
              colDF: colDFCol >= 0 ? cellVal(ws, r, colDFCol) : "",
            });
          }
        }
      }
    }
  }

  return map;
}

// ─── Step 3: Build ProdStructure lookup (from any xlsb file) ──────────────

export async function buildStructureLookup(
  files: File[]
): Promise<Map<string, HierarchyNames>> {
  const map = new Map<string, HierarchyNames>();

  for (const file of files) {
    const buf = await file.arrayBuffer();
    let wb: XLSX.WorkBook;
    try {
      wb = XLSX.read(buf, { type: "array" });
    } catch {
      continue;
    }

    const ws = wb.Sheets["Sh_ProdStructure"];
    if (!ws) continue;

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

    for (let r = 1; r <= range.e.r; r++) {
      const code = cellVal(ws, r, 14); // SUB_CLASS_CODE (col O)
      if (!code || code.length !== 10) continue;

      const divFull     = cellVal(ws, r, 9);  // col J
      const deptFull    = cellVal(ws, r, 10); // col K
      const subdeptFull = cellVal(ws, r, 11); // col L
      const clsFull     = cellVal(ws, r, 12); // col M

      if (divFull && deptFull && subdeptFull && clsFull && !map.has(code)) {
        map.set(code, { divFull, deptFull, subdeptFull, clsFull });
      }
    }
  }

  return map;
}

// ─── Step 4: Parse DATA_SPACEMAN → prefix/UPC lookups ───────────────────────

export async function parsePlanogramLookup(
  file: File,
  onProgress?: (pct: number) => void
): Promise<PlanogramLookupResult> {
  const empty: PlanogramLookupResult = {
    byPrefix: new Map(),
    byUpc: new Map(),
    categories: [],
    subcategories: [],
    descCList: [],
  };

  const buf = await file.arrayBuffer();
  onProgress?.(20);

  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buf, { type: "array" });
  } catch {
    return empty;
  }
  onProgress?.(60);

  const ws = wb.Sheets["QRY_Product_by_POG"];
  if (!ws) return empty;

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

  // Locate columns by header name
  let subcatCol = -1, categoryCol = -1, upcCol = -1, descACol = -1, descBCol = -1, descCCol = -1, totalUnitsCol = -1;
  for (let c = 0; c <= range.e.c; c++) {
    const h = cellVal(ws, 0, c);
    if (h === "SUBCATEGORY") subcatCol = c;
    else if (h === "CATEGORY") categoryCol = c;
    else if (h === "UPC") upcCol = c;
    else if (h === "DESC_A") descACol = c;
    else if (h === "DESC_B") descBCol = c;
    else if (h === "DESC_C") descCCol = c;
    else if (h === "TOTAL_UNITS") totalUnitsCol = c;
  }
  if (subcatCol < 0) return empty;

  const COL_D = 3; // PLANOGRAM column (col D, 0-indexed)

  const freqPlog = new Map<string, Map<string, number>>();
  const byUpc    = new Map<string, SpacemanRowMeta>();
  const catSet   = new Set<string>();
  const subSet   = new Set<string>();
  const descSet  = new Set<string>();

  const mostFrequent = (freq: Map<string, number>): string => {
    let best = ""; let bestCount = 0;
    for (const [val, count] of freq.entries()) {
      if (count > bestCount) { best = val; bestCount = count; }
    }
    return best;
  };

  for (let r = 1; r <= range.e.r; r++) {
    const subcat = cellVal(ws, r, subcatCol);
    if (!subcat) continue;

    const prefix = subcat.slice(0, 6);
    if (!/^\d{6}$/.test(prefix)) continue;

    const plog = cellVal(ws, r, COL_D);
    if (plog) {
      if (!freqPlog.has(prefix)) freqPlog.set(prefix, new Map());
      const m = freqPlog.get(prefix)!;
      m.set(plog, (m.get(plog) || 0) + 1);
    }

    // Build UPC-level meta for config matching + TOTAL_UNITS lookup
    if (upcCol >= 0) {
      const upc = normalizeBarcode(cellVal(ws, r, upcCol));
      if (upc && !byUpc.has(upc)) {
        const category   = categoryCol   >= 0 ? cellVal(ws, r, categoryCol)   : "";
        const descA      = descACol      >= 0 ? cellVal(ws, r, descACol)      : "";
        const descB      = descBCol      >= 0 ? cellVal(ws, r, descBCol)      : "";
        const descC      = descCCol      >= 0 ? cellVal(ws, r, descCCol)      : "";
        const totalUnits = totalUnitsCol >= 0 ? cellVal(ws, r, totalUnitsCol) : "";
        byUpc.set(upc, { category, subcategory: subcat, descA, descB, descC, totalUnits });
        if (category) catSet.add(category);
        if (subcat)   subSet.add(subcat);
        if (descC)    descSet.add(descC);
      }
    }
  }

  const byPrefix = new Map<string, { planogram: string; colAL: string }>();
  for (const prefix of freqPlog.keys()) {
    byPrefix.set(prefix, {
      planogram: mostFrequent(freqPlog.get(prefix)!),
      colAL: "",
    });
  }

  onProgress?.(100);
  return {
    byPrefix,
    byUpc,
    categories:   [...catSet].sort(),
    subcategories: [...subSet].sort(),
    descCList:    [...descSet].sort(),
  };
}

// ─── Extract cascading hierarchy from RECAP cols F-I ──────────────────────

export function extractHierarchy(wb: XLSX.WorkBook): HierarchyMap {
  const ws = wb.Sheets["NEW SCM"];
  if (!ws) return { divToDept: {}, deptToSub: {}, subToCls: {} };

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

  const divToDept = new Map<string, Set<string>>();
  const deptToSub = new Map<string, Set<string>>();
  const subToCls  = new Map<string, Set<string>>();

  for (let r = 4; r <= range.e.r; r++) {
    const div    = cellVal(ws, r, 5); // col F — DIVISION
    const dept   = cellVal(ws, r, 6); // col G — DEPT
    const subDpt = cellVal(ws, r, 7); // col H — SUB-DEPT
    const cls    = cellVal(ws, r, 8); // col I — Class

    if (div && dept) {
      if (!divToDept.has(div)) divToDept.set(div, new Set());
      divToDept.get(div)!.add(dept);
    }
    if (dept && subDpt) {
      if (!deptToSub.has(dept)) deptToSub.set(dept, new Set());
      deptToSub.get(dept)!.add(subDpt);
    }
    if (subDpt && cls) {
      if (!subToCls.has(subDpt)) subToCls.set(subDpt, new Set());
      subToCls.get(subDpt)!.add(cls);
    }
  }

  const toSortedRecord = (m: Map<string, Set<string>>): Record<string, string[]> =>
    Object.fromEntries([...m.entries()].map(([k, s]) => [k, [...s].sort()]));

  return {
    divToDept: toSortedRecord(divToDept),
    deptToSub: toSortedRecord(deptToSub),
    subToCls:  toSortedRecord(subToCls),
  };
}

// ─── Extract existing non-empty values from RECAP cols F-J, N, O ──────────

export function extractExistingValues(
  wb: XLSX.WorkBook
): Partial<Record<string, string[]>> {
  const ws = wb.Sheets["NEW SCM"];
  if (!ws) return {};

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const colMap: Record<string, number> = {
    division: 5,
    dept: 6,
    subDept: 7,
    cls: 8,
    planogram: 9,
    colN: 13,
    colO: 14, // O: percentage string e.g. "100%"
  };

  const sets: Record<string, Set<string>> = {};
  for (const key of Object.keys(colMap)) sets[key] = new Set();

  for (let r = 4; r <= range.e.r; r++) {
    for (const [key, c] of Object.entries(colMap)) {
      const v = cellVal(ws, r, c);
      if (v) sets[key].add(v);
    }
  }

  return Object.fromEntries(
    Object.entries(sets).map(([k, s]) => [k, [...s].sort()])
  );
}

// ─── Exception config helpers ───────────────────────────────────────────────

function matchesConfig(entry: ExceptionConfig, meta: SpacemanRowMeta): boolean {
  const catOk  = entry.category    === "ทั้งหมด" || entry.category    === meta.category;
  const subOk  = entry.subcategory === "ทั้งหมด" || entry.subcategory === meta.subcategory;
  const descOk = entry.descC       === "ทั้งหมด" || entry.descC       === meta.descC;
  return catOk && subOk && descOk;
}

function findMatchingConfig(config: ExceptionConfig[], meta: SpacemanRowMeta): ExceptionConfig | null {
  for (const entry of config) {
    if (entry.status === "inactive") continue;
    if (matchesConfig(entry, meta)) return entry;
  }
  return null;
}

// ─── Step 5: Combine everything ────────────────────────────────────────────

export function processRows(
  missing: MissingRow[],
  barcodeMap: Map<string, SubclassInfo>,
  structureMap: Map<string, HierarchyNames>,
  planogramResult: PlanogramLookupResult,
  config: ExceptionConfig[] = []
): ProcessedRow[] {
  const { byPrefix, byUpc } = planogramResult;

  return missing.map((row) => {
    const info = barcodeMap.get(row.barcode);

    if (!info) {
      // Fallback: look up barcode in DATA_SPACEMAN to fill F/G/H/I
      const spacemanMeta = byUpc.get(row.barcode);
      if (spacemanMeta) {
        const subdeptPrefix = spacemanMeta.subcategory.slice(0, 6);
        const plogEntry = byPrefix.get(subdeptPrefix);
        const matched = config.length > 0 ? findMatchingConfig(config, spacemanMeta) : null;
        return {
          ...row,
          filled: {
            division: spacemanMeta.descA,
            dept:     spacemanMeta.descB,
            subDept:  spacemanMeta.descC,
            cls:      spacemanMeta.category,
            planogram: plogEntry?.planogram || "",
            colN:     "",   // MBC Forecast: ว่างไว้เมื่อข้อมูลมาจาก DATA_SPACEMAN
            colPiece: spacemanMeta.totalUnits,
            colO:     matched ? `${matched.percentage}%` : "100%",
          },
          confidence: "from_spaceman",
          note: "ไม่พบในไฟล์ 100 ช่อง — F/G/H/I จาก DATA_SPACEMAN",
        };
      }
      return {
        ...row,
        filled: null,
        confidence: "not_found",
        note: "ไม่พบบาร์โค้ดในไฟล์ 100 ช่อง — กรุณากรอกเอง",
      };
    }

    const hierarchy = structureMap.get(info.subclassCode);
    if (!hierarchy) {
      return {
        ...row,
        filled: null,
        confidence: "not_found",
        subclassCode: info.subclassCode,
        note: `พบ Sub-Class Code (${info.subclassCode}) แต่ไม่พบใน Sh_ProdStructure`,
      };
    }

    const filled = buildRecapCodes(hierarchy);
    const subdeptPrefix = info.subclassCode.slice(0, 6);
    const entry = byPrefix.get(subdeptPrefix);
    filled.planogram = entry?.planogram || "";
    filled.colN      = info.colDF || "";

    // UPC-level DATA_SPACEMAN lookup — used for both TOTAL_UNITS and config meta.
    const metaEntry = byUpc.get(row.barcode);

    // Col O (Piece 100%): TOTAL_UNITS from DATA_SPACEMAN.
    // No fallback — this is a physical shelf count; only meaningful when the product
    // is explicitly listed in DATA_SPACEMAN.
    filled.colPiece = metaEntry?.totalUnits || "";

    // Col P (%): Match exception config rules (first active match wins, default 100%).
    // Primary: use DATA_SPACEMAN meta (CATEGORY/SUBCATEGORY/DESC_C) from byUpc.
    // Fallback: derive meta from RECAP hierarchy — Class→category, SubDept→descC,
    // SubclassCode+Name→subcategory — so rules apply even when the barcode is absent
    // from DATA_SPACEMAN's UPC column.
    const derivedMeta: SpacemanRowMeta = {
      category:    filled.cls,
      subcategory: info.subclassCode
        ? (info.subclassName.trim()
            ? `${info.subclassCode}: ${info.subclassName.trim()}`
            : info.subclassCode)
        : "",
      descA:      filled.division,
      descB:      filled.dept,
      descC:      filled.subDept,
      totalUnits: "",
    };
    const meta = metaEntry ?? derivedMeta;
    const matched = config.length > 0 ? findMatchingConfig(config, meta) : null;
    filled.colO = matched ? `${matched.percentage}%` : "100%";

    const confidence = filled.planogram ? "confirmed" : "inferred";

    let configNote = "";
    if (config.length > 0) {
      if (matched) {
        configNote = ` | Rule O: ${matched.percentage}%`;
      } else {
        configNote = ` | ไม่มี Rule ตรง (${meta.category || "–"}) → O=100%`;
      }
    }

    return {
      ...row,
      filled,
      confidence,
      subclassCode: info.subclassCode,
      note: (confidence === "confirmed"
        ? `จาก ${info.sourceFile}`
        : `จาก ${info.sourceFile} — PLANOGRAM ไม่พบใน DATA_SPACEMAN`) + configNote,
    };
  });
}

// ─── Step 6: Write back to workbook and download ──────────────────────────

export function applyAndDownload(
  wb: XLSX.WorkBook,
  rows: ProcessedRow[]
): void {
  const ws = wb.Sheets["NEW SCM"];

  for (const row of rows) {
    const data = row.override
      ? { ...row.filled, ...row.override }
      : row.filled;
    if (!data) continue;

    const write = (c: number, v: string) => {
      const addr = XLSX.utils.encode_cell({ r: row.rowIndex, c });
      ws[addr] = { t: "s", v };
    };
    const writeNum = (c: number, v: number) => {
      const addr = XLSX.utils.encode_cell({ r: row.rowIndex, c });
      ws[addr] = { t: "n", v };
    };

    write(5,  data.division  || "");
    write(6,  data.dept      || "");
    write(7,  data.subDept   || "");
    write(8,  data.cls       || "");
    write(9,  data.planogram || ""); // J ← DATA_SPACEMAN col D (PLANOGRAM)
    write(13, data.colN      || ""); // N ← 100 ช่อง col DF (MBC Forecast)
    write(14, data.colPiece  || ""); // O ← DATA_SPACEMAN TOTAL_UNITS (Piece 100%)
    write(15, data.colO      || ""); // P ← config % (Shelf stock ON POG %)

    // Q ← Net = P% × O pieces
    const pctNum   = parseFloat(data.colO    || "0");
    const pieceNum = parseFloat(data.colPiece || "0");
    if (pctNum > 0 && pieceNum > 0) {
      writeNum(16, Math.round((pctNum / 100) * pieceNum * 100) / 100);
    }
  }

  XLSX.writeFile(wb, "RECAP_filled.xlsx");
}

// ═══════════════════════════════════════════════════════════════════════════════
// CHECK SPACE + FILE_INDEX ENHANCEMENT
// These functions are ADDITIVE — they do NOT modify any logic above.
// ═══════════════════════════════════════════════════════════════════════════════

// ─── Internal helpers ──────────────────────────────────────────────────────

/** Convert 0-based column index → Excel column letters ("A", "Z", "AA" …) */
function colLetter(idx: number): string {
  let s = "";
  let n = idx + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Excel column letters → 0-based index ("A"→0, "Z"→25, "AA"→26) */
function colLetterToIdx(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64;
  return n - 1;
}

/** Write a string cell, skipping empty values */
function writeCell(ws: XLSX.WorkSheet, r: number, c: number, v: string): void {
  if (!v) return;
  ws[XLSX.utils.encode_cell({ r, c })] = { t: "s", v };
}

/** Extend ws["!ref"] to include the given row/col */
function extendRef(ws: XLSX.WorkSheet, maxRow: number, maxCol: number): void {
  if (!ws["!ref"]) {
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });
    return;
  }
  const ref = XLSX.utils.decode_range(ws["!ref"]);
  ref.e.r = Math.max(ref.e.r, maxRow);
  ref.e.c = Math.max(ref.e.c, maxCol);
  ws["!ref"] = XLSX.utils.encode_range(ref);
}

/**
 * Collect unique BY_CODEs for an item (deduped, order of first appearance).
 * Skips POGs not found in the index.
 */
function getUniqueByCodes(item: CheckSpaceItem, idx: IndexLookup): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const pog of item.pogs) {
    const code = idx.pogToByCode.get(pog);
    if (code && !seen.has(code)) { seen.add(code); result.push(code); }
  }
  return result;
}

/**
 * Union of store codes across all POGs for one item.
 */
function getStoreUnion(item: CheckSpaceItem, idx: IndexLookup): Set<string> {
  const set = new Set<string>();
  for (const pog of item.pogs) {
    const pogStores = idx.pogToStores.get(pog);
    if (pogStores) for (const s of pogStores) set.add(s);
  }
  return set;
}

/** Strip everything from the first "_" in a POG name (e.g. "FISH SEAWEED_3B" → "FISH SEAWEED") */
function pogRootName(pog: string): string {
  const idx = pog.indexOf("_");
  return (idx >= 0 ? pog.slice(0, idx) : pog).trim();
}

/**
 * Read store-code→column-index mapping from a sheet's header row.
 * If no store codes found in that row, falls back to building from storeList.
 */
function buildStoreColMap(
  ws: XLSX.WorkSheet,
  headerRowIdx: number,
  storeStartCol: number,
  storeList: string[],
  writeHeaderIfMissing = false,
): Map<string, number> {
  const map = new Map<string, number>();
  const wsRef = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { e: { c: storeStartCol + storeList.length } };
  for (let c = storeStartCol; c <= wsRef.e.c; c++) {
    const v = cellVal(ws, headerRowIdx, c);
    if (v) map.set(v, c);
  }
  if (map.size === 0 && storeList.length > 0) {
    storeList.forEach((code, i) => {
      const col = storeStartCol + i;
      map.set(code, col);
      if (writeHeaderIfMissing) {
        ws[XLSX.utils.encode_cell({ r: headerRowIdx, c: col })] = { t: "s", v: code };
      }
    });
  }
  return map;
}

/**
 * Find the first row index AFTER the last non-empty row in a column.
 * Used to append new data without overwriting existing rows.
 */
function findAppendRow(ws: XLSX.WorkSheet, checkCol: number, dataStartRow: number): number {
  if (!ws["!ref"]) return dataStartRow;
  const ref = XLSX.utils.decode_range(ws["!ref"]);
  // Iterate all existing cell keys (sparse) to find last occupied row in the target column
  const colLet = colLetter(checkCol);
  const re = new RegExp(`^${colLet}(\\d+)$`);
  let lastRow = dataStartRow - 1;
  for (const key of Object.keys(ws)) {
    const m = re.exec(key);
    if (m) {
      const r = parseInt(m[1]) - 1; // convert to 0-indexed
      if (r >= dataStartRow && r > lastRow) lastRow = r;
    }
  }
  return lastRow + 1;
}

/**
 * Count existing data rows (non-empty in checkCol) starting from dataStartRow.
 * Used to continue the sequential numbering (NO.) for new items.
 */
function countExistingRows(ws: XLSX.WorkSheet, checkCol: number, dataStartRow: number): number {
  if (!ws["!ref"]) return 0;
  const ref = XLSX.utils.decode_range(ws["!ref"]);
  let count = 0;
  for (let r = dataStartRow; r <= ref.e.r; r++) {
    if (cellVal(ws, r, checkCol)) count++;
  }
  return count;
}

// ─── Parse Check Space.xlsx ────────────────────────────────────────────────

/**
 * Parse Check Space.xlsx (Sheet2).
 * Header: row 5 (index 4). Data: row 6+ (index 5+).
 * Col A = barcode, B = name, C = status, D = remark, E+ = POG matrix.
 */
export async function parseCheckSpace(file: File): Promise<CheckSpaceItem[]> {
  const buf = await file.arrayBuffer();
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buf, { type: "array", cellFormula: false, cellStyles: false });
  } catch {
    return [];
  }

  const ws = wb.Sheets["Sheet2"];
  if (!ws || !ws["!ref"]) return [];

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const HEADER_ROW = 4; // 0-indexed → Excel row 5
  const DATA_START = 5; // 0-indexed → Excel row 6
  const POG_START_COL = 4; // col E

  // Collect POG names from header row (col E onwards)
  const pogCols: Array<{ col: number; name: string }> = [];
  for (let c = POG_START_COL; c <= range.e.c; c++) {
    const h = cellVal(ws, HEADER_ROW, c);
    if (h) pogCols.push({ col: c, name: h });
  }

  const items: CheckSpaceItem[] = [];
  for (let r = DATA_START; r <= range.e.r; r++) {
    const barcode = normalizeBarcode(cellVal(ws, r, 0));
    if (!barcode) continue;
    const pogs: string[] = [];
    for (const { col, name } of pogCols) {
      if (cellVal(ws, r, col)) pogs.push(name);
    }
    items.push({
      barcode,
      name:   cellVal(ws, r, 1),
      status: cellVal(ws, r, 2),
      remark: cellVal(ws, r, 3),
      pogs,
    });
  }
  return items;
}

// ─── Parse FILE_INDEX_1.xlsx ───────────────────────────────────────────────

/**
 * Parse FILE_INDEX_1.xlsx (sheet INDX_BCM).
 * Row 13 (index 12) = store codes starting col O (index 14).
 * Row 14 (index 13) = headers: col C = POG NAME, col I = BY_CODE.
 * Row 15+ (index 14+) = data.
 *
 * Uses sparse cell iteration to avoid scanning 1,929 store columns per row.
 */
export async function parseFileIndex(file: File): Promise<IndexLookup> {
  const empty: IndexLookup = { pogToByCode: new Map(), pogToStores: new Map(), storeList: [] };
  const buf = await file.arrayBuffer();
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buf, { type: "array", cellFormula: false, cellStyles: false });
  } catch {
    return empty;
  }

  const ws = wb.Sheets["INDX_BCM"];
  if (!ws || !ws["!ref"]) return empty;

  const STORE_ROW    = 12; // 0-indexed → Excel row 13
  const DATA_START   = 14; // 0-indexed → Excel row 15
  const POG_NAME_COL = 2;  // col C
  const BY_CODE_COL  = 8;  // col I
  const STORE_START  = 14; // col O (index 14)

  // Build colIndex→storeCode map from store header row (sparse scan)
  const storeColMap = new Map<number, string>(); // colIdx → storeCode
  const storeList: string[] = [];
  const storeRowRe = new RegExp(`^([A-Z]+)${STORE_ROW + 1}$`);
  for (const key of Object.keys(ws)) {
    const m = storeRowRe.exec(key);
    if (!m) continue;
    const ci = colLetterToIdx(m[1]);
    if (ci < STORE_START) continue;
    const v = ws[key]?.v;
    if (v != null) {
      const sc = String(v).split(".")[0]; // strip .0 from numbers
      if (sc) { storeColMap.set(ci, sc); storeList.push(sc); }
    }
  }
  // storeList order: sort by column index
  storeList.sort((a, b) => {
    const ca = [...storeColMap.entries()].find(([, v]) => v === a)?.[0] ?? 0;
    const cb = [...storeColMap.entries()].find(([, v]) => v === b)?.[0] ?? 0;
    return ca - cb;
  });

  const pogToByCode = new Map<string, string>();
  const pogToStores = new Map<string, Set<string>>();

  // Sparse scan: group cell keys by row
  const rowStoreMap = new Map<number, string[]>(); // rowIdx → [storeCode, ...]
  const dataRe = /^([A-Z]+)(\d+)$/;
  for (const key of Object.keys(ws)) {
    const m = dataRe.exec(key);
    if (!m) continue;
    const ci = colLetterToIdx(m[1]);
    const ri = parseInt(m[2]) - 1; // 0-indexed
    if (ri < DATA_START) continue;

    if (ci === POG_NAME_COL) {
      // handled in the next pass
      continue;
    }
    if (ci === BY_CODE_COL) {
      continue; // handled in next pass
    }
    const storeCode = storeColMap.get(ci);
    if (!storeCode) continue;
    const cell = ws[key];
    if (cell?.v == null) continue;
    if (!rowStoreMap.has(ri)) rowStoreMap.set(ri, []);
    rowStoreMap.get(ri)!.push(storeCode);
  }

  // Second pass: read POG_NAME and BY_CODE columns (dense, these are narrow)
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let r = DATA_START; r <= range.e.r; r++) {
    const pogName = cellVal(ws, r, POG_NAME_COL);
    if (!pogName) continue;

    const byCode = cellVal(ws, r, BY_CODE_COL);
    if (byCode && !pogToByCode.has(pogName)) pogToByCode.set(pogName, byCode);

    if (!pogToStores.has(pogName)) pogToStores.set(pogName, new Set());
    const storeSet = pogToStores.get(pogName)!;
    const stores = rowStoreMap.get(r);
    if (stores) for (const s of stores) storeSet.add(s);
  }

  return { pogToByCode, pogToStores, storeList };
}

// ─── Extra Info from 100 ช่อง ─────────────────────────────────────────────

/**
 * Extract barcode → Extra Info string from xlsb/xlsx files (Sheet "Base").
 * Extra Info = col V (0-indexed col 21) or dynamically located by header.
 * Priority 1 fallback for fillNewDeleteIM / fillDelSCM col O.
 */
export async function buildXlsbExtraInfoMap(files: File[]): Promise<Map<string, string>> {
  const map = new Map<string, string>();

  for (const file of files) {
    const buf = await file.arrayBuffer();
    let wb: XLSX.WorkBook;
    try {
      wb = XLSX.read(buf, { type: "array", cellFormula: false, cellStyles: false });
    } catch {
      continue;
    }

    for (const sheetName of ["Base", "Input"]) {
      const ws = wb.Sheets[sheetName];
      if (!ws || !ws["!ref"]) continue;

      const range = XLSX.utils.decode_range(ws["!ref"]);
      let barcodeCol = -1;
      let extraInfoCol = -1;
      let headerRow = -1;

      for (let r = 0; r <= Math.min(range.e.r, 70); r++) {
        for (let c = 0; c <= Math.min(range.e.c, 250); c++) {
          const v = cellVal(ws, r, c);
          if (v.includes("Barcode / PLU")) barcodeCol = c;
          if (/extra\s*info/i.test(v)) extraInfoCol = c;
        }
        if (barcodeCol >= 0) { headerRow = r; break; }
      }
      if (barcodeCol < 0 || headerRow < 0) continue;
      if (extraInfoCol < 0) extraInfoCol = 21; // fallback: col V

      for (let r = headerRow + 1; r <= range.e.r; r++) {
        const bc = normalizeBarcode(cellVal(ws, r, barcodeCol));
        if (!bc) continue;
        const extra = cellVal(ws, r, extraInfoCol);
        if (extra && !map.has(bc)) map.set(bc, extra);
      }
      break; // found sheet, stop looking for "Input"
    }
  }

  return map;
}

// ─── Fill NEW_DELETE_IM sheet ──────────────────────────────────────────────

/**
 * Fill the NEW_DELETE_IM sheet in wb.
 *
 * Layout: header at Excel row 4 (index 3). Data starts at index 4.
 * New group (left side, cols A-G / indices 0-6):
 *   A=NO, D="MBC1", E=BY_CODE, F=status, G=remark
 *   (B, C empty — they are header labels only)
 * Delete group (right side, cols I-O / indices 8-14):
 *   I=NO, L="MBC1", M=BY_CODE, N=status(col C), O=Extra_Info
 *   (J, K empty)
 * Both groups share row space: totalRows = max(newRows, delRows).
 * Multiple BY_CODEs per item → multiple rows; A/I filled only on first row.
 */
export function fillNewDeleteIM(
  wb: XLSX.WorkBook,
  items: CheckSpaceItem[],
  indexLookup: IndexLookup,
  xlsbExtraInfo: Map<string, string>,
): FillRow[] {
  const ws = findWbSheet(wb, "NEW_DELETE_IM");
  if (!ws) return [];

  const HEADER_ROW = 3; // index → Excel row 4
  const DATA_START = 4; // index → Excel row 5

  // Existing row count (for continuing sequence numbers)
  const existingNew = countExistingRows(ws, 4, DATA_START);  // col E = BY_CODE (New side)
  const existingDel = countExistingRows(ws, 12, DATA_START); // col M = BY_CODE (Del side)
  const appendRow = findAppendRow(ws, 4, DATA_START); // append after existing

  const newItems = items.filter((i) => !i.status.toUpperCase().startsWith("DELETE"));
  const delItems = items.filter((i) => i.status.toUpperCase().startsWith("DELETE"));

  // Expand each group into write rows (one row per unique BY_CODE)
  type ExpandedRow = { seqNo: string; byCode: string; status: string; remark: string; extraInfo: string };
  const expand = (group: CheckSpaceItem[], existingCount: number, isNew: boolean): ExpandedRow[] => {
    const rows: ExpandedRow[] = [];
    let seq = existingCount + 1;
    for (const item of group) {
      const codes = getUniqueByCodes(item, indexLookup);
      if (codes.length === 0) codes.push(""); // include item even if no code found
      const extraInfo = xlsbExtraInfo.get(item.barcode) || (isNew ? "" : item.remark);
      codes.forEach((code, i) => {
        rows.push({
          seqNo:     i === 0 ? String(seq) : "",
          byCode:    code,
          status:    i === 0 ? item.status : "",
          remark:    i === 0 ? item.remark : "",
          extraInfo: i === 0 ? extraInfo : "",
        });
      });
      seq++;
    }
    return rows;
  };

  const newRows = expand(newItems, existingNew, true);
  const delRows = expand(delItems, existingDel, false);
  const totalRows = Math.max(newRows.length, delRows.length);
  if (totalRows === 0) return [];

  const fillRows: FillRow[] = [];

  for (let i = 0; i < totalRows; i++) {
    const nr = newRows[i];
    const dr = delRows[i];
    const r = appendRow + i;
    const cells: FillCell[] = [];
    const push = (col: number, v: string) => { if (v) cells.push({ col, value: v }); };

    if (nr) {
      writeCell(ws, r, 0, nr.seqNo);   push(0,  nr.seqNo);
      writeCell(ws, r, 3, "MBC1");     push(3,  "MBC1");
      writeCell(ws, r, 4, nr.byCode);  push(4,  nr.byCode);
      writeCell(ws, r, 5, nr.status);  push(5,  nr.status);
      writeCell(ws, r, 6, nr.remark);  push(6,  nr.remark);
    }
    if (dr) {
      writeCell(ws, r, 8,  dr.seqNo);      push(8,  dr.seqNo);
      writeCell(ws, r, 11, "MBC1");         push(11, "MBC1");
      writeCell(ws, r, 12, dr.byCode);     push(12, dr.byCode);
      writeCell(ws, r, 13, dr.status);     push(13, dr.status);
      writeCell(ws, r, 14, dr.extraInfo);  push(14, dr.extraInfo);
    }
    if (cells.length > 0) fillRows.push({ rowIndex: r, cells });
  }

  extendRef(ws, appendRow + totalRows - 1, 14);
  return fillRows;
}

// ─── Fill NEW SCM sheet (non-F-J columns only) ────────────────────────────

/**
 * Add new rows to NEW SCM for Check Space "New" items.
 * Only writes A, D, E, K, L, M, and store flag columns (R+).
 * Columns F-J (5-9) and N-Q (13-16) are intentionally LEFT EMPTY —
 * the existing processRows() logic fills them after this function.
 *
 * Header: index 3 (Excel row 4). Data starts: index 4 (Excel row 5).
 * Store flags: col R (index 17) onwards.
 */
export function fillNewSCM(
  wb: XLSX.WorkBook,
  items: CheckSpaceItem[],
  indexLookup: IndexLookup,
): FillRow[] {
  const ws = findWbSheet(wb, "NEW SCM");
  if (!ws) return [];

  const HEADER_ROW = 3;  // index → Excel row 4
  const DATA_START = 4;  // index → Excel row 5
  const BARCODE_COL = 3; // col D
  const STORE_START_COL = 17; // col R

  const newItems = items.filter((i) => !i.status.toUpperCase().startsWith("DELETE"));
  if (newItems.length === 0) return [];

  // Build store → colIdx map from existing header (or from storeList fallback)
  const storeColMap = buildStoreColMap(ws, HEADER_ROW, STORE_START_COL, indexLookup.storeList, true);
  const appendRow = findAppendRow(ws, BARCODE_COL, DATA_START);
  const existingCount = countExistingRows(ws, BARCODE_COL, DATA_START);
  let seq = existingCount + 1;

  let maxCol = STORE_START_COL - 1;
  for (const col of storeColMap.values()) maxCol = Math.max(maxCol, col);

  const fillRows: FillRow[] = [];

  newItems.forEach((item, i) => {
    const r = appendRow + i;
    const cells: FillCell[] = [];
    const push = (col: number, v: string) => { if (v) cells.push({ col, value: v }); };

    writeCell(ws, r, 0,  String(seq));  push(0,  String(seq));
    writeCell(ws, r, 3,  item.barcode); push(3,  item.barcode);
    writeCell(ws, r, 4,  item.name);    push(4,  item.name);
    // F-J (5-9) = empty → processRows() will fill (via download worker applyRows)
    writeCell(ws, r, 10, item.status);  push(10, item.status);
    writeCell(ws, r, 11, item.remark);  push(11, item.remark);
    writeCell(ws, r, 12, "7.2");        push(12, "7.2");
    // N-Q (13-16) = empty → processRows() will fill

    // Store flags
    const storeUnion = getStoreUnion(item, indexLookup);
    for (const [storeCode, col] of storeColMap) {
      if (storeUnion.has(storeCode)) {
        ws[XLSX.utils.encode_cell({ r, c: col })] = { t: "s", v: "1" };
        cells.push({ col, value: "1" });
      }
    }
    if (cells.length > 0) fillRows.push({ rowIndex: r, cells });
    seq++;
  });

  extendRef(ws, appendRow + newItems.length - 1, maxCol);
  return fillRows;
}

// ─── Fill DEL SCM sheet ────────────────────────────────────────────────────

/**
 * Add new rows to DEL SCM for Check Space "Delete" items (status starts with "DELETE").
 *
 * Header: index 4 (Excel row 5). Data starts: index 5 (Excel row 6).
 * Store flags: col O (index 14) onwards.
 *   DELETE ALL STORE → no store flags written (all empty).
 *   DELETE SOME STORE → union of stores from all item POGs.
 *
 * PLANOGRAM = root name of first POG (strip at first "_").
 * DIVISION  = "04: DRY FOOD" (hardcoded per spec).
 * Extra_Info: Priority 1 → xlsb col V; Priority 2 → Check Space col D.
 */
export function fillDelSCM(
  wb: XLSX.WorkBook,
  items: CheckSpaceItem[],
  indexLookup: IndexLookup,
  xlsbExtraInfo: Map<string, string>,
): FillRow[] {
  const ws = findWbSheet(wb, "DEL SCM");
  if (!ws) return [];

  const HEADER_ROW  = 4;  // index → Excel row 5
  const DATA_START  = 5;  // index → Excel row 6
  const BARCODE_COL = 3;  // col D
  const STORE_START_COL = 14; // col O

  const delItems = items.filter((i) => i.status.toUpperCase().startsWith("DELETE"));
  if (delItems.length === 0) return [];

  const storeColMap = buildStoreColMap(ws, HEADER_ROW, STORE_START_COL, indexLookup.storeList, true);
  const appendRow = findAppendRow(ws, BARCODE_COL, DATA_START);
  const existingCount = countExistingRows(ws, BARCODE_COL, DATA_START);
  let seq = existingCount + 1;
  const isDeleteAll = (status: string) => /DELETE\s+ALL/i.test(status);

  let maxCol = STORE_START_COL - 1;
  for (const col of storeColMap.values()) maxCol = Math.max(maxCol, col);

  const fillRows: FillRow[] = [];

  delItems.forEach((item, i) => {
    const r = appendRow + i;
    const extraInfo = xlsbExtraInfo.get(item.barcode) || item.remark;
    const firstPog = item.pogs[0] ?? "";
    const cells: FillCell[] = [];
    const push = (col: number, v: string) => { if (v) cells.push({ col, value: v }); };

    writeCell(ws, r, 0, String(seq));           push(0, String(seq));
    writeCell(ws, r, 3, item.barcode);          push(3, item.barcode);
    writeCell(ws, r, 4, item.name);             push(4, item.name);
    writeCell(ws, r, 5, "04: DRY FOOD");        push(5, "04: DRY FOOD");
    writeCell(ws, r, 6, pogRootName(firstPog)); push(6, pogRootName(firstPog));
    writeCell(ws, r, 7, "7.2");                 push(7, "7.2");
    writeCell(ws, r, 8, item.status);           push(8, item.status);
    writeCell(ws, r, 9, extraInfo);             push(9, extraInfo);
    // K-N (10-13) = week 1-4 → empty (Buyer fills)

    // Store flags
    if (!isDeleteAll(item.status)) {
      const storeUnion = getStoreUnion(item, indexLookup);
      for (const [storeCode, col] of storeColMap) {
        if (storeUnion.has(storeCode)) {
          ws[XLSX.utils.encode_cell({ r, c: col })] = { t: "s", v: "1" };
          cells.push({ col, value: "1" });
        }
      }
    }
    if (cells.length > 0) fillRows.push({ rowIndex: r, cells });
    seq++;
  });

  extendRef(ws, appendRow + delItems.length - 1, maxCol);
  return fillRows;
}
