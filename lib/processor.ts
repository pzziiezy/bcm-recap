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
} from "./types";

// ─── Helpers ───────────────────────────────────────────────────────────────

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
    colO: "",      // filled later by processRows (exception config → percentage%)
  };
}

// ─── Step 1: Parse RECAP ───────────────────────────────────────────────────

export function parseMissingRows(wb: XLSX.WorkBook): MissingRow[] {
  const ws = wb.Sheets["NEW SCM"];
  if (!ws) throw new Error('ไม่พบชีต "NEW SCM" ในไฟล์ RECAP');

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const results: MissingRow[] = [];

  for (let r = 4; r <= range.e.r; r++) {
    const barcode = normalizeBarcode(cellVal(ws, r, 3)); // col D
    const fVal = cellVal(ws, r, 5); // col F (DIVISION)
    if (barcode && !fVal) {
      results.push({
        rowIndex: r,
        barcode,
        name: cellVal(ws, r, 4), // col E
      });
    }
  }
  return results;
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
  let subcatCol = -1, categoryCol = -1, upcCol = -1, descCCol = -1;
  for (let c = 0; c <= range.e.c; c++) {
    const h = cellVal(ws, 0, c);
    if (h === "SUBCATEGORY") subcatCol = c;
    else if (h === "CATEGORY") categoryCol = c;
    else if (h === "UPC") upcCol = c;
    else if (h === "DESC_C") descCCol = c;
  }
  if (subcatCol < 0) return empty;

  const COL_D  = 3;  // PLANOGRAM
  const COL_AL = 37; // colAL (kept for backward compat but not used in O anymore)

  const freqPlog = new Map<string, Map<string, number>>();
  const freqAL   = new Map<string, Map<string, number>>();
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

    const al = cellVal(ws, r, COL_AL);
    if (al) {
      if (!freqAL.has(prefix)) freqAL.set(prefix, new Map());
      const m = freqAL.get(prefix)!;
      m.set(al, (m.get(al) || 0) + 1);
    }

    // Build UPC-level meta for exception config matching
    if (upcCol >= 0) {
      const upc = normalizeBarcode(cellVal(ws, r, upcCol));
      if (upc && !byUpc.has(upc)) {
        const category   = categoryCol >= 0 ? cellVal(ws, r, categoryCol) : "";
        const descC      = descCCol   >= 0 ? cellVal(ws, r, descCCol)   : "";
        byUpc.set(upc, { category, subcategory: subcat, descC });
        if (category) catSet.add(category);
        if (subcat)   subSet.add(subcat);
        if (descC)    descSet.add(descC);
      }
    }
  }

  const allPrefixes = new Set([...freqPlog.keys(), ...freqAL.keys()]);
  const byPrefix = new Map<string, { planogram: string; colAL: string }>();
  for (const prefix of allPrefixes) {
    byPrefix.set(prefix, {
      planogram: freqPlog.has(prefix) ? mostFrequent(freqPlog.get(prefix)!) : "",
      colAL:     freqAL.has(prefix)   ? mostFrequent(freqAL.get(prefix)!)   : "",
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

    // Column O: percentage string. Match against exception config (first active rule wins).
    // Primary: look up the barcode in DATA_SPACEMAN's byUpc map for exact CATEGORY/SUBCATEGORY/DESC_C.
    // Fallback: derive meta from the RECAP hierarchy we already computed — Class → category,
    // SubDept → descC, SubclassCode+Name → subcategory. This covers products that have a
    // planogram prefix entry but whose individual UPC row is absent from DATA_SPACEMAN.
    const metaEntry = byUpc.get(row.barcode);
    const derivedMeta: SpacemanRowMeta = {
      category:    filled.cls,
      subcategory: info.subclassCode
        ? (info.subclassName.trim()
            ? `${info.subclassCode}: ${info.subclassName.trim()}`
            : info.subclassCode)
        : "",
      descC: filled.subDept,
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

    write(5,  data.division  || "");
    write(6,  data.dept      || "");
    write(7,  data.subDept   || "");
    write(8,  data.cls       || "");
    write(9,  data.planogram || ""); // J ← DATA_SPACEMAN col D
    write(13, data.colN      || ""); // N ← 100 ช่อง col DF
    write(14, data.colO      || ""); // O ← DATA_SPACEMAN col AL
  }

  XLSX.writeFile(wb, "RECAP_filled.xlsx");
}
