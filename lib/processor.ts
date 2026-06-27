import * as XLSX from "xlsx";
import type {
  MissingRow,
  SubclassInfo,
  ProcessedRow,
  HierarchyNames,
  FilledData,
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

function colLetter(ws: XLSX.WorkSheet): number {
  const ref = ws["!ref"];
  if (!ref) return 0;
  return XLSX.utils.decode_range(ref).e.c;
}

/** "04 DRY FOOD"  →  "04: DRY FOOD" */
function formatLevel(full: string): string {
  if (!full) return "";
  const parts = full.trim().split(" ");
  const num = parts[0];
  const name = parts.slice(1).join(" ");
  return `${num}: ${name}`;
}

/** Build hierarchical RECAP-format codes from the 4 level strings */
function buildRecapCodes(h: HierarchyNames): FilledData {
  const div = h.divFull.trim().split(" ")[0]; // "04"
  const dep = h.deptFull.trim().split(" ")[0]; // "20"
  const sub = h.subdeptFull.trim().split(" ")[0]; // "60"
  const cls = h.clsFull.trim().split(" ")[0]; // "01"

  const divName = h.divFull.trim().slice(div.length).trim(); // "DRY FOOD"
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
    colO: "",      // filled later by parsePlanogramLookup col AL
  };
}

// ─── Step 1: Parse RECAP ───────────────────────────────────────────────────

export function parseMissingRows(wb: XLSX.WorkBook): MissingRow[] {
  const ws = wb.Sheets["NEW SCM"];
  if (!ws) throw new Error('ไม่พบชีต "NEW SCM" ในไฟล์ RECAP');

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const results: MissingRow[] = [];

  for (let r = 4; r <= range.e.r; r++) {
    const barcode = cellVal(ws, r, 3); // col D
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

// ─── Step 2: Parse xlsb/xlsx source files ─────────────────────────────────

export async function parseXlsbFiles(
  files: File[]
): Promise<Map<string, SubclassInfo>> {
  const map = new Map<string, SubclassInfo>();

  for (const file of files) {
    const buf = await file.arrayBuffer();
    let wb: XLSX.WorkBook;
    try {
      wb = XLSX.read(buf, { type: "array", bookSheets: true });
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

      // Find header row (search first 70 rows)
      for (let r = 0; r <= Math.min(range.e.r, 70); r++) {
        for (let c = 0; c <= Math.min(range.e.c, 250); c++) {
          const v = cellVal(ws, r, c);
          if (v.includes("Barcode / PLU")) barcodeCol = c;
          if (v.includes("Sub-Class") && v.includes("รหัส")) subclassCodeCol = c;
          if (v === "Sub-Class Name ชื่อโครงสร้างสินค้า") subclassNameCol = c;
        }
        if (barcodeCol >= 0 && subclassCodeCol >= 0) break;
      }

      if (barcodeCol < 0 || subclassCodeCol < 0) continue;

      for (let r = 0; r <= range.e.r; r++) {
        const bc = cellVal(ws, r, barcodeCol);
        const code = cellVal(ws, r, subclassCodeCol);
        if (bc && code && bc.length >= 7 && code.length === 10) {
          if (!map.has(bc)) {
            map.set(bc, {
              subclassCode: code,
              subclassName:
                subclassNameCol >= 0 ? cellVal(ws, r, subclassNameCol) : "",
              sourceFile: file.name,
              colDF: cellVal(ws, r, 109), // col DF (0-based index 109) → RECAP col N
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
      const code = cellVal(ws, r, 14); // SUB_CLASS_CODE
      if (!code || code.length !== 10) continue;

      const divFull = cellVal(ws, r, 9);
      const deptFull = cellVal(ws, r, 10);
      const subdeptFull = cellVal(ws, r, 11);
      const clsFull = cellVal(ws, r, 12);

      if (divFull && deptFull && subdeptFull && clsFull && !map.has(code)) {
        map.set(code, { divFull, deptFull, subdeptFull, clsFull });
      }
    }
  }

  return map;
}

export interface PlanogramEntry {
  planogram: string; // most-frequent value from col D (index 3)
  colAL: string;     // most-frequent value from col AL (index 37)
}

// ─── Step 4: Parse DATA_SPACEMAN → subdept prefix → planogram + colAL ────

export async function parsePlanogramLookup(
  file: File,
  onProgress?: (pct: number) => void
): Promise<Map<string, PlanogramEntry>> {
  const map = new Map<string, PlanogramEntry>(); // key = 6-digit subdept code

  const buf = await file.arrayBuffer();
  onProgress?.(20);

  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buf, { type: "array" });
  } catch {
    return map;
  }
  onProgress?.(60);

  const ws = wb.Sheets["QRY_Product_by_POG"];
  if (!ws) return map;

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

  // Find SUBCATEGORY column by header name (position may vary)
  let subcatCol = -1;
  for (let c = 0; c <= range.e.c; c++) {
    if (cellVal(ws, 0, c) === "SUBCATEGORY") { subcatCol = c; break; }
  }
  if (subcatCol < 0) return map;

  // Planogram: col D (fixed index 3) — new format
  // ColAL: col AL (fixed index 37)
  const COL_D = 3;
  const COL_AL = 37;

  const freqPlog = new Map<string, Map<string, number>>();
  const freqAL   = new Map<string, Map<string, number>>();

  for (let r = 1; r <= range.e.r; r++) {
    const subcat = cellVal(ws, r, subcatCol); // e.g. "0420600103: CRACKERS"
    if (!subcat) continue;

    const prefix = subcat.slice(0, 6); // "042060"

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
  }

  const mostFrequent = (freq: Map<string, number>): string => {
    let best = ""; let bestCount = 0;
    for (const [val, count] of freq.entries()) {
      if (count > bestCount) { best = val; bestCount = count; }
    }
    return best;
  };

  const allPrefixes = new Set([...freqPlog.keys(), ...freqAL.keys()]);
  for (const prefix of allPrefixes) {
    map.set(prefix, {
      planogram: freqPlog.has(prefix) ? mostFrequent(freqPlog.get(prefix)!) : "",
      colAL:     freqAL.has(prefix)   ? mostFrequent(freqAL.get(prefix)!)   : "",
    });
  }

  onProgress?.(100);
  return map;
}

// ─── Step 5: Combine everything ────────────────────────────────────────────

export function processRows(
  missing: MissingRow[],
  barcodeMap: Map<string, SubclassInfo>,
  structureMap: Map<string, HierarchyNames>,
  planogramMap: Map<string, PlanogramEntry>
): ProcessedRow[] {
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
    const subdeptPrefix = info.subclassCode.slice(0, 6); // e.g. "042060"
    const entry = planogramMap.get(subdeptPrefix);
    filled.planogram = entry?.planogram || "";
    filled.colN      = info.colDF || "";        // 100 ช่อง col DF → RECAP col N
    filled.colO      = entry?.colAL || "";      // DATA_SPACEMAN col AL → RECAP col O

    const confidence = filled.planogram ? "confirmed" : "inferred";

    return {
      ...row,
      filled,
      confidence,
      subclassCode: info.subclassCode,
      note: confidence === "confirmed"
        ? `จาก ${info.sourceFile}`
        : `จาก ${info.sourceFile} — PLANOGRAM ไม่พบใน DATA_SPACEMAN`,
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
