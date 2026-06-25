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
    planogram: "", // filled later
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

// ─── Step 4: Parse DATA_SPACEMAN → subdept prefix → planogram ─────────────

export async function parsePlanogramLookup(
  file: File,
  onProgress?: (pct: number) => void
): Promise<Map<string, string>> {
  const map = new Map<string, string>(); // key = 6-digit subdept code

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

  // Find SUBCATEGORY and PLANOGRAM columns from header row 0
  let subcatCol = -1;
  let plogCol = -1;
  for (let c = 0; c <= range.e.c; c++) {
    const h = cellVal(ws, 0, c);
    if (h === "SUBCATEGORY") subcatCol = c;
    if (h === "PLANOGRAM") plogCol = c;
  }
  if (subcatCol < 0 || plogCol < 0) return map;

  const freq = new Map<string, Map<string, number>>();

  for (let r = 1; r <= range.e.r; r++) {
    const subcat = cellVal(ws, r, subcatCol); // e.g. "0420600103: CRACKERS"
    const plog = cellVal(ws, r, plogCol); // e.g. "BISCUITS_4B_H180"
    if (!subcat || !plog) continue;

    const prefix = subcat.slice(0, 6); // "042060"
    // Take only the root name (before first underscore)
    const root = plog.split("_")[0].trim();
    if (!root) continue;

    if (!freq.has(prefix)) freq.set(prefix, new Map());
    const m = freq.get(prefix)!;
    m.set(root, (m.get(root) || 0) + 1);
  }

  // Pick most-frequent planogram root per subdept prefix
  for (const [prefix, counts] of freq.entries()) {
    let best = "";
    let bestCount = 0;
    for (const [root, count] of counts.entries()) {
      if (count > bestCount) {
        best = root;
        bestCount = count;
      }
    }
    map.set(prefix, best);
  }

  onProgress?.(100);
  return map;
}

// ─── Step 5: Combine everything ────────────────────────────────────────────

export function processRows(
  missing: MissingRow[],
  barcodeMap: Map<string, SubclassInfo>,
  structureMap: Map<string, HierarchyNames>,
  planogramMap: Map<string, string>
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
    filled.planogram = planogramMap.get(subdeptPrefix) || "";

    const confidence =
      filled.planogram
        ? "confirmed"
        : "inferred";

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

    write(5, data.division || "");
    write(6, data.dept || "");
    write(7, data.subDept || "");
    write(8, data.cls || "");
    write(9, data.planogram || "");
  }

  XLSX.writeFile(wb, "RECAP_filled.xlsx");
}
