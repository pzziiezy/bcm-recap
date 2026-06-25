import * as XLSX from "xlsx";
import type {
  FilledData,
  HierarchyNames,
  MissingRow,
  ProcessedRow,
  SubclassInfo,
} from "./types";

function cellVal(ws: XLSX.WorkSheet, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell) return "";

  const value = cell.v;
  if (value == null) return "";

  const stringValue = String(value);
  return stringValue.includes(".") && !isNaN(Number(stringValue))
    ? stringValue.split(".")[0]
    : stringValue;
}

function buildRecapCodes(hierarchy: HierarchyNames): FilledData {
  const div = hierarchy.divFull.trim().split(" ")[0];
  const dep = hierarchy.deptFull.trim().split(" ")[0];
  const sub = hierarchy.subdeptFull.trim().split(" ")[0];
  const cls = hierarchy.clsFull.trim().split(" ")[0];

  const divName = hierarchy.divFull.trim().slice(div.length).trim();
  const depName = hierarchy.deptFull.trim().slice(dep.length).trim();
  const subName = hierarchy.subdeptFull.trim().slice(sub.length).trim();
  const clsName = hierarchy.clsFull.trim().slice(cls.length).trim();

  return {
    division: `${div}: ${divName}`,
    dept: `${div}${dep}: ${depName}`,
    subDept: `${div}${dep}${sub}: ${subName}`,
    cls: `${div}${dep}${sub}${cls}: ${clsName}`,
    planogram: "",
  };
}

function fillBarcodeLookup(
  ws: XLSX.WorkSheet,
  sourceFile: string,
  barcodeMap: Map<string, SubclassInfo>
): void {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  let barcodeCol = -1;
  let subclassCodeCol = -1;
  let subclassNameCol = -1;

  for (let r = 0; r <= Math.min(range.e.r, 70); r++) {
    for (let c = 0; c <= Math.min(range.e.c, 250); c++) {
      const value = cellVal(ws, r, c);
      if (value.includes("Barcode / PLU")) barcodeCol = c;
      if (value.includes("Sub-Class") && value.includes("เธฃเธซเธฑเธช")) {
        subclassCodeCol = c;
      }
      if (value === "Sub-Class Name เธเธทเนเธญเนเธเธฃเธเธชเธฃเนเธฒเธเธชเธดเธเธเนเธฒ") {
        subclassNameCol = c;
      }
    }

    if (barcodeCol >= 0 && subclassCodeCol >= 0) break;
  }

  if (barcodeCol < 0 || subclassCodeCol < 0) return;

  for (let r = 0; r <= range.e.r; r++) {
    const barcode = cellVal(ws, r, barcodeCol);
    const subclassCode = cellVal(ws, r, subclassCodeCol);

    if (
      !barcode ||
      !subclassCode ||
      barcode.length < 7 ||
      subclassCode.length !== 10 ||
      barcodeMap.has(barcode)
    ) {
      continue;
    }

    barcodeMap.set(barcode, {
      subclassCode,
      subclassName:
        subclassNameCol >= 0 ? cellVal(ws, r, subclassNameCol) : "",
      sourceFile,
    });
  }
}

function fillStructureLookup(
  ws: XLSX.WorkSheet,
  structureMap: Map<string, HierarchyNames>
): void {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

  for (let r = 1; r <= range.e.r; r++) {
    const code = cellVal(ws, r, 14);
    if (!code || code.length !== 10 || structureMap.has(code)) continue;

    const divFull = cellVal(ws, r, 9);
    const deptFull = cellVal(ws, r, 10);
    const subdeptFull = cellVal(ws, r, 11);
    const clsFull = cellVal(ws, r, 12);

    if (!divFull || !deptFull || !subdeptFull || !clsFull) continue;

    structureMap.set(code, { divFull, deptFull, subdeptFull, clsFull });
  }
}

export function parseMissingRows(wb: XLSX.WorkBook): MissingRow[] {
  const ws = wb.Sheets["NEW SCM"];
  if (!ws) throw new Error('ไม่พบชีต "NEW SCM" ในไฟล์ RECAP');

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const results: MissingRow[] = [];

  for (let r = 4; r <= range.e.r; r++) {
    const barcode = cellVal(ws, r, 3);
    const division = cellVal(ws, r, 5);

    if (!barcode || division) continue;

    results.push({
      rowIndex: r,
      barcode,
      name: cellVal(ws, r, 4),
    });
  }

  return results;
}

export async function parseXlsbFiles(
  files: File[]
): Promise<Map<string, SubclassInfo>> {
  return (await buildXlsbLookups(files)).barcodeMap;
}

export async function buildStructureLookup(
  files: File[]
): Promise<Map<string, HierarchyNames>> {
  return (await buildXlsbLookups(files)).structureMap;
}

export async function buildXlsbLookups(
  files: File[]
): Promise<{
  barcodeMap: Map<string, SubclassInfo>;
  structureMap: Map<string, HierarchyNames>;
}> {
  const barcodeMap = new Map<string, SubclassInfo>();
  const structureMap = new Map<string, HierarchyNames>();

  for (const file of files) {
    const buffer = await file.arrayBuffer();

    let workbook: XLSX.WorkBook;
    try {
      workbook = XLSX.read(buffer, { type: "array" });
    } catch {
      continue;
    }

    for (const sheetName of ["Base", "Input"]) {
      const ws = workbook.Sheets[sheetName];
      if (ws) fillBarcodeLookup(ws, file.name, barcodeMap);
    }

    const structureSheet = workbook.Sheets["Sh_ProdStructure"];
    if (structureSheet) fillStructureLookup(structureSheet, structureMap);
  }

  return { barcodeMap, structureMap };
}

export async function parsePlanogramLookup(
  file: File,
  onProgress?: (pct: number) => void
): Promise<Map<string, string>> {
  const map = new Map<string, string>();

  const buffer = await file.arrayBuffer();
  onProgress?.(20);

  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, { type: "array" });
  } catch {
    return map;
  }

  onProgress?.(60);

  const ws = workbook.Sheets["QRY_Product_by_POG"];
  if (!ws) return map;

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  let subcategoryCol = -1;
  let planogramCol = -1;

  for (let c = 0; c <= range.e.c; c++) {
    const header = cellVal(ws, 0, c);
    if (header === "SUBCATEGORY") subcategoryCol = c;
    if (header === "PLANOGRAM") planogramCol = c;
  }

  if (subcategoryCol < 0 || planogramCol < 0) return map;

  const frequencies = new Map<string, Map<string, number>>();

  for (let r = 1; r <= range.e.r; r++) {
    const subcategory = cellVal(ws, r, subcategoryCol);
    const planogram = cellVal(ws, r, planogramCol);
    if (!subcategory || !planogram) continue;

    const prefix = subcategory.slice(0, 6);
    const root = planogram.split("_")[0].trim();
    if (!root) continue;

    if (!frequencies.has(prefix)) frequencies.set(prefix, new Map());
    const counts = frequencies.get(prefix)!;
    counts.set(root, (counts.get(root) || 0) + 1);
  }

  for (const [prefix, counts] of frequencies.entries()) {
    let bestRoot = "";
    let bestCount = 0;

    for (const [root, count] of counts.entries()) {
      if (count > bestCount) {
        bestRoot = root;
        bestCount = count;
      }
    }

    map.set(prefix, bestRoot);
  }

  onProgress?.(100);
  return map;
}

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
        note: "ไม่พบบาร์โค้ดในไฟล์ 100 ช่อง กรุณากรอกเอง",
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
    filled.planogram = planogramMap.get(subdeptPrefix) || "";

    const confidence = filled.planogram ? "confirmed" : "inferred";

    return {
      ...row,
      filled,
      confidence,
      subclassCode: info.subclassCode,
      note:
        confidence === "confirmed"
          ? `จาก ${info.sourceFile}`
          : `จาก ${info.sourceFile} - PLANOGRAM ไม่พบใน DATA_SPACEMAN`,
    };
  });
}

export function applyAndDownload(
  wb: XLSX.WorkBook,
  rows: ProcessedRow[]
): void {
  const ws = wb.Sheets["NEW SCM"];

  for (const row of rows) {
    const data = row.override ? { ...row.filled, ...row.override } : row.filled;
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
