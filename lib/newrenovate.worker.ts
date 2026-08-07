/**
 * NewRenovate worker — patches cell values directly inside the template XLSX ZIP.
 * Preserves all original formatting: styles, colours, merged cells, column widths,
 * freeze panes, conditional formatting, etc.
 *
 * Flow:
 *  1. XLSX.read source files (DATA_SPACEMAN, INDEX, QRY) → lookup maps
 *  2. XLSX.read target template (no styles) → detect header cols + read barcodes
 *  3. Build row→column patch map
 *  4. unzipSync(targetBuf) → patch sheet XMLs → append SST strings → zipSync
 */

import * as XLSX from "xlsx";
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import type { ExceptionConfig } from "./types";

// Worker postMessage with transfer support
const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

// ─── Types ─────────────────────────────────────────────────────────────────────

type InMsg = {
  type: "run";
  targetBuf:   ArrayBuffer;
  spacemanBuf: ArrayBuffer;
  indexBuf:    ArrayBuffer;
  qryBuf:      ArrayBuffer;
  exceptionConfig: ExceptionConfig[];
};

interface Stats {
  sheet1: { processed: number; matched: number };
  sheet2: { processed: number; matched: number };
}

interface SpacemanEntry {
  planofolder02: string;
  planofolder03: string;
  planofolder04: string;
  planogram: string;
  category: string;
  subcategory: string;
  descC: string;
}

interface QryEntry {
  segment: string;
  locationId: string;
  totalUnits: string;
}

type CellPatch =
  | { t: "s"; v: string }
  | { t: "n"; v: number }
  | { t: "f"; f: string; v: number };

// ─── Helpers ───────────────────────────────────────────────────────────────────

function progress(pct: number, msg: string) {
  ctx.postMessage({ type: "progress", pct, msg });
}

function normalizeBarcode(val: unknown): string {
  if (val == null || val === "") return "";
  const s = String(val).trim();
  if (!s) return "";
  const n = Number(s);
  if (!isNaN(n) && n > 0) return String(Math.round(n));
  return s;
}

function getOrderingPct(cfg: ExceptionConfig[], cat: string, sub: string, descC: string): number {
  for (const rule of cfg) {
    if (rule.status === "inactive" || rule.status === "deleted") continue;
    const catOk  = rule.category    === "ทั้งหมด" || rule.category    === cat;
    const subOk  = rule.subcategory === "ทั้งหมด" || rule.subcategory === sub;
    const descOk = rule.descC       === "ทั้งหมด" || rule.descC       === descC;
    if (catOk && subOk && descOk) return Number(rule.percentage) / 100;
  }
  return 1.0;
}

function cellStrXlsx(ws: XLSX.WorkSheet, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  return cell?.v != null ? String(cell.v).trim() : "";
}

// ─── XML / ZIP helpers (same pattern as download.worker.ts) ───────────────────

function encodeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function decodeXml(s: string): string {
  return s
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

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

function colLetterIdx(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64;
  return n - 1;
}

function parseSST(xml: string): string[] {
  const out: string[] = [];
  const siRe = /<si>([\s\S]*?)<\/si>/g;
  let m: RegExpExecArray | null;
  while ((m = siRe.exec(xml)) !== null) {
    let text = "";
    const tRe = /<t(?:\s[^>]*)?>([^<]*)<\/t>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRe.exec(m[1])) !== null) text += decodeXml(tm[1]);
    out.push(text);
  }
  return out;
}

function appendSST(xml: string, newStrings: string[]): string {
  if (!newStrings.length) return xml;
  const newSis = newStrings.map(s => `<si><t>${encodeXml(s)}</t></si>`).join("");
  const at = xml.lastIndexOf("</sst>");
  let result = xml.slice(0, at) + newSis + xml.slice(at);
  result = result
    .replace(/\bcount="(\d+)"/, (_, n) => `count="${+n + newStrings.length}"`)
    .replace(/\buniqueCount="(\d+)"/, (_, n) => `uniqueCount="${+n + newStrings.length}"`);
  return result;
}

function buildSST(strings: string[]): string {
  const n = strings.length;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
    ` count="${n}" uniqueCount="${n}">` +
    strings.map(s => `<si><t>${encodeXml(s)}</t></si>`).join("") +
    `</sst>`
  );
}

function findSheetPath(wbXml: string, relsXml: string, name: string): string | null {
  const xmlName = encodeXml(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const sheetMatch = new RegExp(`<sheet\\b[^>]*name="${xmlName}"[^>]*/?>`, "i").exec(wbXml);
  if (!sheetMatch) return null;
  const ridMatch = /\br:id="([^"]+)"/.exec(sheetMatch[0]);
  if (!ridMatch) return null;
  const rm = new RegExp(`<Relationship\\b[^>]+Id="${ridMatch[1]}"[^>]+Target="([^"]+)"`, "i").exec(relsXml);
  if (!rm) return null;
  const t = rm[1];
  if (t.startsWith("/xl/")) return t.slice(1);
  if (t.startsWith("worksheets/")) return `xl/${t}`;
  return `xl/worksheets/${t}`;
}

function getSheetNamesFromXml(wbXml: string): string[] {
  const names: string[] = [];
  const pat = /<sheet\b[^>]*name="([^"]*)"[^>]*/g;
  let m: RegExpExecArray | null;
  while ((m = pat.exec(wbXml)) !== null) names.push(decodeXml(m[1]));
  return names;
}

// ─── Cell patching ────────────────────────────────────────────────────────────

function buildCellXml(ref: string, patch: CellPatch, getSsIdx: (v: string) => number, sAttr: string): string {
  if (patch.t === "s") return `<c r="${ref}"${sAttr} t="s"><v>${getSsIdx(patch.v)}</v></c>`;
  if (patch.t === "n") return `<c r="${ref}"${sAttr}><v>${patch.v}</v></c>`;
  return `<c r="${ref}"${sAttr}><f>${patch.f}</f><v>${patch.v}</v></c>`;
}

function patchCellInRow(inner: string, ref: string, ci: number, patch: CellPatch, getSsIdx: (v: string) => number): string {
  const pat = new RegExp(`<c r="${ref}"([^>]*?)(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const m = pat.exec(inner);
  const sM = m ? /\bs="(\d+)"/.exec(m[1]) : null;
  const sAttr = sM ? ` s="${sM[1]}"` : "";
  const newCell = buildCellXml(ref, patch, getSsIdx, sAttr);

  if (m) return inner.slice(0, m.index) + newCell + inner.slice(m.index + m[0].length);

  // Insert in column order
  const scanPat = /<c\s+r="([A-Z]+)\d+"/g;
  let at = -1;
  let im: RegExpExecArray | null;
  while ((im = scanPat.exec(inner)) !== null) {
    if (colLetterIdx(im[1]) > ci) { at = im.index; break; }
  }
  return at >= 0 ? inner.slice(0, at) + newCell + inner.slice(at) : inner + newCell;
}

/**
 * Apply rowPatches to sheetXml.
 * rowPatches: Map<rowNum (1-based), Map<colIdx (0-based), CellPatch>>
 */
function patchSheetXml(
  sheetXml: string,
  sstStrings: string[],
  rowPatches: Map<number, Map<number, CellPatch>>
): { sheetXml: string; newStrings: string[] } {
  if (!rowPatches.size) return { sheetXml, newStrings: [] };

  const allStrings = [...sstStrings];
  const getSsIdx = (v: string): number => {
    let i = allStrings.indexOf(v);
    if (i < 0) { i = allStrings.length; allStrings.push(v); }
    return i;
  };

  // Pass 1: expand self-closing rows that need patching
  let result = sheetXml.replace(/<row\b([^>]*?)\/>/g, (full, attrs) => {
    const rm = /\br="(\d+)"/.exec(attrs);
    if (!rm) return full;
    const patches = rowPatches.get(+rm[1]);
    if (!patches) return full;
    const rowNum = +rm[1];
    let cells = "";
    for (const [ci, patch] of [...patches.entries()].sort((a, b) => a[0] - b[0])) {
      const ref = `${colLetter(ci)}${rowNum}`;
      cells += buildCellXml(ref, patch, getSsIdx, "");
    }
    return `<row${attrs}>${cells}</row>`;
  });

  // Pass 2: patch open/close rows
  result = result.replace(/(<row\b[^>]*>)([\s\S]*?)(<\/row>)/g, (full, open, inner, close) => {
    const rm = /\br="(\d+)"/.exec(open);
    if (!rm) return full;
    const patches = rowPatches.get(+rm[1]);
    if (!patches) return full;
    const rowNum = +rm[1];
    let cells = inner;
    for (const [ci, patch] of patches) {
      cells = patchCellInRow(cells, `${colLetter(ci)}${rowNum}`, ci, patch, getSsIdx);
    }
    return open + cells + close;
  });

  return { sheetXml: result, newStrings: allStrings.slice(sstStrings.length) };
}

// ─── Main worker handler ───────────────────────────────────────────────────────

addEventListener("message", (e: MessageEvent<InMsg>) => {
  if (e.data.type !== "run") return;
  const { targetBuf, spacemanBuf, indexBuf, qryBuf, exceptionConfig } = e.data;

  try {
    // ── 1. Parse DATA_SPACEMAN ────────────────────────────────────────────────
    progress(5, "อ่านไฟล์ DATA_SPACEMAN...");

    const spacemanWb = XLSX.read(spacemanBuf, { type: "array", cellText: false, cellHTML: false, cellNF: false, cellDates: false });
    const spacemanWs = spacemanWb.Sheets["QRY_Product_by_POG"];
    if (!spacemanWs) throw new Error('ไม่พบ sheet "QRY_Product_by_POG" ใน DATA_SPACEMAN');

    const sRange = XLSX.utils.decode_range(spacemanWs["!ref"] || "A1");
    const hdrs: string[] = [];
    for (let c = 0; c <= sRange.e.c; c++) hdrs.push(cellStrXlsx(spacemanWs, 0, c));

    const upcIdx   = hdrs.indexOf("UPC");
    const pf02Idx  = hdrs.indexOf("PLANOFOLDER02");
    const pf03Idx  = hdrs.indexOf("PLANOFOLDER03");
    const pf04Idx  = hdrs.indexOf("PLANOFOLDER04");
    const plogIdx  = hdrs.indexOf("PLANOGRAM") >= 0 ? hdrs.indexOf("PLANOGRAM") : 3;
    const catIdx   = hdrs.indexOf("CATEGORY");
    const subIdx   = hdrs.indexOf("SUBCATEGORY");
    const descCIdx = hdrs.indexOf("DESC_C");

    const spacemanMap = new Map<string, SpacemanEntry>();
    for (let r = 1; r <= sRange.e.r; r++) {
      const upc = upcIdx >= 0 ? normalizeBarcode(cellStrXlsx(spacemanWs, r, upcIdx)) : "";
      if (!upc) continue;
      if (!spacemanMap.has(upc)) {
        spacemanMap.set(upc, {
          planofolder02: pf02Idx  >= 0 ? cellStrXlsx(spacemanWs, r, pf02Idx)  : "",
          planofolder03: pf03Idx  >= 0 ? cellStrXlsx(spacemanWs, r, pf03Idx)  : "",
          planofolder04: pf04Idx  >= 0 ? cellStrXlsx(spacemanWs, r, pf04Idx)  : "",
          planogram:     cellStrXlsx(spacemanWs, r, plogIdx),
          category:      catIdx   >= 0 ? cellStrXlsx(spacemanWs, r, catIdx)   : "",
          subcategory:   subIdx   >= 0 ? cellStrXlsx(spacemanWs, r, subIdx)   : "",
          descC:         descCIdx >= 0 ? cellStrXlsx(spacemanWs, r, descCIdx) : "",
        });
      }
      if (r % 10000 === 0)
        progress(5 + Math.floor((r / sRange.e.r) * 25), `DATA_SPACEMAN: ${r.toLocaleString()} rows...`);
    }
    progress(30, `DATA_SPACEMAN: ${spacemanMap.size.toLocaleString()} barcodes`);

    // ── 2. Parse INDEX ────────────────────────────────────────────────────────
    progress(32, "อ่านไฟล์ INDEX...");

    const indexWb = XLSX.read(indexBuf, { type: "array" });
    const indexWs = indexWb.Sheets[indexWb.SheetNames[0]];
    if (!indexWs) throw new Error("ไม่พบ sheet ใน INDEX");

    type IndexRow = Record<string, unknown>;
    const indexRows = XLSX.utils.sheet_to_json<IndexRow>(indexWs, { defval: "" });
    const indexMap = new Map<string, { status: string; store: string }>();
    for (const row of indexRows) {
      const plog = String(row["PLANOGRAM"] ?? row["Planogram"] ?? row["POG"] ?? "").trim();
      if (!plog) continue;
      if (!indexMap.has(plog)) {
        indexMap.set(plog, {
          status: String(row["Status"] ?? row["STATUS"] ?? row["สถานะ"] ?? ""),
          store:  String(row["Store"]  ?? row["STORE"]  ?? row["สาขา"]  ?? ""),
        });
      }
    }
    progress(42, `INDEX: ${indexMap.size.toLocaleString()} planograms`);

    // ── 3. Parse QRY_Product_by_POG_by_Position ───────────────────────────────
    progress(44, "อ่านไฟล์ QRY_Product_by_POG_by_Position...");

    const qryWb = XLSX.read(qryBuf, { type: "array" });
    const qryWs = qryWb.Sheets[qryWb.SheetNames[0]];
    if (!qryWs) throw new Error("ไม่พบ sheet ใน QRY_Product_by_POG_by_Position");

    type QryRow = Record<string, unknown>;
    const qryRows = XLSX.utils.sheet_to_json<QryRow>(qryWs, { defval: "" });
    const qryMap = new Map<string, QryEntry>();
    for (const row of qryRows) {
      const bc = normalizeBarcode(row["BARCODE"] ?? row["UPC"] ?? row["Barcode"]);
      if (!bc) continue;
      if (!qryMap.has(bc)) {
        qryMap.set(bc, {
          segment:    String(row["SEGMENT"]     ?? ""),
          locationId: String(row["LOCATION_ID"] ?? ""),
          totalUnits: String(row["TOTAL_UNITS"] ?? ""),
        });
      }
    }
    progress(55, `QRY Position: ${qryMap.size.toLocaleString()} barcodes`);

    // ── 4. Parse target template (XLSX.read — for header/barcode detection only) ──
    progress(57, "อ่านไฟล์ Template (detect headers)...");

    // NOTE: XLSX.read does not detach the ArrayBuffer, so we can still use
    // targetBuf later in unzipSync for the format-preserving ZIP patch.
    const targetWb = XLSX.read(targetBuf, { type: "array", cellText: false, cellHTML: false, cellNF: false, cellDates: false });

    // ── 5. Build patches for Sheet 1 ─────────────────────────────────────────
    progress(60, "สร้าง patch data Sheet 1...");

    const s1Name = targetWb.SheetNames.find(n => n.trimStart().startsWith("New&Exsiting For Oder"));
    if (!s1Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New&Exsiting For Oder'");
    const ws1 = targetWb.Sheets[s1Name];

    const ref1 = XLSX.utils.decode_range(ws1["!ref"] ?? "A1:V7998");
    const lastRow1 = Math.min(ref1.e.r + 1, 7998);

    // Detect column positions from row 6 (index 5)
    const colMap1 = new Map<string, number>();
    for (let c = 0; c <= ref1.e.c; c++) {
      const h = cellStrXlsx(ws1, 5, c);
      if (h) colMap1.set(h, c);
    }

    const c1 = (name: string, fb: number) => colMap1.get(name) ?? fb;
    const findCol = (needle: string) => {
      for (const [k, v] of colMap1)
        if (k.toLowerCase().includes(needle.toLowerCase())) return v;
      return undefined;
    };

    const BARCODE_COL    = c1("BARCODE", 5);
    const DIVISION_COL   = c1("DIVISION", 1);
    const PF03_COL       = colMap1.get("PLANOFOLDER03") ?? 2;
    const PF04_COL       = colMap1.get("PLANOFOLDER04") ?? 3;
    const STATUS_COL     = findCol("Status");
    const STORE_COL      = findCol("Store");
    const FIXTYPE_COL    = findCol("Fixture Type");
    const W_COL          = findCol(" W") ?? findCol("W");
    const H_COL          = findCol(" H") ?? findCol("H");
    const D_COL          = findCol(" D") ?? findCol("D");
    const NEWFIXTURE_COL = findCol("New Fixture");
    const NOBAY_COL      = findCol("No.Bay");
    const SEQ_COL        = findCol("SEQ");
    const SHELFSTOCK_COL = findCol("SHELF STOCK") ?? c1("SHELF STOCK FOR ORDER (Piece)", 19);
    const PCT_COL        = findCol("% Ordering") ?? 20;
    const NETCAP_COL     = findCol("Net Capacity") ?? 21;

    // Map<rowNum (1-based), Map<colIdx, CellPatch>>
    const s1Patches = new Map<number, Map<number, CellPatch>>();

    let s1Processed = 0, s1Matched = 0;

    for (let row1 = 7; row1 <= lastRow1; row1++) {
      const r = row1 - 1;
      const isSpecialRow = row1 <= 8;
      const bc = normalizeBarcode(ws1[XLSX.utils.encode_cell({ r, c: BARCODE_COL })]?.v);

      const rowCols = new Map<number, CellPatch>();

      if (!isSpecialRow && bc) {
        s1Processed++;
        const spaceman = spacemanMap.get(bc);
        const qry      = qryMap.get(bc);
        const planogram = spaceman?.planogram ?? "";
        if (spaceman || qry) s1Matched++;

        const setS = (col: number | undefined, v: string) => {
          if (col !== undefined && v) rowCols.set(col, { t: "s", v });
        };
        const setN = (col: number | undefined, v: unknown) => {
          if (col === undefined || v == null || v === "") return;
          const n = Number(v);
          if (!isNaN(n)) rowCols.set(col, { t: "n", v: n });
        };

        setS(DIVISION_COL, spaceman?.planofolder02 ?? "");
        setS(PF03_COL,     spaceman?.planofolder03 ?? "");
        setS(PF04_COL,     spaceman?.planofolder04 ?? "");

        const idx = planogram ? indexMap.get(planogram) : undefined;
        setS(STATUS_COL, idx?.status ?? "");
        setS(STORE_COL,  idx?.store  ?? "");

        if (FIXTYPE_COL    !== undefined) rowCols.set(FIXTYPE_COL,    { t: "n", v: 0 });
        if (W_COL          !== undefined) rowCols.set(W_COL,          { t: "n", v: 2 });
        if (H_COL          !== undefined) rowCols.set(H_COL,          { t: "n", v: 1 });
        if (D_COL          !== undefined) rowCols.set(D_COL,          { t: "n", v: 1 });
        if (NEWFIXTURE_COL !== undefined) rowCols.set(NEWFIXTURE_COL, { t: "s", v: "" });

        setS(NOBAY_COL, qry?.segment    ?? "");
        setS(SEQ_COL,   qry?.locationId ?? "");
        setN(SHELFSTOCK_COL, qry?.totalUnits);
      }

      if (!isSpecialRow || bc) {
        const spaceman = bc ? spacemanMap.get(bc) : undefined;
        const pct = getOrderingPct(exceptionConfig, spaceman?.category ?? "", spaceman?.subcategory ?? "", spaceman?.descC ?? "");
        rowCols.set(PCT_COL, { t: "n", v: pct });
        // Net Capacity as formula referencing the actual cells
        const tRef = `${colLetter(SHELFSTOCK_COL)}${row1}`;
        const uRef = `${colLetter(PCT_COL)}${row1}`;
        rowCols.set(NETCAP_COL, { t: "f", f: `${tRef}*${uRef}`, v: 0 });
      }

      if (rowCols.size) s1Patches.set(row1, rowCols);
    }

    // ── 6. Build patches for Sheet 2 ─────────────────────────────────────────
    progress(70, "สร้าง patch data Sheet 2...");

    const s2Name = targetWb.SheetNames.find(n => n.trim().startsWith("New for Link_IM"));
    if (!s2Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New for Link_IM'");
    const ws2 = targetWb.Sheets[s2Name];

    const ref2 = XLSX.utils.decode_range(ws2["!ref"] ?? "A1:H15522");
    const lastRow2 = Math.min(ref2.e.r + 1, 15522);

    const colMap2 = new Map<string, number>();
    for (let c = 0; c <= ref2.e.c; c++) {
      const h = cellStrXlsx(ws2, 5, c);
      if (h) colMap2.set(h, c);
    }
    const BC2_COL   = colMap2.get("BARCODE")    ?? 4;
    const DIV2_COL  = colMap2.get("DIVISION")   ?? 1;
    const DEPT2_COL = colMap2.get("DEPARTMENT") ?? 2;

    const s2Patches = new Map<number, Map<number, CellPatch>>();
    let s2Processed = 0, s2Matched = 0;

    for (let row2 = 7; row2 <= lastRow2; row2++) {
      s2Processed++;
      const bc = normalizeBarcode(ws2[XLSX.utils.encode_cell({ r: row2 - 1, c: BC2_COL })]?.v);
      const spaceman = bc ? spacemanMap.get(bc) : undefined;
      if (spaceman) {
        s2Matched++;
        const rowCols = new Map<number, CellPatch>();
        if (spaceman.planofolder02) rowCols.set(DIV2_COL,  { t: "s", v: spaceman.planofolder02 });
        if (spaceman.planofolder03) rowCols.set(DEPT2_COL, { t: "s", v: spaceman.planofolder03 });
        if (rowCols.size) s2Patches.set(row2, rowCols);
      }
    }

    // ── 7. ZIP-patch the template (format-preserving) ─────────────────────────
    progress(78, "เปิด Template ZIP...");

    const files = unzipSync(new Uint8Array(targetBuf));
    const wbXml   = strFromU8(files["xl/workbook.xml"]);
    const relsXml = strFromU8(files["xl/_rels/workbook.xml.rels"]);
    const sstPath  = "xl/sharedStrings.xml";
    const sstStrings = files[sstPath] ? parseSST(strFromU8(files[sstPath])) : [];

    // Accumulate all new strings across both sheets
    const allNewStrings: string[] = [];

    // Patch Sheet 1
    const path1 = findSheetPath(wbXml, relsXml, s1Name);
    if (!path1 || !files[path1]) throw new Error(`ไม่พบ path สำหรับ sheet "${s1Name}" ใน ZIP`);
    progress(82, "Patch Sheet 1...");
    const r1 = patchSheetXml(strFromU8(files[path1]), [...sstStrings, ...allNewStrings], s1Patches);
    files[path1] = strToU8(r1.sheetXml);
    allNewStrings.push(...r1.newStrings);

    // Patch Sheet 2
    const path2 = findSheetPath(wbXml, relsXml, s2Name);
    if (!path2 || !files[path2]) throw new Error(`ไม่พบ path สำหรับ sheet "${s2Name}" ใน ZIP`);
    progress(88, "Patch Sheet 2...");
    const r2 = patchSheetXml(strFromU8(files[path2]), [...sstStrings, ...allNewStrings], s2Patches);
    files[path2] = strToU8(r2.sheetXml);
    allNewStrings.push(...r2.newStrings);

    // Update SST
    if (allNewStrings.length > 0) {
      files[sstPath] = strToU8(
        files[sstPath]
          ? appendSST(strFromU8(files[sstPath]), allNewStrings)
          : buildSST([...sstStrings, ...allNewStrings])
      );
    }

    // Rezip — every other file (styles.xml, workbook.xml, rels, images …) unchanged
    progress(94, "บีบอัดไฟล์ผลลัพธ์...");
    const zipped = zipSync(files);
    const output = zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength);
    progress(100, "เสร็จสิ้น!");

    ctx.postMessage(
      { type: "done", buffer: output, stats: { sheet1: { processed: s1Processed, matched: s1Matched }, sheet2: { processed: s2Processed, matched: s2Matched } } },
      [output]
    );

  } catch (err) {
    ctx.postMessage({ type: "error", message: String(err) });
  }
});
