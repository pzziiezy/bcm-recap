/**
 * NewRenovate worker — ZIP-patch approach (preserves all template formatting).
 *
 * Data flow (QRY-driven):
 *
 *   QRY_Product_by_POG_by_Position  ← primary source of rows
 *     │  BARCODE → DATA_SPACEMAN    → DIVISION(PF02) / PF03 / PF04 / PLANOGRAM / CAT / SUB / DESC_C
 *     │  BARCODE → Master Assortment→ SALE PACK CODE (BAR_SINGLE) / Pack Size (SKU_PACK) / Extra info (EXTRA_INFO)
 *     │  PLANOGRAM + SEGMENT
 *     │            → Fixture Index  → New Fixture (Code Fixture)
 *     │  PLANOGRAM → INDEX          → Status / Store
 *     │  QRY itself                 → No.Bay (SEGMENT) / SEQ (LOCATION_ID) / SHELF STOCK (TOTAL_UNITS)
 *     └─ Config Rules               → % Ordering (default 100%)
 *        Net Capacity = SHELF STOCK × % Ordering
 *
 * Template is written via ZIP-patch: unzip → patch sheet XMLs → append SST → rezip.
 */

import * as XLSX from "xlsx";
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import type { ExceptionConfig } from "./types";

const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

// ─── Types ─────────────────────────────────────────────────────────────────────

type InMsg = {
  type: "run";
  targetBuf:   ArrayBuffer;
  qryBuf:      ArrayBuffer;
  spacemanBuf: ArrayBuffer;
  masterBuf:   ArrayBuffer;
  indexBuf:    ArrayBuffer;
  fixtureBuf:  ArrayBuffer;
  exceptionConfig: ExceptionConfig[];
};

interface Stats {
  total: number;
  matchedSpaceman: number;
  matchedMaster: number;
  matchedIndex: number;
  matchedFixture: number;
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

interface MasterEntry {
  barSingle: string;
  skuPack: string;
  extraInfo: string;
}

interface QrySourceRow {
  barcode: string;
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
    if ((rule.category    === "ทั้งหมด" || rule.category    === cat)  &&
        (rule.subcategory === "ทั้งหมด" || rule.subcategory === sub)  &&
        (rule.descC       === "ทั้งหมด" || rule.descC       === descC))
      return Number(rule.percentage) / 100;
  }
  return 1.0;
}

// ─── XML / ZIP helpers ────────────────────────────────────────────────────────

function encodeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function decodeXml(s: string): string {
  return s.replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">")
          .replace(/&quot;/g,'"').replace(/&apos;/g,"'");
}
function colLetter(idx: number): string {
  let s = "", n = idx + 1;
  while (n > 0) { const r = (n-1)%26; s = String.fromCharCode(65+r)+s; n = Math.floor((n-1)/26); }
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
  let r = xml.slice(0, at) + newSis + xml.slice(at);
  r = r.replace(/\bcount="(\d+)"/,       (_,n) => `count="${+n + newStrings.length}"`)
       .replace(/\buniqueCount="(\d+)"/,  (_,n) => `uniqueCount="${+n + newStrings.length}"`);
  return r;
}
function buildSST(strings: string[]): string {
  const n = strings.length;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
    ` count="${n}" uniqueCount="${n}">` +
    strings.map(s => `<si><t>${encodeXml(s)}</t></si>`).join("") + `</sst>`;
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

// ─── Cell patching ────────────────────────────────────────────────────────────

function buildCellXml(ref: string, patch: CellPatch, getSsIdx: (v: string) => number, sAttr: string): string {
  if (patch.t === "s") return `<c r="${ref}"${sAttr} t="s"><v>${getSsIdx(patch.v)}</v></c>`;
  if (patch.t === "n") return `<c r="${ref}"${sAttr}><v>${patch.v}</v></c>`;
  return `<c r="${ref}"${sAttr}><f>${patch.f}</f><v>${patch.v}</v></c>`;
}

function patchCellInRow(inner: string, ref: string, ci: number, patch: CellPatch, getSsIdx: (v: string) => number): string {
  const pat = new RegExp(`<c r="${ref}"([^>]*?)(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const m = pat.exec(inner);
  const sAttr = m ? (/\bs="(\d+)"/.exec(m[1]) ? ` s="${/\bs="(\d+)"/.exec(m[1])![1]}"` : "") : "";
  const newCell = buildCellXml(ref, patch, getSsIdx, sAttr);
  if (m) return inner.slice(0, m.index) + newCell + inner.slice(m.index + m[0].length);
  const scanPat = /<c\s+r="([A-Z]+)\d+"/g;
  let at = -1, im: RegExpExecArray | null;
  while ((im = scanPat.exec(inner)) !== null) {
    if (colLetterIdx(im[1]) > ci) { at = im.index; break; }
  }
  return at >= 0 ? inner.slice(0, at) + newCell + inner.slice(at) : inner + newCell;
}

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
  const patchedRows = new Set<number>();

  // Pass 1: self-closing rows
  let result = sheetXml.replace(/<row\b([^>]*?)\/>/g, (full, attrs) => {
    const rm = /\br="(\d+)"/.exec(attrs);
    if (!rm) return full;
    const patches = rowPatches.get(+rm[1]);
    if (!patches) return full;
    patchedRows.add(+rm[1]);
    let cells = "";
    for (const [ci, patch] of [...patches.entries()].sort((a, b) => a[0] - b[0]))
      cells += buildCellXml(`${colLetter(ci)}${rm[1]}`, patch, getSsIdx, "");
    return `<row${attrs}>${cells}</row>`;
  });

  // Pass 2: open/close rows
  result = result.replace(/(<row\b[^>]*>)([\s\S]*?)(<\/row>)/g, (full, open, inner, close) => {
    const rm = /\br="(\d+)"/.exec(open);
    if (!rm) return full;
    const patches = rowPatches.get(+rm[1]);
    if (!patches) return full;
    patchedRows.add(+rm[1]);
    let cells = inner;
    for (const [ci, patch] of patches)
      cells = patchCellInRow(cells, `${colLetter(ci)}${rm[1]}`, ci, patch, getSsIdx);
    return open + cells + close;
  });

  // Pass 3: insert missing rows
  const newRowXml = [...rowPatches.entries()]
    .filter(([rowNum]) => !patchedRows.has(rowNum))
    .sort((a, b) => a[0] - b[0])
    .map(([rowNum, patches]) => {
      let cells = "";
      for (const [ci, patch] of [...patches.entries()].sort((a, b) => a[0] - b[0]))
        cells += buildCellXml(`${colLetter(ci)}${rowNum}`, patch, getSsIdx, "");
      return `<row r="${rowNum}">${cells}</row>`;
    }).join("");

  if (newRowXml) {
    const sdClose = result.lastIndexOf("</sheetData>");
    const sdSelf  = result.indexOf("<sheetData/>");
    if (sdClose >= 0) {
      result = result.slice(0, sdClose) + newRowXml + result.slice(sdClose);
    } else if (sdSelf >= 0) {
      result = result.slice(0, sdSelf) + `<sheetData>${newRowXml}</sheetData>` + result.slice(sdSelf + 12);
    } else {
      const wsEnd = result.lastIndexOf("</worksheet>");
      if (wsEnd >= 0)
        result = result.slice(0, wsEnd) + `<sheetData>${newRowXml}</sheetData>` + result.slice(wsEnd);
    }
  }

  return { sheetXml: result, newStrings: allStrings.slice(sstStrings.length) };
}

// ─── Main worker handler ───────────────────────────────────────────────────────

addEventListener("message", (e: MessageEvent<InMsg>) => {
  if (e.data.type !== "run") return;
  const { targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf, fixtureBuf, exceptionConfig } = e.data;

  try {
    // ── 1. QRY_Product_by_POG_by_Position → ordered list of rows ─────────────
    progress(3, "อ่านไฟล์ QRY_Product_by_POG_by_Position...");

    const qryWb = XLSX.read(qryBuf, { type: "array" });
    const qryWs = qryWb.Sheets[qryWb.SheetNames[0]];
    if (!qryWs) throw new Error("ไม่พบ sheet ใน QRY_Product_by_POG_by_Position");

    const qrySourceRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(qryWs, { defval: "" });
    const qryRows: QrySourceRow[] = [];
    for (const row of qrySourceRows) {
      const bc = normalizeBarcode(row["BARCODE"] ?? row["UPC"] ?? row["Barcode"]);
      if (!bc) continue;
      qryRows.push({
        barcode:    bc,
        segment:    String(row["SEGMENT"]     ?? ""),
        locationId: String(row["LOCATION_ID"] ?? ""),
        totalUnits: String(row["TOTAL_UNITS"] ?? ""),
      });
    }
    progress(12, `QRY: ${qryRows.length.toLocaleString()} rows (ข้อมูลตั้งต้น)`);

    // ── 2. DATA_SPACEMAN → map by BARCODE/UPC ────────────────────────────────
    progress(14, "อ่านไฟล์ DATA_SPACEMAN...");

    const spacemanWb = XLSX.read(spacemanBuf, { type: "array", cellText: false, cellHTML: false, cellNF: false, cellDates: false });
    const spacemanWs = spacemanWb.Sheets["QRY_Product_by_POG"];
    if (!spacemanWs) throw new Error('ไม่พบ sheet "QRY_Product_by_POG" ใน DATA_SPACEMAN');

    const sRange = XLSX.utils.decode_range(spacemanWs["!ref"] || "A1");
    const sHdrs: string[] = [];
    for (let c = 0; c <= sRange.e.c; c++) {
      const cell = spacemanWs[XLSX.utils.encode_cell({ r: 0, c })];
      sHdrs.push(cell?.v != null ? String(cell.v).trim() : "");
    }
    const upcIdx   = sHdrs.indexOf("UPC");
    const pf02Idx  = sHdrs.indexOf("PLANOFOLDER02");
    const pf03Idx  = sHdrs.indexOf("PLANOFOLDER03");
    const pf04Idx  = sHdrs.indexOf("PLANOFOLDER04");
    const plogIdx  = sHdrs.indexOf("PLANOGRAM") >= 0 ? sHdrs.indexOf("PLANOGRAM") : 3;
    const catIdx   = sHdrs.indexOf("CATEGORY");
    const subIdx   = sHdrs.indexOf("SUBCATEGORY");
    const descCIdx = sHdrs.indexOf("DESC_C");

    const getS = (r: number, c: number) => {
      const cell = spacemanWs[XLSX.utils.encode_cell({ r, c })];
      return cell?.v != null ? String(cell.v).trim() : "";
    };

    const spacemanMap = new Map<string, SpacemanEntry>();
    for (let r = 1; r <= sRange.e.r; r++) {
      const upc = upcIdx >= 0 ? normalizeBarcode(getS(r, upcIdx)) : "";
      if (!upc) continue;
      if (!spacemanMap.has(upc)) {
        spacemanMap.set(upc, {
          planofolder02: pf02Idx  >= 0 ? getS(r, pf02Idx)  : "",
          planofolder03: pf03Idx  >= 0 ? getS(r, pf03Idx)  : "",
          planofolder04: pf04Idx  >= 0 ? getS(r, pf04Idx)  : "",
          planogram:     getS(r, plogIdx),
          category:      catIdx   >= 0 ? getS(r, catIdx)   : "",
          subcategory:   subIdx   >= 0 ? getS(r, subIdx)   : "",
          descC:         descCIdx >= 0 ? getS(r, descCIdx) : "",
        });
      }
      if (r % 10000 === 0)
        progress(14 + Math.floor((r / sRange.e.r) * 20), `DATA_SPACEMAN: ${r.toLocaleString()} rows...`);
    }
    progress(34, `DATA_SPACEMAN: ${spacemanMap.size.toLocaleString()} barcodes`);

    // ── 3. Master Assortment Orderable → map by BARCODE ───────────────────────
    progress(36, "อ่านไฟล์ Master Assortment Orderable...");

    const masterWb = XLSX.read(new Uint8Array(masterBuf), { type: "array" });
    const masterSheetName =
      masterWb.SheetNames.find(n => masterWb.Sheets[n]) ??
      masterWb.SheetNames[0];
    if (!masterSheetName)
      throw new Error(`Master Assortment: ไม่พบ sheet ใดๆ ในไฟล์ — SheetNames: [${masterWb.SheetNames.join(", ") || "ว่าง"}]`);
    const masterWs = masterWb.Sheets[masterSheetName];
    if (!masterWs)
      throw new Error(`Master Assortment: ไม่พบ sheet "${masterSheetName}" — SheetNames: [${masterWb.SheetNames.join(", ")}]`);

    const masterRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(masterWs, { defval: "" });
    const masterMap = new Map<string, MasterEntry>();
    for (const row of masterRows) {
      const bc = normalizeBarcode(row["BARCODE"]);
      if (!bc) continue;
      if (!masterMap.has(bc)) {
        masterMap.set(bc, {
          barSingle: String(row["BAR_SINGLE"] ?? ""),
          skuPack:   String(row["SKU_PACK"]   ?? ""),
          extraInfo: String(row["EXTRA_INFO"] ?? ""),
        });
      }
    }
    progress(46, `Master Assortment: ${masterMap.size.toLocaleString()} barcodes`);

    // ── 4. INDEX → map by PLANOGRAM ───────────────────────────────────────────
    progress(48, "อ่านไฟล์ INDEX...");

    const indexWb = XLSX.read(new Uint8Array(indexBuf), { type: "array" });
    const indexSheetName = indexWb.SheetNames.find(n => indexWb.Sheets[n]) ?? indexWb.SheetNames[0];
    if (!indexSheetName) throw new Error(`INDEX: ไม่พบ sheet — SheetNames: [${indexWb.SheetNames.join(", ") || "ว่าง"}]`);
    const indexWs = indexWb.Sheets[indexSheetName];
    if (!indexWs) throw new Error(`INDEX: ไม่พบ sheet "${indexSheetName}"`);

    const indexRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(indexWs, { defval: "" });
    const indexMap = new Map<string, { status: string; store: string }>();
    for (const row of indexRows) {
      const plog = String(row["PLANOGRAM"] ?? row["Planogram"] ?? row["POG"] ?? "").trim();
      if (!plog) continue;
      if (!indexMap.has(plog)) {
        indexMap.set(plog, {
          status: String(row["Status"]  ?? row["STATUS"] ?? row["สถานะ"] ?? ""),
          store:  String(row["Store"]   ?? row["STORE"]  ?? row["สาขา"]  ?? ""),
        });
      }
    }
    progress(55, `INDEX: ${indexMap.size.toLocaleString()} planograms`);

    // ── 5. Fixture Index → map by "SEG|POG" ───────────────────────────────────
    progress(57, "อ่านไฟล์ Fixture Index...");

    const fixtureWb = XLSX.read(new Uint8Array(fixtureBuf), { type: "array" });
    const fixtureSheetName =
      fixtureWb.SheetNames.find(n => n === "Fixture_2026") ??
      fixtureWb.SheetNames.find(n => n.toLowerCase().startsWith("fixture")) ??
      fixtureWb.SheetNames.find(n => fixtureWb.Sheets[n]) ??
      fixtureWb.SheetNames[0];
    if (!fixtureSheetName)
      throw new Error(`Fixture Index: ไม่พบ sheet — SheetNames: [${fixtureWb.SheetNames.join(", ") || "ว่าง"}]`);
    const fixtureWs = fixtureWb.Sheets[fixtureSheetName];
    if (!fixtureWs)
      throw new Error(`Fixture Index: ไม่พบ sheet "${fixtureSheetName}" — SheetNames: [${fixtureWb.SheetNames.join(", ")}]`);
    // range:1 = skip remark row, row 2 becomes headers (SEG / POG / Code Fixture)
    const fixtureRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(fixtureWs, { defval: "", range: 1 });

    const fixtureMap = new Map<string, string>(); // "SEG|POG" → Code Fixture
    for (const row of fixtureRows) {
      const seg  = String(row["SEG"]          ?? "").trim();
      const pog  = String(row["POG"]          ?? "").trim();
      const code = String(row["Code Fixture"] ?? "").trim();
      if (seg && pog && code) fixtureMap.set(`${seg}|${pog}`, code);
    }
    progress(63, `Fixture Index: ${fixtureMap.size.toLocaleString()} SEG|POG entries`);

    // ── 6. Detect column positions from template ───────────────────────────────
    progress(65, "อ่านไฟล์ Template (detect headers)...");

    // XLSX.read does NOT detach ArrayBuffer — safe to unzip later
    const targetWb = XLSX.read(targetBuf, { type: "array", cellText: false, cellHTML: false, cellNF: false, cellDates: false });

    const s1Name = targetWb.SheetNames.find(n => n.trimStart().startsWith("New&Exsiting For Oder"));
    if (!s1Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New&Exsiting For Oder'");
    const ws1 = targetWb.Sheets[s1Name];
    const ref1 = XLSX.utils.decode_range(ws1["!ref"] ?? "A1:V7");

    const colMap1 = new Map<string, number>();
    for (let c = 0; c <= Math.max(ref1.e.c, 25); c++) {
      const cell = ws1[XLSX.utils.encode_cell({ r: 5, c })];
      const h = cell?.v != null ? String(cell.v).trim() : "";
      if (h) colMap1.set(h, c);
    }

    const c1  = (name: string, fb: number) => colMap1.get(name) ?? fb;
    const fc1 = (needle: string) => {
      for (const [k, v] of colMap1)
        if (k.toLowerCase().includes(needle.toLowerCase())) return v;
      return undefined;
    };

    const BARCODE_COL    = c1("BARCODE", 5);
    const DIVISION_COL   = c1("DIVISION", 1);
    const PF03_COL       = colMap1.get("PLANOFOLDER03") ?? 2;
    const PF04_COL       = colMap1.get("PLANOFOLDER04") ?? 3;
    const SALEPACK_COL   = fc1("SALE PACK CODE") ?? fc1("SALE PACK") ?? 7;
    const PACKSIZE_COL   = fc1("Pack Size") ?? 8;
    const EXTRA_COL      = fc1("Extra info") ?? 9;
    const STATUS_COL     = fc1("Status");
    const STORE_COL      = fc1("Store");
    const FIXTYPE_COL    = fc1("Fixture Type");
    const W_COL          = fc1(" W") ?? fc1("W");
    const H_COL          = fc1(" H") ?? fc1("H");
    const D_COL          = fc1(" D") ?? fc1("D");
    const NEWFIXTURE_COL = fc1("New Fixture");
    const NOBAY_COL      = fc1("No.Bay");
    const SEQ_COL        = fc1("SEQ");
    const SHELFSTOCK_COL = fc1("SHELF STOCK") ?? c1("SHELF STOCK FOR ORDER (Piece)", 19);
    const PCT_COL        = fc1("% Ordering") ?? 20;
    const NETCAP_COL     = fc1("Net Capacity") ?? 21;

    const s2Name = targetWb.SheetNames.find(n => n.trim().startsWith("New for Link_IM"));
    if (!s2Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New for Link_IM'");
    const ws2 = targetWb.Sheets[s2Name];
    const ref2 = XLSX.utils.decode_range(ws2["!ref"] ?? "A1:H7");

    const colMap2 = new Map<string, number>();
    for (let c = 0; c <= Math.max(ref2.e.c, 10); c++) {
      const cell = ws2[XLSX.utils.encode_cell({ r: 5, c })];
      const h = cell?.v != null ? String(cell.v).trim() : "";
      if (h) colMap2.set(h, c);
    }
    const BARCODE2_COL = colMap2.get("BARCODE")    ?? 4;
    const DIV2_COL     = colMap2.get("DIVISION")   ?? 1;
    const DEPT2_COL    = colMap2.get("DEPARTMENT") ?? 2;

    // ── 7. Build patches row-by-row from QRY ──────────────────────────────────
    progress(68, `สร้าง patches จาก ${qryRows.length.toLocaleString()} QRY rows...`);

    const DATA_START_ROW = 7;
    const s1Patches = new Map<number, Map<number, CellPatch>>();
    const s2Patches = new Map<number, Map<number, CellPatch>>();

    let matchedSpaceman = 0, matchedMaster = 0, matchedIndex = 0, matchedFixture = 0;

    for (let i = 0; i < qryRows.length; i++) {
      const qry      = qryRows[i];
      const rowNum   = DATA_START_ROW + i;
      const spaceman = spacemanMap.get(qry.barcode);
      const master   = masterMap.get(qry.barcode);
      const planogram = spaceman?.planogram ?? "";
      const idxEntry  = planogram ? indexMap.get(planogram) : undefined;
      const fixtureKey = qry.segment && planogram ? `${qry.segment}|${planogram}` : "";
      const fixtureCode = fixtureKey ? (fixtureMap.get(fixtureKey) ?? "") : "";

      if (spaceman)    matchedSpaceman++;
      if (master)      matchedMaster++;
      if (idxEntry)    matchedIndex++;
      if (fixtureCode) matchedFixture++;

      // ── Sheet 1 ──────────────────────────────────────────────────────────
      const cols1 = new Map<number, CellPatch>();

      const setS1 = (col: number | undefined, v: string) => {
        if (col !== undefined && v !== "") cols1.set(col, { t: "s", v });
      };
      const setN1 = (col: number | undefined, v: unknown) => {
        if (col === undefined) return;
        const n = Number(v);
        if (!isNaN(n)) cols1.set(col, { t: "n", v: n });
      };

      // BARCODE (from QRY)
      cols1.set(BARCODE_COL, { t: "s", v: qry.barcode });

      // From DATA_SPACEMAN
      setS1(DIVISION_COL, spaceman?.planofolder02 ?? "");
      setS1(PF03_COL,     spaceman?.planofolder03 ?? "");
      setS1(PF04_COL,     spaceman?.planofolder04 ?? "");

      // From Master Assortment
      if (master) {
        const barSingleNum = Number(master.barSingle);
        if (!isNaN(barSingleNum) && master.barSingle !== "") {
          cols1.set(SALEPACK_COL, { t: "n", v: barSingleNum });
        } else {
          setS1(SALEPACK_COL, master.barSingle);
        }
        setN1(PACKSIZE_COL, master.skuPack);
        setS1(EXTRA_COL,    master.extraInfo);
      }

      // From INDEX
      setS1(STATUS_COL, idxEntry?.status ?? "");
      setS1(STORE_COL,  idxEntry?.store  ?? "");

      // Constants
      if (FIXTYPE_COL !== undefined) cols1.set(FIXTYPE_COL, { t: "n", v: 0 });
      if (W_COL       !== undefined) cols1.set(W_COL,       { t: "n", v: 2 });
      if (H_COL       !== undefined) cols1.set(H_COL,       { t: "n", v: 1 });
      if (D_COL       !== undefined) cols1.set(D_COL,       { t: "n", v: 1 });

      // From Fixture Index
      setS1(NEWFIXTURE_COL, fixtureCode);

      // From QRY itself
      setS1(NOBAY_COL, qry.segment);
      setS1(SEQ_COL,   qry.locationId);
      setN1(SHELFSTOCK_COL, qry.totalUnits || undefined);

      // % Ordering (Config Rules or default 100%)
      const pct = getOrderingPct(exceptionConfig, spaceman?.category ?? "", spaceman?.subcategory ?? "", spaceman?.descC ?? "");
      cols1.set(PCT_COL, { t: "n", v: pct });

      // Net Capacity = SHELF_STOCK_CELL × PCT_CELL
      cols1.set(NETCAP_COL, { t: "f", f: `${colLetter(SHELFSTOCK_COL)}${rowNum}*${colLetter(PCT_COL)}${rowNum}`, v: 0 });

      s1Patches.set(rowNum, cols1);

      // ── Sheet 2 ──────────────────────────────────────────────────────────
      const cols2 = new Map<number, CellPatch>();
      cols2.set(BARCODE2_COL, { t: "s", v: qry.barcode });
      if (spaceman?.planofolder02) cols2.set(DIV2_COL,  { t: "s", v: spaceman.planofolder02 });
      if (spaceman?.planofolder03) cols2.set(DEPT2_COL, { t: "s", v: spaceman.planofolder03 });
      s2Patches.set(rowNum, cols2);
    }

    // ── 8. ZIP-patch template ─────────────────────────────────────────────────
    progress(75, "เปิด Template ZIP...");

    const files   = unzipSync(new Uint8Array(targetBuf));
    const wbXml   = strFromU8(files["xl/workbook.xml"]);
    const relsXml = strFromU8(files["xl/_rels/workbook.xml.rels"]);
    const sstPath = "xl/sharedStrings.xml";
    const sstStrings = files[sstPath] ? parseSST(strFromU8(files[sstPath])) : [];

    const allNewStrings: string[] = [];

    const path1 = findSheetPath(wbXml, relsXml, s1Name);
    if (!path1 || !files[path1]) throw new Error(`ไม่พบ ZIP path สำหรับ sheet "${s1Name}"`);
    progress(80, `Patch Sheet 1 (${s1Patches.size.toLocaleString()} rows)...`);
    const r1 = patchSheetXml(strFromU8(files[path1]), [...sstStrings, ...allNewStrings], s1Patches);
    files[path1] = strToU8(r1.sheetXml);
    allNewStrings.push(...r1.newStrings);

    const path2 = findSheetPath(wbXml, relsXml, s2Name);
    if (!path2 || !files[path2]) throw new Error(`ไม่พบ ZIP path สำหรับ sheet "${s2Name}"`);
    progress(90, `Patch Sheet 2 (${s2Patches.size.toLocaleString()} rows)...`);
    const r2 = patchSheetXml(strFromU8(files[path2]), [...sstStrings, ...allNewStrings], s2Patches);
    files[path2] = strToU8(r2.sheetXml);
    allNewStrings.push(...r2.newStrings);

    if (allNewStrings.length > 0) {
      files[sstPath] = strToU8(
        files[sstPath]
          ? appendSST(strFromU8(files[sstPath]), allNewStrings)
          : buildSST([...sstStrings, ...allNewStrings])
      );
    }

    progress(96, "บีบอัดไฟล์ผลลัพธ์...");
    const zipped = zipSync(files);
    const output = zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength);
    progress(100, "เสร็จสิ้น!");

    ctx.postMessage(
      { type: "done", buffer: output, stats: { total: qryRows.length, matchedSpaceman, matchedMaster, matchedIndex, matchedFixture } },
      [output]
    );

  } catch (err) {
    ctx.postMessage({ type: "error", message: String(err) });
  }
});
