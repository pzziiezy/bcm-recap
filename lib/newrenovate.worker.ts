/**
 * NewRenovate worker — ZIP-patch approach (preserves all template formatting).
 *
 * Data flow (QRY-driven):
 *
 *   QRY_Product_by_POG_by_Position  ← primary source of rows
 *     │  BARCODE → DATA_SPACEMAN    → DIVISION(PF01) / PF03 / PF04 / PLANOGRAM
 *     │  BARCODE → Master Assortment→ SALE PACK CODE / Pack Size / Extra info / Status / Store / Name
 *     │  PLANOGRAM + SEGMENT → Fixture Index → New Fixture (Code Fixture)
 *     │  PLANOGRAM → INDEX   → Status(fallback) / Store(fallback) / PLANOGRAM NAME
 *     │  QRY itself          → No.Bay(SEGMENT) / SEQ(LOCATION_ID) / SHELF STOCK(TOTAL_UNITS) / Name
 *     └─ Config Rules        → % Ordering (default 100, stored as integer)
 *        Net Capacity = SHELF STOCK × (% Ordering / 100)
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

interface SpacemanEntry {
  planofolder01: string;
  planofolder02: string;
  planofolder03: string;
  planofolder04: string;
  planogram:     string;
  category:      string;
  subcategory:   string;
  descC:         string;
}

interface MasterEntry {
  name:      string;
  barSingle: string;
  skuPack:   string;
  extraInfo: string;
  status:    string;
  store:     string;
}

interface IndexEntry {
  status:        string;
  store:         string;
  planogramName: string;
}

interface QrySourceRow {
  barcode:    string;
  segment:    string;
  locationId: string;
  totalUnits: string;
  name:       string;
}

type CellPatch =
  | { t: "s"; v: string }
  | { t: "n"; v: number }
  | { t: "f"; f: string; v: number };

// ─── Helpers ───────────────────────────────────────────────────────────────────

function progress(pct: number, msg: string) {
  ctx.postMessage({ type: "progress", pct, msg });
}

/**
 * Normalize a barcode to a canonical string.
 * KEY FIX: if the value is already a string that starts with "0",
 * preserve the leading zeros — don't convert to Number and back.
 * Also accepts an optional `formatted` (cell.w) which preserves Excel custom-number-format
 * leading zeros (e.g., format "0000000000000" applied to a number).
 */
function normalizeBarcode(val: unknown, formatted?: string): string {
  // Prefer formatted text (w property from cellText:true) — it preserves
  // leading zeros that come from Excel custom number formats.
  if (formatted != null && formatted !== "") {
    const w = formatted.replace(/[, ]/g, "").trim();
    if (w) return w;
  }
  if (val == null || val === "") return "";
  const s = String(val).trim();
  if (!s) return "";
  // Preserve strings that already have leading zeros (text-type cells)
  if (s.startsWith("0")) return s;
  const n = Number(s);
  if (!isNaN(n) && n > 0) return String(Math.round(n));
  return s;
}

/**
 * Returns the ordering percentage as an integer (0–100).
 * Default = 100. Matches Config Rules by DIVISION(PF01)+PF03+PF04.
 */
function getOrderingPct(
  cfg: ExceptionConfig[],
  pf01: string,
  pf03: string,
  pf04: string,
): number {
  for (const rule of cfg) {
    if (rule.status === "inactive" || rule.status === "deleted") continue;
    const catOk = rule.category    === "ทั้งหมด" || rule.category    === pf01;
    const subOk = rule.subcategory === "ทั้งหมด" || rule.subcategory === pf03;
    const dscOk = rule.descC       === "ทั้งหมด" || rule.descC       === pf04;
    if (catOk && subOk && dscOk) return Number(rule.percentage);
  }
  return 100;
}

/** Normalize key for INDEX map: trim + collapse whitespace */
function normalizeKey(s: string): string {
  return s.replace(/\s+/g, " ").trim();
}

// ─── XML / ZIP helpers ────────────────────────────────────────────────────────

function encodeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function decodeXml(s: string): string {
  return s
    .replace(/&amp;/g,  "&").replace(/&lt;/g,  "<").replace(/&gt;/g,  ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}
function colLetter(idx: number): string {
  let s = "", n = idx + 1;
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
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
  r = r.replace(/\bcount="(\d+)"/,      (_, n) => `count="${+n + newStrings.length}"`)
       .replace(/\buniqueCount="(\d+)"/, (_, n) => `uniqueCount="${+n + newStrings.length}"`);
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
  rowPatches: Map<number, Map<number, CellPatch>>,
): { sheetXml: string; newStrings: string[] } {
  if (!rowPatches.size) return { sheetXml, newStrings: [] };

  const allStrings = [...sstStrings];
  const getSsIdx = (v: string): number => {
    let i = allStrings.indexOf(v);
    if (i < 0) { i = allStrings.length; allStrings.push(v); }
    return i;
  };
  const patchedRows = new Set<number>();

  // Pass 1: self-closing rows <row r="N" .../>
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

  // Pass 2: open/close rows <row r="N" ...>...</row>
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

  // Pass 3: insert rows that have no XML element yet
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
    // ── 1. QRY → ordered row list ─────────────────────────────────────────────
    progress(3, "อ่านไฟล์ QRY_Product_by_POG_by_Position...");

    // cellText:true → cell.w (formatted text) is populated, needed for barcode leading zeros.
    const qryWb = XLSX.read(new Uint8Array(qryBuf), { type: "array", cellText: true });
    const qryWs = qryWb.Sheets[qryWb.SheetNames[0]];
    if (!qryWs) throw new Error("ไม่พบ sheet ใน QRY_Product_by_POG_by_Position");

    const qryRange = XLSX.utils.decode_range(qryWs["!ref"] || "A1");

    // Auto-detect header row: scan first 5 rows for a row that contains a barcode-like column.
    let qHdrRow = 0;
    for (let r = 0; r <= Math.min(4, qryRange.e.r); r++) {
      for (let c = 0; c <= qryRange.e.c; c++) {
        const v = String(qryWs[XLSX.utils.encode_cell({ r, c })]?.v ?? "").toUpperCase().trim();
        if (v === "BARCODE" || v === "UPC" || v === "EAN") { qHdrRow = r; break; }
      }
      if (qHdrRow === r && r > 0) break;
    }

    // Build header array (case-preserved for display, searched case-insensitively below)
    const qHdrs: string[] = [];
    for (let c = 0; c <= qryRange.e.c; c++) {
      const cell = qryWs[XLSX.utils.encode_cell({ r: qHdrRow, c })];
      qHdrs.push(cell?.v != null ? String(cell.v).trim() : "");
    }

    // Case-insensitive column finder with multiple name fallbacks
    const findQCol = (...names: string[]): number => {
      const targets = new Set(names.map(n => n.toLowerCase()));
      return qHdrs.findIndex(h => targets.has(h.toLowerCase()));
    };

    const qBarcodeCol = findQCol("BARCODE", "UPC", "EAN", "Barcode", "barcode");
    const qSegCol     = findQCol("SEGMENT", "Segment");
    const qLocCol     = findQCol("LOCATION_ID", "LOCATIONID", "Location_ID");
    const qUnitsCol   = findQCol("TOTAL_UNITS", "TOTALUNITS", "Total_Units", "QTY", "UNITS");
    // NAME column in QRY is the primary source for product name.
    const qNameCol    = findQCol("NAME", "PRODUCT_NAME", "DESCRIPTION", "LONG_DESC", "SHORT_DESC");

    if (qBarcodeCol < 0)
      throw new Error(`QRY: ไม่พบ column BARCODE — headers ที่พบ: [${qHdrs.filter(Boolean).join(", ")}]`);

    const qryRows: QrySourceRow[] = [];
    for (let r = qHdrRow + 1; r <= qryRange.e.r; r++) {
      const barcodeCell = qryWs[XLSX.utils.encode_cell({ r, c: qBarcodeCol })];
      // Prefer formatted text (w) to recover leading zeros from Excel custom number formats.
      const bc = normalizeBarcode(barcodeCell?.v, barcodeCell?.w);
      if (!bc) continue;

      const str = (col: number): string => {
        if (col < 0) return "";
        const cell = qryWs[XLSX.utils.encode_cell({ r, c: col })];
        if (!cell) return "";
        return (cell.w != null ? String(cell.w) : cell.v != null ? String(cell.v) : "").trim();
      };
      const numStr = (col: number): string => {
        if (col < 0) return "";
        const cell = qryWs[XLSX.utils.encode_cell({ r, c: col })];
        return cell?.v != null ? String(cell.v) : "";
      };

      qryRows.push({
        barcode:    bc,
        segment:    str(qSegCol),
        locationId: str(qLocCol),
        totalUnits: numStr(qUnitsCol),
        name:       str(qNameCol),
      });
    }
    progress(12, `QRY: ${qryRows.length.toLocaleString()} rows`);

    // ── 2. DATA_SPACEMAN → map by UPC ────────────────────────────────────────
    progress(14, "อ่านไฟล์ DATA_SPACEMAN...");

    const spacemanWb = XLSX.read(new Uint8Array(spacemanBuf), {
      type: "array", cellText: true, cellHTML: false, cellNF: false, cellDates: false,
    });
    const spacemanWs = spacemanWb.Sheets["QRY_Product_by_POG"];
    if (!spacemanWs) throw new Error('ไม่พบ sheet "QRY_Product_by_POG" ใน DATA_SPACEMAN');

    const sRange = XLSX.utils.decode_range(spacemanWs["!ref"] || "A1");
    const sHdrs: string[] = [];
    for (let c = 0; c <= sRange.e.c; c++) {
      const cell = spacemanWs[XLSX.utils.encode_cell({ r: 0, c })];
      sHdrs.push(cell?.v != null ? String(cell.v).trim() : "");
    }
    const upcIdx   = sHdrs.indexOf("UPC");
    const pf01Idx  = sHdrs.indexOf("PLANOFOLDER01");
    const pf02Idx  = sHdrs.indexOf("PLANOFOLDER02");
    const pf03Idx  = sHdrs.indexOf("PLANOFOLDER03");
    const pf04Idx  = sHdrs.indexOf("PLANOFOLDER04");
    const plogIdx  = sHdrs.indexOf("PLANOGRAM") >= 0 ? sHdrs.indexOf("PLANOGRAM") : 3;
    const catIdx   = sHdrs.indexOf("CATEGORY");
    const subIdx   = sHdrs.indexOf("SUBCATEGORY");
    const descCIdx = sHdrs.indexOf("DESC_C");

    const getS = (r: number, c: number, useW = false): string => {
      const cell = spacemanWs[XLSX.utils.encode_cell({ r, c })];
      if (!cell) return "";
      if (useW && cell.w != null) return String(cell.w).replace(/,/g, "").trim();
      return cell.v != null ? String(cell.v).trim() : "";
    };

    const spacemanMap = new Map<string, SpacemanEntry>();
    for (let r = 1; r <= sRange.e.r; r++) {
      const upc = upcIdx >= 0 ? normalizeBarcode(getS(r, upcIdx), getS(r, upcIdx, true)) : "";
      if (!upc) continue;
      if (!spacemanMap.has(upc)) {
        spacemanMap.set(upc, {
          planofolder01: pf01Idx >= 0 ? getS(r, pf01Idx) : "",
          planofolder02: pf02Idx >= 0 ? getS(r, pf02Idx) : "",
          planofolder03: pf03Idx >= 0 ? getS(r, pf03Idx) : "",
          planofolder04: pf04Idx >= 0 ? getS(r, pf04Idx) : "",
          planogram:     getS(r, plogIdx),
          category:      catIdx   >= 0 ? getS(r, catIdx)   : "",
          subcategory:   subIdx   >= 0 ? getS(r, subIdx)   : "",
          descC:         descCIdx >= 0 ? getS(r, descCIdx) : "",
        });
      }
      if (r % 10000 === 0)
        progress(14 + Math.floor((r / sRange.e.r) * 18), `DATA_SPACEMAN: ${r.toLocaleString()} rows...`);
    }
    progress(32, `DATA_SPACEMAN: ${spacemanMap.size.toLocaleString()} barcodes`);

    // ── 3. Master Assortment → map by BARCODE ────────────────────────────────
    progress(34, "อ่านไฟล์ Master Assortment Orderable...");

    const masterWb = XLSX.read(new Uint8Array(masterBuf), { type: "array", cellText: true });
    const masterSheetName = masterWb.SheetNames.find(n => masterWb.Sheets[n]) ?? masterWb.SheetNames[0];
    if (!masterSheetName)
      throw new Error(`Master Assortment: ไม่พบ sheet — SheetNames: [${masterWb.SheetNames.join(", ") || "ว่าง"}]`);
    const masterWs = masterWb.Sheets[masterSheetName];
    if (!masterWs)
      throw new Error(`Master Assortment: ไม่พบ sheet "${masterSheetName}" — SheetNames: [${masterWb.SheetNames.join(", ")}]`);

    const masterAllRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(masterWs, { defval: "" });
    const masterMap = new Map<string, MasterEntry>();

    // Detect which header key holds the barcode (case-insensitive scan of first row)
    const masterHdrKeys = Object.keys(masterAllRows[0] ?? {});
    const findHdr = (...names: string[]) =>
      names.find(n => masterHdrKeys.some(k => k.toUpperCase() === n.toUpperCase())) ??
      names.find(n => masterHdrKeys.some(k => k.toUpperCase().includes(n.toUpperCase())));
    const masterBcKey    = findHdr("BARCODE","EAN","UPC") ?? "BARCODE";
    const masterNameKey  = findHdr("DESCRIPTION","PRODUCT_NAME","LONG_DESC","NAME","SHORT_DESC","PROD_NAME") ?? "";
    const masterBarSingle= findHdr("BAR_SINGLE","BAR_SINGLE_CODE","SALE_PACK_CODE") ?? "BAR_SINGLE";
    const masterSkuPack  = findHdr("SKU_PACK","PACK_SIZE","PACK") ?? "SKU_PACK";
    const masterExtraKey = findHdr("EXTRA_INFO","EXTRA","REMARK") ?? "EXTRA_INFO";
    const masterStatusKey= findHdr("STATUS","ITEM_STATUS","ORD_STATUS") ?? "STATUS";
    const masterStoreKey = findHdr("STORE","STORE_CODE","STORE_NAME","BRANCH") ?? "STORE";

    for (const row of masterAllRows) {
      // Prefer formatted-text barcode from sheet — but sheet_to_json raw gives us the .v already.
      // For leading-zero safety: treat numeric BARCODE values with the same normalizeBarcode.
      const bc = normalizeBarcode(row[masterBcKey]);
      if (!bc) continue;
      if (!masterMap.has(bc)) {
        masterMap.set(bc, {
          name:      masterNameKey ? String(row[masterNameKey] ?? "") : "",
          barSingle: String(row[masterBarSingle] ?? ""),
          skuPack:   String(row[masterSkuPack]   ?? ""),
          extraInfo: String(row[masterExtraKey]  ?? ""),
          status:    String(row[masterStatusKey] ?? ""),
          store:     String(row[masterStoreKey]  ?? ""),
        });
      }
    }
    progress(46, `Master Assortment: ${masterMap.size.toLocaleString()} barcodes`);

    // ── 4. INDEX → map by PLANOGRAM (normalized key) ──────────────────────────
    progress(48, "อ่านไฟล์ INDEX...");

    const indexWb = XLSX.read(new Uint8Array(indexBuf), { type: "array" });
    const indexSheetName = indexWb.SheetNames.find(n => indexWb.Sheets[n]) ?? indexWb.SheetNames[0];
    if (!indexSheetName)
      throw new Error(`INDEX: ไม่พบ sheet — SheetNames: [${indexWb.SheetNames.join(", ") || "ว่าง"}]`);
    const indexWs = indexWb.Sheets[indexSheetName];
    if (!indexWs) throw new Error(`INDEX: ไม่พบ sheet "${indexSheetName}"`);

    const indexAllRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(indexWs, { defval: "" });
    const indexMap = new Map<string, IndexEntry>(); // key = normalized planogram name

    const idxHdrKeys = Object.keys(indexAllRows[0] ?? {});
    const findIdxHdr = (...names: string[]) =>
      names.find(n => idxHdrKeys.some(k => k.toUpperCase() === n.toUpperCase())) ??
      names.find(n => idxHdrKeys.some(k => k.toUpperCase().includes(n.toUpperCase())));
    const idxPlogKey  = findIdxHdr("PLANOGRAM","POG","POG_NAME","PLANOGRAM_NAME","POG NAME") ?? "PLANOGRAM";
    const idxStatKey  = findIdxHdr("STATUS","สถานะ","ITEM_STATUS") ?? "STATUS";
    const idxStoreKey = findIdxHdr("STORE","สาขา","STORE_NAME","STORE_CODE","BRANCH") ?? "STORE";
    const idxNameKey  = findIdxHdr("PLANOGRAM NAME","POG NAME","PLANOGRAM_NAME","POG_NAME","NAME") ?? "";

    for (const row of indexAllRows) {
      const plog = normalizeKey(String(row[idxPlogKey] ?? ""));
      if (!plog) continue;
      if (!indexMap.has(plog)) {
        indexMap.set(plog, {
          status:        String(row[idxStatKey]  ?? ""),
          store:         String(row[idxStoreKey] ?? ""),
          planogramName: idxNameKey ? String(row[idxNameKey] ?? "") : "",
        });
      }
      // Also store under uppercase variant for case-insensitive fallback
      const upper = plog.toUpperCase();
      if (!indexMap.has(upper)) indexMap.set(upper, indexMap.get(plog)!);
    }
    progress(55, `INDEX: ${(indexMap.size / 2) | 0} planograms`);

    // ── 5. Fixture Index → map by "SEG|POG" ──────────────────────────────────
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

    // range:1 = skip the REMARK row; actual headers (SEG / POG / Code Fixture) are in row 2
    const fixtureRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(fixtureWs, { defval: "", range: 1 });
    const fixtureMap = new Map<string, string>();
    for (const row of fixtureRows) {
      const seg  = normalizeKey(String(row["SEG"]          ?? ""));
      const pog  = normalizeKey(String(row["POG"]          ?? ""));
      const code = String(row["Code Fixture"] ?? "").trim();
      if (seg && pog && code) fixtureMap.set(`${seg}|${pog}`, code);
    }
    progress(63, `Fixture Index: ${fixtureMap.size.toLocaleString()} SEG|POG entries`);

    // ── 6. Detect column positions from template header row 6 (index 5) ───────
    progress(65, "อ่านไฟล์ Template (detect headers)...");

    const targetWb = XLSX.read(targetBuf, { type: "array", cellText: false, cellHTML: false, cellNF: false, cellDates: false });

    const s1Name = targetWb.SheetNames.find(n => n.trimStart().startsWith("New&Exsiting For Oder"));
    if (!s1Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New&Exsiting For Oder'");
    const ws1  = targetWb.Sheets[s1Name];
    const ref1 = XLSX.utils.decode_range(ws1["!ref"] ?? "A1:Z7");

    const colMap1 = new Map<string, number>();
    for (let c = 0; c <= Math.max(ref1.e.c, 30); c++) {
      const cell = ws1[XLSX.utils.encode_cell({ r: 5, c })];
      const h = cell?.v != null ? String(cell.v).replace(/\s+/g, " ").trim() : "";
      if (h) colMap1.set(h, c);
    }

    // Exact match (case-sensitive), with fallback index
    const c1 = (name: string, fb: number) => colMap1.get(name) ?? fb;
    // Fuzzy: first header whose normalised form contains the needle (case-insensitive)
    const fc1 = (needle: string): number | undefined => {
      const nl = needle.toLowerCase().replace(/\s+/g, " ");
      for (const [k, v] of colMap1)
        if (k.toLowerCase().replace(/\s+/g, " ").includes(nl)) return v;
      return undefined;
    };

    const BARCODE_COL       = c1("BARCODE", 5);
    // DIVISION ← PLANOFOLDER01 (topmost hierarchy level)
    const DIVISION_COL      = c1("DIVISION", 1);
    const PF03_COL          = colMap1.get("PLANOFOLDER03") ?? fc1("PF03") ?? 2;
    const PF04_COL          = colMap1.get("PLANOFOLDER04") ?? fc1("PF04") ?? 3;
    const PLOGNAME_COL      = fc1("PLANOGRAM NAME") ?? fc1("PLANOGRAM") ?? fc1("POG NAME");
    const NAME_COL          = fc1("Name") ?? fc1("Product Name") ?? fc1("Item Name") ?? fc1("ชื่อ");
    const SALEPACK_COL      = fc1("SALE PACK CODE") ?? fc1("SALE PACK") ?? 7;
    const PACKSIZE_COL      = fc1("Pack Size") ?? fc1("PACK SIZE") ?? 8;
    const EXTRA_COL         = fc1("Extra info") ?? fc1("EXTRA INFO") ?? fc1("EXTRA") ?? 9;
    const STATUS_COL        = fc1("Status") ?? fc1("STATUS");
    const STORE_COL         = fc1("Store") ?? fc1("STORE");
    const FIXTYPE_COL       = fc1("Fixture Type") ?? fc1("FIXTURE TYPE");
    const W_COL             = (() => {
      // Look for standalone "W" — avoid matching "STORE" or "NEW FIXTURE"
      for (const [k, v] of colMap1) if (/^\s*W\s*$/.test(k)) return v;
      return undefined;
    })();
    const H_COL             = (() => { for (const [k, v] of colMap1) if (/^\s*H\s*$/.test(k)) return v; return undefined; })();
    const D_COL             = (() => { for (const [k, v] of colMap1) if (/^\s*D\s*$/.test(k)) return v; return undefined; })();
    const NEWFIXTURE_COL    = fc1("New Fixture") ?? fc1("NEW FIXTURE");
    const NOBAY_COL         = fc1("No.Bay") ?? fc1("NO BAY") ?? fc1("NO.BAY");
    const SEQ_COL           = fc1("SEQ");
    const SHELFSTOCK_COL    = fc1("SHELF STOCK") ?? c1("SHELF STOCK FOR ORDER (Piece)", 19);
    const PCT_COL           = fc1("% Ordering") ?? fc1("%Ordering") ?? 20;
    const NETCAP_COL        = fc1("Net Capacity") ?? fc1("NET CAPACITY") ?? 21;

    const s2Name = targetWb.SheetNames.find(n => n.trim().startsWith("New for Link_IM"));
    if (!s2Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New for Link_IM'");
    const ws2  = targetWb.Sheets[s2Name];
    const ref2 = XLSX.utils.decode_range(ws2["!ref"] ?? "A1:J7");

    const colMap2 = new Map<string, number>();
    for (let c = 0; c <= Math.max(ref2.e.c, 15); c++) {
      const cell = ws2[XLSX.utils.encode_cell({ r: 5, c })];
      const h = cell?.v != null ? String(cell.v).replace(/\s+/g, " ").trim() : "";
      if (h) colMap2.set(h, c);
    }
    const fc2 = (needle: string): number | undefined => {
      const nl = needle.toLowerCase().replace(/\s+/g, " ");
      for (const [k, v] of colMap2)
        if (k.toLowerCase().replace(/\s+/g, " ").includes(nl)) return v;
      return undefined;
    };
    const BARCODE2_COL  = colMap2.get("BARCODE")    ?? fc2("BARCODE")    ?? 4;
    const DIV2_COL      = colMap2.get("DIVISION")   ?? fc2("DIVISION")   ?? 1;
    const NAME2_COL     = colMap2.get("Name")       ?? fc2("name")       ?? fc2("ชื่อ");
    const POG04_2_COL   = colMap2.get("POG 04")     ?? fc2("POG 04")     ?? fc2("PLANOFOLDER04") ?? fc2("PF04");
    const POG03_2_COL   = colMap2.get("POG 03")     ?? fc2("POG 03")     ?? fc2("PLANOFOLDER03") ?? fc2("PF03");
    const DEPT2_COL     = colMap2.get("DEPARTMENT") ?? fc2("DEPARTMENT") ?? fc2("DEPT");

    // ── 7. Build row patches ──────────────────────────────────────────────────
    progress(68, `สร้าง patches จาก ${qryRows.length.toLocaleString()} QRY rows...`);

    const DATA_START_ROW = 7;
    const s1Patches = new Map<number, Map<number, CellPatch>>();
    const s2Patches = new Map<number, Map<number, CellPatch>>();

    let matchedSpaceman = 0, matchedMaster = 0, matchedIndex = 0, matchedFixture = 0;

    for (let i = 0; i < qryRows.length; i++) {
      const qry     = qryRows[i];
      const rowNum  = DATA_START_ROW + i;
      const sm      = spacemanMap.get(qry.barcode);
      const master  = masterMap.get(qry.barcode);

      // Normalized planogram key for INDEX lookup
      const planogram    = normalizeKey(sm?.planogram ?? "");
      const idxEntry     = planogram
        ? (indexMap.get(planogram) ?? indexMap.get(planogram.toUpperCase()))
        : undefined;

      const fixtureKey  = qry.segment && planogram ? `${qry.segment}|${planogram}` : "";
      const fixtureCode = fixtureKey ? (fixtureMap.get(fixtureKey) ?? "") : "";

      if (sm)          matchedSpaceman++;
      if (master)      matchedMaster++;
      if (idxEntry)    matchedIndex++;
      if (fixtureCode) matchedFixture++;

      // ── Sheet 1 (New&Exsiting For Oder_SCM+MIS) ──────────────────────────
      const cols1 = new Map<number, CellPatch>();

      const ss = (col: number | undefined, v: string) => {
        if (col !== undefined && v !== "") cols1.set(col, { t: "s", v });
      };
      const sn = (col: number | undefined, v: unknown) => {
        if (col === undefined) return;
        const n = Number(v);
        if (!isNaN(n)) cols1.set(col, { t: "n", v: n });
      };

      // BARCODE — always written as string to preserve leading zeros
      cols1.set(BARCODE_COL, { t: "s", v: qry.barcode });

      // Name — from QRY first, fall back to Master Assortment
      const productName = qry.name || master?.name || "";
      ss(NAME_COL, productName);

      // DATA_SPACEMAN lookups
      ss(DIVISION_COL, sm?.planofolder01 ?? "");  // DIVISION = PLANOFOLDER01
      ss(PF03_COL,     sm?.planofolder03 ?? "");
      ss(PF04_COL,     sm?.planofolder04 ?? "");
      // PLANOGRAM NAME
      const plogNameVal = idxEntry?.planogramName || sm?.planogram || "";
      ss(PLOGNAME_COL, plogNameVal);

      // Master Assortment
      if (master) {
        const barSingleNum = Number(master.barSingle);
        if (!isNaN(barSingleNum) && master.barSingle !== "") {
          cols1.set(SALEPACK_COL, { t: "n", v: barSingleNum });
        } else {
          ss(SALEPACK_COL, master.barSingle);
        }
        sn(PACKSIZE_COL, master.skuPack);
        ss(EXTRA_COL,    master.extraInfo);
        // Status / Store from Master Assortment (primary), INDEX as fallback
        ss(STATUS_COL, master.status || idxEntry?.status || "");
        ss(STORE_COL,  master.store  || idxEntry?.store  || "");
      } else {
        // No Master match — fall back to INDEX for Status/Store
        ss(STATUS_COL, idxEntry?.status ?? "");
        ss(STORE_COL,  idxEntry?.store  ?? "");
      }

      // Constants
      if (FIXTYPE_COL !== undefined) cols1.set(FIXTYPE_COL, { t: "n", v: 0 });
      if (W_COL       !== undefined) cols1.set(W_COL,       { t: "n", v: 2 });
      if (H_COL       !== undefined) cols1.set(H_COL,       { t: "n", v: 1 });
      if (D_COL       !== undefined) cols1.set(D_COL,       { t: "n", v: 1 });

      // Fixture Index
      ss(NEWFIXTURE_COL, fixtureCode);

      // QRY direct fields
      ss(NOBAY_COL, qry.segment);
      ss(SEQ_COL,   qry.locationId);
      sn(SHELFSTOCK_COL, qry.totalUnits || undefined);

      // % Ordering — stored as integer (100 = 100%).
      // Config Rule keys: PLANOFOLDER01(Division) + PF03 + PF04
      const pctVal = getOrderingPct(
        exceptionConfig,
        sm?.planofolder01 ?? "",
        sm?.planofolder03 ?? "",
        sm?.planofolder04 ?? "",
      );
      cols1.set(PCT_COL, { t: "n", v: pctVal });

      // Net Capacity = SHELF_STOCK × (% Ordering / 100)
      cols1.set(NETCAP_COL, {
        t: "f",
        f: `${colLetter(SHELFSTOCK_COL)}${rowNum}*${colLetter(PCT_COL)}${rowNum}/100`,
        v: 0,
      });

      s1Patches.set(rowNum, cols1);

      // ── Sheet 2 (New for Link_IM) ─────────────────────────────────────────
      const cols2 = new Map<number, CellPatch>();
      cols2.set(BARCODE2_COL, { t: "s", v: qry.barcode });
      if (sm?.planofolder01)  cols2.set(DIV2_COL,    { t: "s", v: sm.planofolder01 });
      if (productName)        { if (NAME2_COL  !== undefined) cols2.set(NAME2_COL,  { t: "s", v: productName }); }
      if (sm?.planofolder04)  { if (POG04_2_COL !== undefined) cols2.set(POG04_2_COL, { t: "s", v: sm.planofolder04 }); }
      if (sm?.planofolder03)  { if (POG03_2_COL !== undefined) cols2.set(POG03_2_COL, { t: "s", v: sm.planofolder03 }); }
      if (sm?.planofolder02)  { if (DEPT2_COL   !== undefined) cols2.set(DEPT2_COL,   { t: "s", v: sm.planofolder02 }); }
      s2Patches.set(rowNum, cols2);
    }

    // ── 8. ZIP-patch template ─────────────────────────────────────────────────
    progress(75, "เปิด Template ZIP...");

    const files    = unzipSync(new Uint8Array(targetBuf));
    const wbXml    = strFromU8(files["xl/workbook.xml"]);
    const relsXml  = strFromU8(files["xl/_rels/workbook.xml.rels"]);
    const sstPath  = "xl/sharedStrings.xml";
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
          : buildSST([...sstStrings, ...allNewStrings]),
      );
    }

    progress(96, "บีบอัดไฟล์ผลลัพธ์...");
    const zipped = zipSync(files);
    const output = zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength);
    progress(100, "เสร็จสิ้น!");

    ctx.postMessage(
      {
        type: "done",
        buffer: output,
        stats: { total: qryRows.length, matchedSpaceman, matchedMaster, matchedIndex, matchedFixture },
      },
      [output],
    );

  } catch (err) {
    ctx.postMessage({ type: "error", message: String(err) });
  }
});
