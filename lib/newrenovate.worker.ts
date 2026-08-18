/**
 * NewRenovate worker — ZIP-patch approach (preserves all template formatting).
 *
 * Data flow (QRY-driven):
 *   QRY_Product_by_POG_by_Position  ← primary source of rows
 *     │  BARCODE → DATA_SPACEMAN    → DIVISION(PF01|PF02) / PF03 / PF04 / PLANOGRAM
 *     │  BARCODE → Master Assortment→ SALE PACK CODE / Pack Size / Extra info / Status / Store / Name
 *     │  PLANOGRAM + SEGMENT → Fixture Index → New Fixture (Code Fixture)
 *     │  PLANOGRAM → INDEX   → Status(fallback) / Store(fallback) / PLANOGRAM NAME
 *     │  QRY itself          → No.Bay / SEQ / SHELF STOCK / Name
 *     └─ Config Rules        → % Ordering (decimal 1.0=100%; stored with column % format)
 *        Net Capacity = SHELF STOCK × % Ordering
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
  fixtureRows: Record<string, string>[];  // pre-parsed rows from Fixture Index tab
  exceptionConfig: ExceptionConfig[];
};

interface SpacemanEntry {
  planofolder01: string;
  planofolder02: string;
  planofolder03: string;
  planofolder04: string;
  planofolder05: string;
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
  existingStores: string[];
  newStores:      string[];
  deleteStores:   string[];
  asIsStores:     string[]; // existingStores + deleteStores — stores with AS IS value
  toBeStores:     string[]; // existingStores + newStores   — stores with TO BE value
}

type GroupEntry = {
  barcode:     string;
  storeVal:    string;
  qry:         QrySourceRow;
  sm:          SpacemanEntry | undefined;
  master:      MasterEntry | undefined;
  productName: string;
  divisionVal: string;
  fixtureCode: string;
};

interface QrySourceRow {
  barcode:    string; // display form (preserves leading zeros)
  matchKey:   string; // stripped leading zeros — used for all map lookups
  planogram:  string; // PLANOGRAM column from QRY — primary key for INDEX lookup
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
 * Normalize barcode for DISPLAY — preserves leading zeros that come from
 * cell.w (Excel formatted text) or string-type cells.
 */
function normalizeBarcode(val: unknown, formatted?: string): string {
  if (formatted != null && formatted !== "") {
    const w = formatted.replace(/[, ]/g, "").trim();
    if (w) return w;
  }
  if (val == null || val === "") return "";
  const s = String(val).trim();
  if (!s) return "";
  if (s.startsWith("0")) return s; // preserve leading zeros from text cells
  const n = Number(s);
  if (!isNaN(n) && n > 0) return String(Math.round(n));
  return s;
}

/**
 * Return a match key for map lookups by stripping leading zeros.
 * "0012345678901" and "12345678901" map to the same key "12345678901".
 * This prevents mismatches when one file stores barcodes as numbers (no zeros)
 * and another stores as text (with leading zeros).
 */
function barcodeMatchKey(bc: string): string {
  const stripped = bc.replace(/^0+/, "");
  return stripped || bc; // don't return "" if bc was all zeros
}

/**
 * Returns ordering percentage as a decimal (1.0 = 100%, 0.8 = 80%).
 * Config Rule keys: PLANOFOLDER01(Division) + PF03 + PF04.
 * Default = 1.0 (100%).
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
    if (catOk && subOk && dscOk) return Number(rule.percentage) / 100;
  }
  return 1.0; // default 100%
}

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

/**
 * Parse the <cols> section of a sheet XML and return a map of
 * 0-indexed column → style attribute string.
 * Used to apply column default styles to newly inserted cells so that
 * number formats (e.g., percentage) are preserved.
 */
function parseColStyles(sheetXml: string): Map<number, string> {
  const colStyles = new Map<number, string>();
  const colsMatch = /<cols\b[^>]*>([\s\S]*?)<\/cols>/.exec(sheetXml);
  if (!colsMatch) return colStyles;
  const colRe = /<col\b([^>]*?)(?:\/>|><\/col>)/g;
  let m: RegExpExecArray | null;
  while ((m = colRe.exec(colsMatch[1])) !== null) {
    const attrs = m[1];
    const minM  = /\bmin="(\d+)"/.exec(attrs);
    const maxM  = /\bmax="(\d+)"/.exec(attrs);
    const styleM = /\bstyle="(\d+)"/.exec(attrs);
    if (!minM || !maxM || !styleM) continue;
    const lo = parseInt(minM[1]) - 1; // 1-indexed → 0-indexed
    const hi = parseInt(maxM[1]) - 1;
    for (let c = lo; c <= hi; c++) colStyles.set(c, styleM[1]);
  }
  return colStyles;
}

// ─── Cell patching ────────────────────────────────────────────────────────────

function buildCellXml(
  ref: string,
  patch: CellPatch,
  getSsIdx: (v: string) => number,
  sAttr: string,
): string {
  if (patch.t === "s") return `<c r="${ref}"${sAttr} t="s"><v>${getSsIdx(patch.v)}</v></c>`;
  if (patch.t === "n") return `<c r="${ref}"${sAttr}><v>${patch.v}</v></c>`;
  return `<c r="${ref}"${sAttr}><f>${patch.f}</f><v>${patch.v}</v></c>`;
}

function patchCellInRow(
  inner: string,
  ref: string,
  ci: number,
  patch: CellPatch,
  getSsIdx: (v: string) => number,
): string {
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
  colStyles: Map<number, string>,          // column default styles from <cols>
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
    for (const [ci, patch] of [...patches.entries()].sort((a, b) => a[0] - b[0])) {
      const sAttr = colStyles.has(ci) ? ` s="${colStyles.get(ci)}"` : "";
      cells += buildCellXml(`${colLetter(ci)}${rm[1]}`, patch, getSsIdx, sAttr);
    }
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
      for (const [ci, patch] of [...patches.entries()].sort((a, b) => a[0] - b[0])) {
        // Apply the column's default style so number formats (%, date, etc.) are preserved
        const sAttr = colStyles.has(ci) ? ` s="${colStyles.get(ci)}"` : "";
        cells += buildCellXml(`${colLetter(ci)}${rowNum}`, patch, getSsIdx, sAttr);
      }
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

  try {
    const { targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf, fixtureRows, exceptionConfig } = e.data;
    // ── 1. QRY → ordered row list ─────────────────────────────────────────────
    progress(3, "อ่านไฟล์ QRY_Product_by_POG_by_Position...");

    const qryWb = XLSX.read(new Uint8Array(qryBuf), { type: "array", cellText: true });
    const qryWs = qryWb.Sheets[qryWb.SheetNames[0]];
    if (!qryWs) throw new Error("ไม่พบ sheet ใน QRY_Product_by_POG_by_Position");

    const qryRange = XLSX.utils.decode_range(qryWs["!ref"] || "A1");

    // Auto-detect header row: scan first 5 rows for a row containing BARCODE/UPC/EAN
    let qHdrRow = 0;
    outer: for (let r = 0; r <= Math.min(4, qryRange.e.r); r++) {
      for (let c = 0; c <= qryRange.e.c; c++) {
        const v = String(qryWs[XLSX.utils.encode_cell({ r, c })]?.v ?? "").toUpperCase().trim();
        if (v === "BARCODE" || v === "UPC" || v === "EAN") { qHdrRow = r; break outer; }
      }
    }

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
    const qPlogCol    = findQCol("PLANOGRAM", "POG", "POG_NAME", "PLANOGRAM_NAME");
    const qSegCol     = findQCol("SEGMENT", "Segment", "segment");
    const qLocCol     = findQCol("LOCATION_ID", "LOCATIONID", "Location_ID", "location_id");
    const qUnitsCol   = findQCol("TOTAL_UNITS", "TOTALUNITS", "Total_Units", "QTY", "UNITS", "total_units");
    const qNameCol    = findQCol("NAME", "PRODUCT_NAME", "DESCRIPTION", "LONG_DESC", "SHORT_DESC");

    if (qBarcodeCol < 0)
      throw new Error(`QRY: ไม่พบ column BARCODE — headers พบ: [${qHdrs.filter(Boolean).join(", ")}]`);

    const qryRows: QrySourceRow[] = [];
    for (let r = qHdrRow + 1; r <= qryRange.e.r; r++) {
      const barcodeCell = qryWs[XLSX.utils.encode_cell({ r, c: qBarcodeCol })];
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
        matchKey:   barcodeMatchKey(bc),   // stripped for all map lookups
        planogram:  str(qPlogCol),         // PLANOGRAM from QRY — primary for INDEX lookup
        segment:    str(qSegCol),
        locationId: str(qLocCol),
        totalUnits: numStr(qUnitsCol),
        name:       str(qNameCol),
      });
    }
    progress(12, `QRY: ${qryRows.length.toLocaleString()} rows`);

    // ── 2. DATA_SPACEMAN → map by UPC matchKey ───────────────────────────────
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
      sHdrs.push(cell?.v != null ? String(cell.v).replace(/\s+/g, "").toUpperCase() : "");
    }
    // Case-insensitive + whitespace-stripped exact match
    const sIdx = (name: string) => sHdrs.indexOf(name.replace(/\s+/g, "").toUpperCase());
    const upcIdx   = sIdx("UPC");
    const pf01Idx  = sIdx("PLANOFOLDER01");
    const pf02Idx  = sIdx("PLANOFOLDER02");
    const pf03Idx  = sIdx("PLANOFOLDER03");
    const pf04Idx  = sIdx("PLANOFOLDER04");
    const pf05Idx  = sIdx("PLANOFOLDER05");
    const plogIdx  = sIdx("PLANOGRAM") >= 0 ? sIdx("PLANOGRAM") : 3;
    const catIdx   = sIdx("CATEGORY");
    const subIdx   = sIdx("SUBCATEGORY");
    const descCIdx = sIdx("DESC_C");
    progress(14, `DATA_SPACEMAN headers: UPC=${upcIdx} PF02=${pf02Idx} PF04=${pf04Idx} PF05=${pf05Idx}`);

    const getS = (r: number, c: number, useW = false): string => {
      const cell = spacemanWs[XLSX.utils.encode_cell({ r, c })];
      if (!cell) return "";
      if (useW && cell.w != null) return String(cell.w).replace(/,/g, "").trim();
      return cell.v != null ? String(cell.v).trim() : "";
    };

    const spacemanMap = new Map<string, SpacemanEntry>(); // key = barcodeMatchKey
    for (let r = 1; r <= sRange.e.r; r++) {
      const rawUpc = upcIdx >= 0 ? normalizeBarcode(getS(r, upcIdx), getS(r, upcIdx, true)) : "";
      const key    = barcodeMatchKey(rawUpc);
      if (!key) continue;
      if (!spacemanMap.has(key)) {
        spacemanMap.set(key, {
          planofolder01: pf01Idx >= 0 ? getS(r, pf01Idx) : "",
          planofolder02: pf02Idx >= 0 ? getS(r, pf02Idx) : "",
          planofolder03: pf03Idx >= 0 ? getS(r, pf03Idx) : "",
          planofolder04: pf04Idx >= 0 ? getS(r, pf04Idx) : "",
          planofolder05: pf05Idx >= 0 ? getS(r, pf05Idx) : "",
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

    // ── 3. Master Assortment → map by BARCODE matchKey ───────────────────────
    progress(34, "อ่านไฟล์ Master Assortment Orderable...");

    // Read Master Assortment directly from XLSX ZIP — bypasses SheetJS which silently
    // fails to parse very large sheets (166k+ rows) leaving Sheets["Sheet1"] = undefined.
    const mNorm = (s: string) => s.toUpperCase().replace(/[^A-Z0-9]/g, "");

    // Helper: decode XML entities
    const mDecode = (s: string) =>
      s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
       .replace(/&quot;/g, '"').replace(/&apos;/g, "'");

    // Helper: Excel column letters → 0-based index  (A→0, B→1, Z→25, AA→26, …)
    const mColIdx = (col: string): number => {
      let n = 0;
      for (let i = 0; i < col.length; i++) n = n * 26 + col.charCodeAt(i) - 64;
      return n - 1;
    };

    const mZip = unzipSync(new Uint8Array(masterBuf));
    const mZipKeys = Object.keys(mZip);

    // Lookup ZIP key: case-insensitive + leading-slash-insensitive.
    // Returns the ACTUAL key (with whatever casing/slashes the ZIP uses),
    // so mZip[mFindKey(path)] always works correctly.
    const norm = (s: string) => s.replace(/^[/\\]+/, "").toLowerCase();
    const mFindKey = (path: string): string => {
      const n = norm(path);
      return mZipKeys.find(k => norm(k) === n) ?? "";
    };

    // ── resolve sheet name → XML file path ───────────────────────────────────
    const mWbKey    = mFindKey("xl/workbook.xml");
    const mRelsKey  = mFindKey("xl/_rels/workbook.xml.rels");

    if (!mWbKey)
      throw new Error(
        `Master Assortment: ไม่พบ xl/workbook.xml ใน ZIP — paths ที่พบ: [${mZipKeys.slice(0, 20).join(", ")}]`
      );

    const mWbXml = strFromU8(mZip[mWbKey]);

    // workbook.xml: <sheet name="Sheet1" sheetId="1" r:Id="rId1"/>  OR  <sheet ...>
    // Extract name + sheetId only — do NOT require r:Id (some generators omit/vary it)
    interface MSheetDef { name: string; sheetId: number; rId: string }
    const mSheetDefs: MSheetDef[] = [];
    for (const m of mWbXml.matchAll(/<sheet\b([^>]*?)(?:\/>|>)/g)) {
      const attrs   = m[1];
      const name    = /\bname="([^"]*)"/.exec(attrs)?.[1];
      const sheetId = parseInt(/\bsheetId="(\d+)"/.exec(attrs)?.[1] ?? "0", 10);
      // Any namespace-prefixed :Id attr (r:Id, rel:Id…) — \b ensures we skip "sheetId"
      const rId     = /\b\w+:[Ii]d="([^"]*)"/.exec(attrs)?.[1] ?? "";
      if (name) mSheetDefs.push({ name, sheetId: isNaN(sheetId) ? 99 : sheetId, rId });
    }
    mSheetDefs.sort((a, b) => a.sheetId - b.sheetId);

    // _rels: rId → worksheet ZIP key (used only when r:Id is present)
    const mRelsXml = mRelsKey ? strFromU8(mZip[mRelsKey]) : "";
    const mWsRelsByRId: Record<string, string> = {};
    for (const m of mRelsXml.matchAll(/<Relationship\b([^>]*?)(?:\/>|>)/g)) {
      const attrs  = m[1];
      const id     = /\bId="([^"]*)"/.exec(attrs)?.[1] ?? "";
      const target = /\bTarget="([^"]*)"/.exec(attrs)?.[1] ?? "";
      const type   = /\bType="([^"]*)"/.exec(attrs)?.[1] ?? "";
      if (id && target && type.toLowerCase().includes("worksheet")) {
        const clean = target.replace(/^\/+/, "");  // strip leading slashes e.g. /xl/...
        const raw   = clean.startsWith("xl/") ? clean : `xl/${clean}`;
        mWsRelsByRId[id] = mFindKey(raw) || raw;
      }
    }

    // Fallback list: all xl/worksheets/sheetN.xml in ZIP, sorted by N
    const mAllWsKeys = mZipKeys
      .filter(k => /xl[/\\]worksheets[/\\]sheet\d+\.xml$/i.test(k))
      .sort((a, b) => {
        const na = parseInt(a.match(/(\d+)\.xml$/i)?.[1] ?? "0");
        const nb = parseInt(b.match(/(\d+)\.xml$/i)?.[1] ?? "0");
        return na - nb;
      });

    // Map sheet name → ZIP key (primary: rId lookup; fallback: index order)
    const mSheetByName: Record<string, string> = {};
    mSheetDefs.forEach(({ name, rId }, i) => {
      const key = (rId && mWsRelsByRId[rId]) ? mWsRelsByRId[rId] : (mAllWsKeys[i] ?? "");
      if (name && key) mSheetByName[name] = key;
    });

    const mAllSheetNames = Object.keys(mSheetByName);
    if (!mAllSheetNames.length) {
      const sheetTagPos = mWbXml.indexOf("<sheet");
      throw new Error(
        `Master Assortment: ไม่พบ sheet — defs=${mSheetDefs.length} wsKeys=[${mAllWsKeys.join(",")}] ` +
        `sheetTagAt=${sheetTagPos} near="${mWbXml.slice(Math.max(0, sheetTagPos), sheetTagPos + 150)}"`
      );
    }

    // Pick data sheet: prefer Sheet1/Data/Master; skip Explanation/Legend/etc.
    const mSkipSheets  = new Set(["explanation","legend","readme","info","instructions","instruction","manual","note","notes","remark"]);
    const mPreferNames = ["Sheet1","Sheet 1","DATA","Data","MASTER","Master","SHEET1"];
    const masterSheetName = (() => {
      for (const p of mPreferNames) {
        if (mAllSheetNames.includes(p)) return p;
      }
      for (const p of mPreferNames) {
        const f = mAllSheetNames.find(n => n.toLowerCase() === p.toLowerCase());
        if (f) return f;
      }
      for (const n of mAllSheetNames) {
        if (!mSkipSheets.has(n.toLowerCase())) return n;
      }
      return mAllSheetNames[0];
    })();

    const mSheetZipKey = mSheetByName[masterSheetName] ?? "";
    const mSheetData  = mSheetZipKey ? mZip[mSheetZipKey] : undefined;
    if (!mSheetZipKey || !mSheetData)
      throw new Error(
        `Master Assortment: ไม่พบ sheet XML "${masterSheetName}" ` +
        `(key="${mSheetZipKey}" allSheets=[${mAllSheetNames.join(",")}] ` +
        `wsRelsByRId=${JSON.stringify(mWsRelsByRId)} allWsKeys=[${mAllWsKeys.join(",")}])`
      );

    // ── parse shared strings (SST is small — safe to load as string) ─────────
    const mSstKey2 = mFindKey("xl/sharedStrings.xml");
    const mSstRaw  = mSstKey2 ? strFromU8(mZip[mSstKey2]) : "";
    const mSst: string[] = [];
    for (const si of mSstRaw.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      const texts: string[] = [];
      for (const t of si[1].matchAll(/<t(?:[^>]*)>([^<]*)<\/t>/g)) texts.push(t[1]);
      mSst.push(mDecode(texts.join("")));
    }

    // ── cell parser (operates on one row's XML string) ────────────────────────
    const mGetCell = (rowXml: string): Record<number, string> => {
      const cells: Record<number, string> = {};
      for (const c of rowXml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
        const rAttr  = /\br="([A-Za-z]+)\d+"/i.exec(c[1])?.[1]?.toUpperCase();
        if (!rAttr) continue;
        const colIdx   = mColIdx(rAttr);
        const isShared = /\bt="s"/.test(c[1]);
        const isInline = /\bt="inlineStr"/.test(c[1]);
        if (isInline) {
          cells[colIdx] = mDecode(/<t[^>]*>([^<]*)<\/t>/.exec(c[2])?.[1] ?? "");
        } else {
          const val = /<v>([^<]*)<\/v>/.exec(c[2])?.[1] ?? "";
          cells[colIdx] = isShared ? (mSst[parseInt(val)] ?? "") : mDecode(val);
        }
      }
      return cells;
    };

    // ── stream-parse sheet XML in 4 MB chunks — avoids V8 string-length limit ─
    // mSheetData is the raw decompressed Uint8Array (can be 500 MB+).
    // Converting it to a JS string with strFromU8 fails silently on large files.
    // Instead we decode chunk-by-chunk, yielding one complete <row>…</row> at a time.
    function* streamRows(data: Uint8Array): Generator<{ rowNum: number; rowXml: string }> {
      const td   = new TextDecoder("utf-8");
      const CHUNK = 4 << 20; // 4 MB
      let   buf   = "";
      let   seq   = 0;
      for (let off = 0; off < data.length; off += CHUNK) {
        const isLast = off + CHUNK >= data.length;
        buf += td.decode(data.subarray(off, isLast ? data.length : off + CHUNK), { stream: !isLast });
        let pos = 0;
        while (true) {
          const rs = buf.indexOf("<row", pos);
          if (rs < 0) { buf = ""; break; }
          // Find the matching </row> — look past the opening tag first
          const tagEnd = buf.indexOf(">", rs);
          if (tagEnd < 0) { buf = buf.slice(rs); break; }
          // Self-closing <row ... /> has no </row>; real data rows always have children
          if (buf[tagEnd - 1] === "/") { pos = tagEnd + 1; continue; }
          const re = buf.indexOf("</row>", tagEnd);
          if (re < 0) { buf = buf.slice(rs); break; }
          seq++;
          const rowXml = buf.slice(rs, re + 6);
          const rn = /\br="(\d+)"/.exec(buf.slice(rs, tagEnd + 1))?.[1];
          yield { rowNum: rn ? parseInt(rn) : seq, rowXml };
          pos = re + 6;
        }
      }
    }

    // ── single-pass: detect header then accumulate masterMap ──────────────────
    let mHdrRowNum  = -1;
    let mBcCol      = -1;
    let mBarSingleCol = -1;
    let mSkuPackCol   = -1;
    let mExtraCol     = -1;
    const mHdrCols: Record<number, string> = {};
    let   mRow1Cells: Record<number, string> = {}; // row-1 backup for header fallback
    const masterMap = new Map<string, MasterEntry>();
    let   mHdrDone  = false;

    const mSetupCols = () => {
      const mNormToCol: Record<string, number> = {};
      for (const [ci, label] of Object.entries(mHdrCols)) mNormToCol[mNorm(label)] = +ci;
      const fmc = (...cands: string[]): number => {
        for (const c of cands) { const k = mNorm(c); if (k in mNormToCol) return mNormToCol[k]; }
        for (const c of cands) {
          const k = mNorm(c);
          for (const [nl, col] of Object.entries(mNormToCol))
            if (nl.includes(k) || k.includes(nl)) return col;
        }
        return -1;
      };
      mBcCol        = fmc("BARCODE","EAN","UPC");
      mBarSingleCol = fmc("BAR_SINGLE","BAR_INGREDIENT","BARSINGLE","BARINGREDIENT","SALEPACKCODE","SALE_PACK_CODE");
      mSkuPackCol   = fmc("SKU_PACK","SKUPACK","PACK_SIZE","PACKSIZE","PACK SIZE","PACK");
      mExtraCol     = fmc("EXTRA_INFO","EXTRAINFO","EXTRA INFO","EXTRA","REMARK");
    };

    for (const { rowNum, rowXml } of streamRows(mSheetData)) {
      if (!mHdrDone) {
        // Scan first 15 rows for header
        const cells = mGetCell(rowXml);
        if (rowNum === 1) mRow1Cells = { ...cells };
        const vals     = Object.values(cells);
        const hasHdrKw = vals.some(v => {
          const n = mNorm(v);
          return n === "BARCODE" || n === "EAN" || n === "UPC"
              || n === "BARSINGLE" || n === "BARINGREDIENT"
              || n === "SKUPACK"   || n === "EXTRAINFO";
        });
        if (hasHdrKw || vals.filter(v => v.trim()).length >= 5) {
          mHdrRowNum = rowNum;
          Object.assign(mHdrCols, cells);
          mSetupCols();
          mHdrDone = true;
        } else if (rowNum >= 15) {
          // No clear header found — fall back to row 1
          mHdrRowNum = 1;
          Object.assign(mHdrCols, mRow1Cells);
          mSetupCols();
          mHdrDone = true;
          if (mBcCol < 0)
            throw new Error(
              `Master Assortment: ไม่พบคอลัมน์ BARCODE — ` +
              `headers=[${Object.values(mHdrCols).join("|")}] sst=${mSst.length}`
            );
        }
        continue; // skip header rows as data
      }
      // Data rows
      if (rowNum <= mHdrRowNum) continue;
      const cells = mGetCell(rowXml);
      const bc    = normalizeBarcode(cells[mBcCol] ?? "");
      const key   = barcodeMatchKey(bc);
      if (!key) continue;
      if (!masterMap.has(key)) {
        masterMap.set(key, {
          name:      "",
          barSingle: mBarSingleCol >= 0 ? (cells[mBarSingleCol] ?? "") : "",
          skuPack:   mSkuPackCol   >= 0 ? (cells[mSkuPackCol]   ?? "") : "",
          extraInfo: mExtraCol     >= 0 ? (cells[mExtraCol]     ?? "") : "",
          status:    "",
          store:     "",
        });
      }
    }

    if (!mHdrDone || mBcCol < 0)
      throw new Error(
        `Master Assortment: ไม่พบคอลัมน์ BARCODE — ` +
        `headers=[${Object.values(mHdrCols).join("|")}] sst=${mSst.length} hdrRow=${mHdrRowNum}`
      );

    progress(46, `Master: ${masterMap.size} barcodes | sheet="${masterSheetName}" hdrRow=${mHdrRowNum} sst=${mSst.length} cols=[BC:${mBcCol} SALEPACK:${mBarSingleCol} PACK:${mSkuPackCol} EXTRA:${mExtraCol}]`);

    // ── 4. INDEX → map by PLANOGRAM (normalised key) ─────────────────────────
    // Handles INDEX BIG C mini cross-tab structure: status is derived from
    // AS IS / TO BE store-matrix columns (not a dedicated STATUS column).
    progress(48, "อ่านไฟล์ INDEX...");

    const indexWb = XLSX.read(new Uint8Array(indexBuf), { type: "array" });
    const indexSheetName = indexWb.SheetNames.find(n => indexWb.Sheets[n]) ?? indexWb.SheetNames[0];
    if (!indexSheetName)
      throw new Error(`INDEX: ไม่พบ sheet — SheetNames: [${indexWb.SheetNames.join(", ") || "ว่าง"}]`);
    const indexWs = indexWb.Sheets[indexSheetName];
    if (!indexWs) throw new Error(`INDEX: ไม่พบ sheet "${indexSheetName}"`);

    const iRange  = XLSX.utils.decode_range(indexWs["!ref"] || "A1");
    const iCell   = (r: number, c: number): string => {
      const cell = indexWs[XLSX.utils.encode_cell({ r, c })];
      return cell?.v != null ? String(cell.v).trim() : "";
    };

    // Expand merged cells so "AS IS"/"TO BE" that span multiple columns are detected
    const expandRowMerges = (row: number): Map<number, string> => {
      const map = new Map<number, string>();
      const merges: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }> =
        (indexWs["!merges"] as Array<{ s: { r: number; c: number }; e: { r: number; c: number } }> | undefined) ?? [];
      for (let c = 0; c <= iRange.e.c; c++) {
        const v = iCell(row, c);
        if (v) map.set(c, v);
      }
      for (const mg of merges) {
        if (mg.s.r <= row && row <= mg.e.r) {
          const v = iCell(mg.s.r, mg.s.c);
          if (v) for (let c = mg.s.c; c <= mg.e.c; c++) if (!map.has(c)) map.set(c, v);
        }
      }
      return map;
    };

    // Scan rows 0–30 for the "AS IS" / "TO BE" header row
    let asisToBeRow = -1;
    const asIsColumns = new Set<number>();
    const toBeColumns = new Set<number>();

    for (let r = 0; r <= Math.min(30, iRange.e.r); r++) {
      const rowMap = expandRowMerges(r);
      let foundA = false, foundB = false;
      for (const [c, v] of rowMap) {
        const up = v.toUpperCase();
        if (up === "AS IS")  { asIsColumns.add(c); foundA = true; }
        else if (up === "TO BE") { toBeColumns.add(c); foundB = true; }
      }
      if (foundA || foundB) { asisToBeRow = r; break; }
    }

    // Scan rows 0–20 for columns whose header contains "POG NAME" (case-insensitive)
    let pogHdrRow = -1;
    const pogNameCols: number[] = [];

    for (let r = 0; r <= Math.min(20, iRange.e.r); r++) {
      const found: number[] = [];
      for (let c = 0; c <= Math.min(iRange.e.c, 30); c++) {
        const v = iCell(r, c).toUpperCase();
        if (v === "POG NAME" || v === "PLANOGRAM" || v === "PLANOGRAM NAME") found.push(c);
      }
      if (found.length > 0) { pogHdrRow = r; pogNameCols.push(...found); break; }
    }

    const iDataStart = pogHdrRow >= 0 ? pogHdrRow + 1
                     : asisToBeRow >= 0 ? asisToBeRow + 8
                     : 15;

    const indexMap = new Map<string, IndexEntry>(); // key = normalizeKey(planogram)

    if (asIsColumns.size === 0 && toBeColumns.size === 0) {
      // ── Fallback: flat-table INDEX with a STATUS column ────────────────────
      const indexAllRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(indexWs, { defval: "" });
      const iHdrs = Object.keys(indexAllRows[0] ?? {});
      const findIHdr = (...names: string[]) =>
        names.find(n => iHdrs.some(k => k.toUpperCase() === n.toUpperCase())) ??
        names.find(n => iHdrs.some(k => k.toUpperCase().includes(n.toUpperCase())));
      const iPlogKey  = findIHdr("PLANOGRAM","POG","POG_NAME","PLANOGRAM_NAME","POG NAME") ?? "PLANOGRAM";
      const iStatKey  = findIHdr("STATUS","สถานะ","ITEM_STATUS") ?? "STATUS";
      const iStoreKey = findIHdr("STORE","สาขา","STORE_NAME","STORE_CODE","BRANCH") ?? "STORE";
      const iNameKey  = findIHdr("PLANOGRAM NAME","POG NAME","PLANOGRAM_NAME","POG_NAME") ?? "";
      for (const row of indexAllRows) {
        const plog = normalizeKey(String(row[iPlogKey] ?? ""));
        if (!plog) continue;
        if (!indexMap.has(plog)) {
          const _fStatus = String(row[iStatKey]  ?? "");
          const _fStore  = String(row[iStoreKey] ?? "").split(",").map(s => s.trim()).filter(Boolean);
          const _fUp     = _fStatus.toUpperCase();
          const _isExisting = _fUp === "EXISTING";
          const _isNew      = _fUp === "NEW" || _fUp === "NEW EXPAND";
          const _isDelete   = _fUp.startsWith("DELETE");
          const entry: IndexEntry = {
            status:        _fUp === "NEW" ? "NEW EXPAND" : _fStatus,
            store:         String(row[iStoreKey] ?? ""),
            planogramName: iNameKey ? String(row[iNameKey] ?? "") : plog,
            existingStores: _isExisting             ? _fStore : [],
            newStores:      _isNew                  ? _fStore : [],
            deleteStores:   _isDelete               ? _fStore : [],
            asIsStores:     _isExisting || _isDelete ? _fStore : [],
            toBeStores:     _isExisting || _isNew    ? _fStore : [],
          };
          indexMap.set(plog, entry);
          const upper = plog.toUpperCase();
          if (!indexMap.has(upper)) indexMap.set(upper, entry);
        }
      }
    } else {
      // ── Primary: INDEX BIG C mini cross-tab — derive STATUS from AS IS/TO BE ─
      const effectivePogCols = pogNameCols;

      // Pair AS IS + TO BE columns by store (sorted order matches alternating pairs)
      const sortedAsIs = [...asIsColumns].sort((a, b) => a - b);
      const sortedToBe = [...toBeColumns].sort((a, b) => a - b);
      const pairCount  = Math.min(sortedAsIs.length, sortedToBe.length);

      // Find Store Code row — label column is immediately left of first AS IS column
      const labelCol = sortedAsIs.length > 0 ? sortedAsIs[0] - 1 : -1;
      let storeCodeRow = -1;
      if (labelCol >= 0) {
        const scanFrom = asisToBeRow >= 0 ? asisToBeRow + 1 : 0;
        for (let r = scanFrom; r <= Math.min(scanFrom + 25, iRange.e.r); r++) {
          if (iCell(r, labelCol).toUpperCase().replace(/\s+/g, "") === "STORECODE") {
            storeCodeRow = r; break;
          }
        }
      }
      // Pre-read store code per pair index (each pair's AS IS col shares the store code)
      const pairStoreCodes: string[] = Array.from({ length: pairCount }, (_, i) => {
        if (storeCodeRow < 0) return "";
        const cell = indexWs[XLSX.utils.encode_cell({ r: storeCodeRow, c: sortedAsIs[i] })];
        return cell?.v != null ? String(cell.v).trim() : "";
      });

      for (let r = iDataStart; r <= iRange.e.r; r++) {
        // Collect all POG NAME values in this row (skip header text / numeric totals)
        const pogNames: string[] = [];
        for (const c of effectivePogCols) {
          const v = normalizeKey(iCell(r, c));
          if (!v) continue;
          const up = v.toUpperCase();
          if (up.includes("POG NAME") || up.startsWith("DEPARTMENT") || /^\d+$/.test(v)) continue;
          pogNames.push(v);
        }
        if (pogNames.length === 0) continue;

        // Check presence within same-store pairs — EXISTING only if BOTH AS IS + TO BE in one store
        let hasExisting = false, hasAsIs = false, hasToBe = false;
        const asIsStores = new Set<string>();
        const toBeStores = new Set<string>();
        for (let i = 0; i < pairCount; i++) {
          const aCell = indexWs[XLSX.utils.encode_cell({ r, c: sortedAsIs[i] })];
          const bCell = indexWs[XLSX.utils.encode_cell({ r, c: sortedToBe[i] })];
          const aHas = aCell?.v != null && aCell.v !== "" && aCell.v !== 0;
          const bHas = bCell?.v != null && bCell.v !== "" && bCell.v !== 0;
          if (aHas && bHas) hasExisting = true;
          else if (aHas) hasAsIs = true;
          else if (bHas) hasToBe = true;
          if (aHas && pairStoreCodes[i]) asIsStores.add(pairStoreCodes[i]);
          if (bHas && pairStoreCodes[i]) toBeStores.add(pairStoreCodes[i]);
        }

        // Per-store status: each store classified independently
        const existingStoresList = [...asIsStores].filter(s =>  toBeStores.has(s));
        const newStoresList      = [...toBeStores].filter(s => !asIsStores.has(s));
        const deleteStoresList   = [...asIsStores].filter(s => !toBeStores.has(s));

        // Dominant planogram status (for reporting / progress messages)
        let derivedStatus = "";
        if      (hasExisting) derivedStatus = "EXISTING";
        else if (hasAsIs)     derivedStatus = "DELETE";
        else if (hasToBe)     derivedStatus = "NEW EXPAND";

        // derivedStore kept for compat (EXISTING/NEW → toBeStores; DELETE → asIsStores)
        const derivedStore = derivedStatus === "DELETE"
          ? [...asIsStores].join(",")
          : [...toBeStores].join(",");

        const primaryPog = pogNames[0];
        const entry: IndexEntry = {
          status:         derivedStatus,
          store:          derivedStore,
          planogramName:  primaryPog,
          existingStores: existingStoresList,
          newStores:      newStoresList,
          deleteStores:   deleteStoresList,
          asIsStores:     [...asIsStores], // stores with AS IS value (existing + delete)
          toBeStores:     [...toBeStores], // stores with TO BE value (existing + new)
        };

        // Index by ALL POG NAME column values so any variant of planogram name matches
        for (const pog of pogNames) {
          if (!indexMap.has(pog)) {
            indexMap.set(pog, entry);
            const upper = pog.toUpperCase();
            if (!indexMap.has(upper)) indexMap.set(upper, entry);
          }
        }
      }
    }
    progress(55, `INDEX: ${Math.ceil(indexMap.size / 2)} planograms (asisRow=${asisToBeRow} AS IS cols=${asIsColumns.size} TO BE cols=${toBeColumns.size})`);

    // ── 5. Fixture Index → map by "SEG|POG" (pre-parsed rows from Fixture Index tab) ──
    progress(57, "สร้าง Fixture Index map...");

    const fixtureMap = new Map<string, string>();
    for (const row of fixtureRows) {
      const seg  = normalizeKey(String(row["SEG"]          ?? ""));
      const pog  = normalizeKey(String(row["POG"]          ?? ""));
      const code = String(row["Code Fixture"] ?? "").trim();
      if (seg && pog && code) fixtureMap.set(`${seg}|${pog}`, code);
    }
    progress(63, `Fixture Index: ${fixtureMap.size.toLocaleString()} SEG|POG entries`);

    // ── 6. Detect column positions from template header row 6 (index 5) ───────
    progress(65, "อ่านไฟล์ Template...");

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
    const c1 = (name: string, fb: number) => colMap1.get(name) ?? fb;
    const fc1 = (needle: string): number | undefined => {
      const nl = needle.toLowerCase().replace(/\s+/g, " ");
      for (const [k, v] of colMap1)
        if (k.toLowerCase().replace(/\s+/g, " ").includes(nl)) return v;
      return undefined;
    };

    const BARCODE_COL    = c1("BARCODE", 5);
    // DIVISION = PLANOFOLDER01 (topmost). Fallback: PF02 if PF01 is absent in DS.
    const DIVISION_COL   = c1("DIVISION", 1);
    const PF03_COL       = colMap1.get("DEPARTMENT") ?? fc1("DEPARTMENT") ?? fc1("DEPT") ?? 2;
    const PF04_COL       = colMap1.get("POG CATE")   ?? fc1("POG CATE") ?? fc1("POG")  ?? 3;
    const PLOGNAME_COL   = fc1("PLANOGRAM NAME") ?? fc1("POG NAME") ?? fc1("PLANOGRAM");
    // Exact case-insensitive match first — avoids false hits on "PLANOGRAM NAME" etc.
    const exactCI1 = (needle: string) => {
      const nl = needle.toLowerCase();
      for (const [k, v] of colMap1) if (k.toLowerCase() === nl) return v;
      return undefined;
    };
    const NAME_COL = exactCI1("name") ?? exactCI1("product name") ?? exactCI1("item name") ??
                     exactCI1("ชื่อสินค้า") ?? exactCI1("ชื่อ");
    const SALEPACK_COL   = fc1("SALE PACK CODE") ?? fc1("SALE PACK") ?? 7;
    const PACKSIZE_COL   = fc1("Pack Size") ?? fc1("PACK SIZE") ?? 8;
    const EXTRA_COL      = fc1("Extra info") ?? fc1("EXTRA INFO") ?? fc1("EXTRA") ?? 9;
    const STATUS_COL     = fc1("Status") ?? fc1("STATUS");
    const STORE_COL      = fc1("Store") ?? fc1("STORE");
    const FIXTYPE_COL    = fc1("Fixture Type") ?? fc1("FIXTURE TYPE");
    const W_COL          = (() => { for (const [k, v] of colMap1) if (/^\s*W\s*$/.test(k)) return v; return undefined; })();
    const H_COL          = (() => { for (const [k, v] of colMap1) if (/^\s*H\s*$/.test(k)) return v; return undefined; })();
    const D_COL          = (() => { for (const [k, v] of colMap1) if (/^\s*D\s*$/.test(k)) return v; return undefined; })();
    const NEWFIXTURE_COL = fc1("New Fixture") ?? fc1("NEW FIXTURE");
    const NOBAY_COL      = fc1("No.Bay") ?? fc1("NO BAY") ?? fc1("NO.BAY");
    const SEQ_COL        = fc1("SEQ");
    const SHELFSTOCK_COL = fc1("SHELF STOCK") ?? c1("SHELF STOCK FOR ORDER (Piece)", 19);
    const PCT_COL        = fc1("% Ordering") ?? fc1("%Ordering") ?? 20;
    const NETCAP_COL     = fc1("Net Capacity") ?? fc1("NET CAPACITY") ?? 21;

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
    const BARCODE2_COL     = colMap2.get("BARCODE")    ?? fc2("BARCODE")    ?? 4;
    const DIV2_COL         = colMap2.get("DIVISION")   ?? fc2("DIVISION")   ?? 1;
    const NAME2_COL        = colMap2.get("Name")       ?? fc2("name")       ?? fc2("ชื่อ");
    const POG04_2_COL      = colMap2.get("POG CATE") ?? fc2("POG CATE") ?? fc2("POG");
    const POG03_2_COL      = POG04_2_COL; // same column — POG CATE is a single column in this sheet
    const DEPT2_COL        = colMap2.get("DEPARTMENT") ?? fc2("DEPARTMENT") ?? fc2("DEPT");
    const STATUS2_COL      = colMap2.get("STATUS")     ?? fc2("status");
    const STORECOUNT2_COL  = fc2("number store")       ?? fc2("store");

    // ── Delete for Exception_ IM sheet ───────────────────────────────────────
    const s3Name = targetWb.SheetNames.find(n => /delete.*exception/i.test(n.trim()));
    let DIV3_COL = 1, DEPT3_COL = 2, POG3_COL = 3, BARCODE3_COL = 4;
    let NAME3_COL = 5, STATUS3_COL = 6, STORE3_COL = 7;
    if (s3Name) {
      const ws3  = targetWb.Sheets[s3Name];
      const ref3 = XLSX.utils.decode_range(ws3?.["!ref"] ?? "A1:I7");
      const colMap3 = new Map<string, number>();
      for (let c = 0; c <= Math.max(ref3.e.c, 15); c++) {
        const cell = ws3[XLSX.utils.encode_cell({ r: 5, c })];
        const h = cell?.v != null ? String(cell.v).replace(/\s+/g, " ").trim() : "";
        if (h) colMap3.set(h, c);
      }
      const fc3 = (needle: string): number | undefined => {
        const nl = needle.toLowerCase().replace(/\s+/g, " ");
        for (const [k, v] of colMap3)
          if (k.toLowerCase().replace(/\s+/g, " ").includes(nl)) return v;
        return undefined;
      };
      DIV3_COL     = colMap3.get("DIVISION")   ?? fc3("division")   ?? 1;
      DEPT3_COL    = colMap3.get("DEPARTMENT") ?? fc3("department") ?? 2;
      POG3_COL     = fc3("pog cate") ?? fc3("pog") ?? 3;
      BARCODE3_COL = colMap3.get("BARCODE")    ?? fc3("barcode")    ?? 4;
      NAME3_COL    = fc3("name") ?? fc3("ชื่อ") ?? 5;
      STATUS3_COL  = colMap3.get("STATUS")     ?? fc3("status")     ?? 6;
      STORE3_COL   = fc3("store exception") ?? fc3("store") ?? 7;
    }

    // ── 7. Build row patches ──────────────────────────────────────────────────
    progress(68, `สร้าง patches จาก ${qryRows.length.toLocaleString()} QRY rows...`);

    const DATA_START_ROW   = 7; // header at row 6 (idx 5), blank row 7, data row 8 (Sheet 1 & 2)
    const S3_DATA_START_ROW = 7; // Sheet 3: header at Excel row 6, data starts at Excel row 7
    const s1Patches = new Map<number, Map<number, CellPatch>>();
    const s2Patches = new Map<number, Map<number, CellPatch>>();
    const s3Patches = new Map<number, Map<number, CellPatch>>();

    // All three sheets staged first, then sorted before row numbers are assigned
    type S1Row = { storeVal: string; planogramName: string; noBay: string; seq: string; cols: Map<number, CellPatch> };
    type StoreRow = { storeVal: string; cols: Map<number, CellPatch> };
    const s1Staging: S1Row[]    = [];
    const s2Staging: StoreRow[] = [];
    const s3Staging: StoreRow[] = [];

    let matchedSpaceman = 0, matchedMaster = 0, matchedIndex = 0, matchedFixture = 0;

    // ── Phase 1: Build Group A (AS IS stores) and Group B (TO BE stores) ─────────
    // Each group is a flat list of (barcode, store, qryRow, lookupData) tuples.
    // Group A = stores with AS IS value for this planogram (existingStores + deleteStores).
    // Group B = stores with TO BE value for this planogram (existingStores + newStores).
    const groupA: GroupEntry[] = [];
    const groupB: GroupEntry[] = [];
    const unmatchedPlanograms = new Set<string>(); // QRY planograms not found in INDEX

    for (let i = 0; i < qryRows.length; i++) {
      const qry = qryRows[i];

      const sm     = spacemanMap.get(qry.matchKey);
      const master = masterMap.get(qry.matchKey);

      const planogram = normalizeKey(qry.planogram);
      const idxEntry  = planogram
        ? (indexMap.get(planogram) ?? indexMap.get(planogram.toUpperCase()))
        : undefined;

      if (!idxEntry) { if (qry.planogram) unmatchedPlanograms.add(qry.planogram); continue; }

      const fixtureKey  = qry.segment && planogram ? `${qry.segment}|${planogram}` : "";
      const fixtureCode = fixtureKey ? (fixtureMap.get(fixtureKey) ?? "") : "";
      const productName = qry.name || master?.name || "";
      const divisionVal = sm?.planofolder02 || "";

      if (sm)          matchedSpaceman++;
      if (master)      matchedMaster++;
      matchedIndex++;
      if (fixtureCode) matchedFixture++;

      for (const storeVal of idxEntry.asIsStores) {
        groupA.push({ barcode: qry.barcode, storeVal, qry, sm, master, productName, divisionVal, fixtureCode });
      }
      for (const storeVal of idxEntry.toBeStores) {
        groupB.push({ barcode: qry.barcode, storeVal, qry, sm, master, productName, divisionVal, fixtureCode });
      }
    }

    // ── Phase 2: Build cross-group lookup sets — key = barcode only ─────────────
    // Status is determined per barcode (across all stores/planograms).
    // The store in each group entry is used only when filling the report rows.
    const setA = new Set(groupA.map(e => e.barcode));
    const setB = new Set(groupB.map(e => e.barcode));

    // ── Build Preview Data ────────────────────────────────────────────────────
    type PreviewRow = { barcode: string; planograms: string[]; storesA: string[]; storesB: string[]; status: "EXISTING" | "NEW EXPAND" | "DELETE" };
    const prevMap = new Map<string, { planograms: Set<string>; storesA: Set<string>; storesB: Set<string> }>();
    const ensurePrev = (bc: string) => {
      if (!prevMap.has(bc)) prevMap.set(bc, { planograms: new Set(), storesA: new Set(), storesB: new Set() });
      return prevMap.get(bc)!;
    };
    for (const e of groupA) { const r = ensurePrev(e.barcode); r.planograms.add(e.qry.planogram); r.storesA.add(e.storeVal); }
    for (const e of groupB) { const r = ensurePrev(e.barcode); r.planograms.add(e.qry.planogram); r.storesB.add(e.storeVal); }
    const nsort = (a: string, b: string): number => { const na = +a, nb = +b; return isFinite(na) && isFinite(nb) ? na - nb : a.localeCompare(b); };
    const previewRows: PreviewRow[] = [...prevMap.entries()].map(([barcode, r]) => {
      const inA = r.storesA.size > 0, inB = r.storesB.size > 0;
      return {
        barcode,
        planograms: [...r.planograms].sort(),
        storesA:    [...r.storesA].sort(nsort),
        storesB:    [...r.storesB].sort(nsort),
        status:     (inA && inB) ? "EXISTING" : inA ? "DELETE" : "NEW EXPAND",
      };
    });

    // ── Phase 3: Process Group B → Sheet 1 (EXISTING + NEW EXPAND) & Sheet 2 (NEW EXPAND) ──
    for (const bEntry of groupB) {
      const rowStatus = setA.has(bEntry.barcode) ? "EXISTING" : "NEW EXPAND";

      const cols1 = new Map<number, CellPatch>();
      const ss = (col: number | undefined, v: string) => {
        if (col !== undefined && v !== "") cols1.set(col, { t: "s", v });
      };
      const sn = (col: number | undefined, v: unknown) => {
        if (col === undefined) return;
        const n = Number(v);
        if (!isNaN(n)) cols1.set(col, { t: "n", v: n });
      };

      cols1.set(BARCODE_COL, { t: "s", v: bEntry.barcode });
      ss(NAME_COL,      bEntry.productName);
      ss(DIVISION_COL,  bEntry.divisionVal);
      ss(PF03_COL,      bEntry.sm?.planofolder04 ?? "");
      ss(PF04_COL,      bEntry.sm?.planofolder05 ?? "");
      // PLANOGRAM NAME = QRY planogram name (from Group B's row), per requirement
      ss(PLOGNAME_COL,  bEntry.qry.planogram);

      if (bEntry.master) {
        ss(SALEPACK_COL, bEntry.master.barSingle);
        sn(PACKSIZE_COL, bEntry.master.skuPack);
        ss(EXTRA_COL,    bEntry.master.extraInfo);
      }

      ss(STATUS_COL, rowStatus);

      if (FIXTYPE_COL !== undefined) cols1.set(FIXTYPE_COL, { t: "n", v: 0 });
      if (W_COL       !== undefined) cols1.set(W_COL,       { t: "n", v: 2 });
      if (H_COL       !== undefined) cols1.set(H_COL,       { t: "n", v: 1 });
      if (D_COL       !== undefined) cols1.set(D_COL,       { t: "n", v: 1 });

      ss(NEWFIXTURE_COL, bEntry.fixtureCode);
      ss(NOBAY_COL,      bEntry.qry.segment);
      ss(SEQ_COL,        bEntry.qry.locationId);
      sn(SHELFSTOCK_COL, bEntry.qry.totalUnits || undefined);

      const pctVal  = getOrderingPct(
        exceptionConfig,
        bEntry.sm?.planofolder01 ?? "",
        bEntry.sm?.planofolder03 ?? "",
        bEntry.sm?.planofolder04 ?? "",
      );
      const pctText = Math.round(pctVal * 100) + "%";
      cols1.set(PCT_COL, { t: "s", v: pctText });

      const shelfStock = Number(bEntry.qry.totalUnits) || 0;
      const netCapVal  = Math.round(shelfStock * pctVal);
      if (NETCAP_COL !== undefined && netCapVal > 0)
        cols1.set(NETCAP_COL, { t: "n", v: netCapVal });

      if (STORE_COL !== undefined && bEntry.storeVal)
        cols1.set(STORE_COL, { t: "s", v: bEntry.storeVal });

      s1Staging.push({
        storeVal:      bEntry.storeVal,
        planogramName: bEntry.qry.planogram,
        noBay:         bEntry.qry.segment,
        seq:           bEntry.qry.locationId,
        cols:          cols1,
      });

      // Sheet 2: NEW EXPAND only
      if (rowStatus === "NEW EXPAND") {
        const cols2 = new Map<number, CellPatch>();
        const ss2 = (col: number | undefined, v: string) => { if (col !== undefined && v) cols2.set(col, { t: "s", v }); };
        ss2(BARCODE2_COL, bEntry.barcode);
        ss2(DIV2_COL,     bEntry.divisionVal);
        ss2(NAME2_COL,    bEntry.productName);
        ss2(POG04_2_COL,  bEntry.sm?.planofolder05 ?? "");
        ss2(POG03_2_COL,  bEntry.sm?.planofolder05 ?? "");
        ss2(DEPT2_COL,    bEntry.sm?.planofolder04 ?? "");
        ss2(STATUS2_COL,  rowStatus);
        if (STORECOUNT2_COL !== undefined && bEntry.storeVal)
          cols2.set(STORECOUNT2_COL, { t: "s", v: bEntry.storeVal });
        s2Staging.push({ storeVal: bEntry.storeVal, cols: cols2 });
      }
    }

    // ── Phase 4: Process Group A → Sheet 3 (DELETE: barcode in A but not in B) ──
    if (s3Name) {
      for (const aEntry of groupA) {
        if (setB.has(aEntry.barcode)) continue; // barcode also in B → EXISTING, handled in Sheet 1

        const cols3 = new Map<number, CellPatch>();
        const ss3 = (col: number, v: string) => { if (v) cols3.set(col, { t: "s", v }); };
        ss3(DIV3_COL,     aEntry.sm?.planofolder02 || "");
        ss3(DEPT3_COL,    aEntry.sm?.planofolder04 || "");
        ss3(POG3_COL,     aEntry.sm?.planofolder05 || "");
        ss3(BARCODE3_COL, aEntry.barcode);
        ss3(NAME3_COL,    aEntry.productName);
        ss3(STATUS3_COL,  "DELETE");
        ss3(STORE3_COL,   aEntry.storeVal);
        s3Staging.push({ storeVal: aEntry.storeVal, cols: cols3 });
      }
    }

    // ── Sort all staging arrays then assign row numbers ───────────────────────
    const cmpNum = (a: string, b: string): number => {
      const na = parseFloat(a), nb = parseFloat(b);
      return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b);
    };
    const cmpStr = (a: string, b: string) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

    // Sheet 1: priority 1=store 2=planogramName 3=noBay 4=seq
    s1Staging.sort((a, b) =>
      cmpNum(a.storeVal,      b.storeVal)      ||
      cmpStr(a.planogramName, b.planogramName) ||
      cmpNum(a.noBay,         b.noBay)         ||
      cmpNum(a.seq,           b.seq)
    );
    s1Staging.forEach(({ cols }, i) => s1Patches.set(DATA_START_ROW + i, cols));

    // Sheet 2 & 3: sort by store code only
    const sortByStore = (arr: StoreRow[]) =>
      arr.sort((a, b) => cmpNum(a.storeVal, b.storeVal));

    sortByStore(s2Staging).forEach(({ cols }, i) => s2Patches.set(DATA_START_ROW + i, cols));
    sortByStore(s3Staging).forEach(({ cols }, i) => s3Patches.set(S3_DATA_START_ROW + i, cols));

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
    const sheet1Xml  = strFromU8(files[path1]);
    const colStyles1 = parseColStyles(sheet1Xml);  // extract column % / date formats etc.
    progress(80, `Patch Sheet 1 (${s1Patches.size.toLocaleString()} rows)...`);
    const r1 = patchSheetXml(sheet1Xml, [...sstStrings, ...allNewStrings], s1Patches, colStyles1);
    files[path1] = strToU8(r1.sheetXml);
    allNewStrings.push(...r1.newStrings);

    const path2 = findSheetPath(wbXml, relsXml, s2Name);
    if (!path2 || !files[path2]) throw new Error(`ไม่พบ ZIP path สำหรับ sheet "${s2Name}"`);
    const sheet2Xml  = strFromU8(files[path2]);
    const colStyles2 = parseColStyles(sheet2Xml);
    progress(90, `Patch Sheet 2 (${s2Patches.size.toLocaleString()} rows)...`);
    const r2 = patchSheetXml(sheet2Xml, [...sstStrings, ...allNewStrings], s2Patches, colStyles2);
    files[path2] = strToU8(r2.sheetXml);
    allNewStrings.push(...r2.newStrings);

    // Delete for Exception_ IM sheet
    if (s3Name && s3Patches.size > 0) {
      const path3 = findSheetPath(wbXml, relsXml, s3Name);
      if (path3 && files[path3]) {
        const sheet3Xml  = strFromU8(files[path3]);
        const colStyles3 = parseColStyles(sheet3Xml);
        progress(92, `Patch Delete sheet (${s3Patches.size.toLocaleString()} rows)...`);
        const r3 = patchSheetXml(sheet3Xml, [...sstStrings, ...allNewStrings], s3Patches, colStyles3);
        files[path3] = strToU8(r3.sheetXml);
        allNewStrings.push(...r3.newStrings);
      }
    }

    if (allNewStrings.length > 0) {
      files[sstPath] = strToU8(
        files[sstPath]
          ? appendSST(strFromU8(files[sstPath]), allNewStrings)
          : buildSST([...sstStrings, ...allNewStrings]),
      );
    }

    // Debug: show sample keys from masterMap vs QRY to diagnose mismatch
    const sampleMasterKeys = [...masterMap.keys()].slice(0, 3).join(" | ");
    const sampleQryKeys    = qryRows.slice(0, 3).map(r => r.matchKey).join(" | ");
    progress(94, `[DEBUG] masterMap size=${masterMap.size} | sample keys: [${sampleMasterKeys || "ว่าง"}] vs QRY: [${sampleQryKeys || "ว่าง"}]`);

    progress(96, "บีบอัดไฟล์ผลลัพธ์...");
    const zipped = zipSync(files);
    const output = zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength);
    progress(100, "เสร็จสิ้น!");

    ctx.postMessage(
      {
        type: "done",
        buffer: output,
        stats: { total: qryRows.length, matchedSpaceman, matchedMaster, matchedIndex, matchedFixture, masterMapSize: masterMap.size, masterSheetName },
        preview: { rows: previewRows, unmatchedPlanograms: [...unmatchedPlanograms].sort() },
      },
      [output],
    );

  } catch (err) {
    ctx.postMessage({ type: "error", message: String(err) });
  }
});
