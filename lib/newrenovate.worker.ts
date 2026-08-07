import * as XLSX from "xlsx";
import type { ExceptionConfig } from "./types";

// Worker postMessage with transfer support
const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

// ─── Message types ─────────────────────────────────────────────────────────────

type InMsg = {
  type: "run";
  targetBuf:   ArrayBuffer;
  spacemanBuf: ArrayBuffer;
  indexBuf:    ArrayBuffer;
  qryBuf:      ArrayBuffer;
  exceptionConfig: ExceptionConfig[];
};

type OutMsg =
  | { type: "progress"; pct: number; msg: string }
  | { type: "done"; buffer: ArrayBuffer; stats: Stats }
  | { type: "error"; message: string };

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

// ─── Helpers ───────────────────────────────────────────────────────────────────

function progress(pct: number, msg: string) {
  ctx.postMessage({ type: "progress", pct, msg } satisfies OutMsg);
}

function normalizeBarcode(val: unknown): string {
  if (val === null || val === undefined || val === "") return "";
  const s = String(val).trim();
  if (!s) return "";
  const n = Number(s);
  if (!isNaN(n) && n > 0) return String(Math.round(n));
  return s;
}

function cellStr(ws: XLSX.WorkSheet, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  return cell?.v != null ? String(cell.v).trim() : "";
}

function detectCols(ws: XLSX.WorkSheet, headerRowIdx: number): Map<string, number> {
  const m = new Map<string, number>();
  const ref = XLSX.utils.decode_range(ws["!ref"] ?? "A1");
  for (let c = 0; c <= ref.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRowIdx, c })];
    if (cell?.v != null) {
      const name = String(cell.v).trim();
      if (name) m.set(name, c);
    }
  }
  return m;
}

function getOrderingPct(
  exceptionConfig: ExceptionConfig[],
  category: string,
  subcategory: string,
  descC: string
): number {
  for (const rule of exceptionConfig) {
    if (rule.status === "inactive" || rule.status === "deleted") continue;
    const catOk  = rule.category    === "ทั้งหมด" || rule.category    === category;
    const subOk  = rule.subcategory === "ทั้งหมด" || rule.subcategory === subcategory;
    const descOk = rule.descC       === "ทั้งหมด" || rule.descC       === descC;
    if (catOk && subOk && descOk) return Number(rule.percentage) / 100;
  }
  return 1.0;
}

// ─── Main worker handler ───────────────────────────────────────────────────────

addEventListener("message", (e: MessageEvent<InMsg>) => {
  if (e.data.type !== "run") return;
  const { targetBuf, spacemanBuf, indexBuf, qryBuf, exceptionConfig } = e.data;

  try {
    // ── 1. Parse DATA_SPACEMAN ──────────────────────────────────────────────
    progress(5, "อ่านไฟล์ DATA_SPACEMAN...");

    const spacemanWb = XLSX.read(spacemanBuf, {
      type: "array",
      cellText: false,
      cellHTML: false,
      cellNF: false,
      cellDates: false,
    });
    const spacemanWs = spacemanWb.Sheets["QRY_Product_by_POG"];
    if (!spacemanWs) throw new Error('ไม่พบ sheet "QRY_Product_by_POG" ใน DATA_SPACEMAN');

    const range = XLSX.utils.decode_range(spacemanWs["!ref"] || "A1");
    const hdrs: string[] = [];
    for (let c = 0; c <= range.e.c; c++) hdrs.push(cellStr(spacemanWs, 0, c));

    const upcIdx   = hdrs.indexOf("UPC");
    const pf02Idx  = hdrs.indexOf("PLANOFOLDER02");
    const pf03Idx  = hdrs.indexOf("PLANOFOLDER03");
    const pf04Idx  = hdrs.indexOf("PLANOFOLDER04");
    const plogIdx  = hdrs.indexOf("PLANOGRAM") >= 0 ? hdrs.indexOf("PLANOGRAM") : 3;
    const catIdx   = hdrs.indexOf("CATEGORY");
    const subIdx   = hdrs.indexOf("SUBCATEGORY");
    const descCIdx = hdrs.indexOf("DESC_C");

    const spacemanMap = new Map<string, SpacemanEntry>();
    const totalSpaceman = range.e.r;

    for (let r = 1; r <= totalSpaceman; r++) {
      const upc = upcIdx >= 0 ? normalizeBarcode(cellStr(spacemanWs, r, upcIdx)) : "";
      if (!upc) continue;
      if (!spacemanMap.has(upc)) {
        spacemanMap.set(upc, {
          planofolder02: pf02Idx  >= 0 ? cellStr(spacemanWs, r, pf02Idx)  : "",
          planofolder03: pf03Idx  >= 0 ? cellStr(spacemanWs, r, pf03Idx)  : "",
          planofolder04: pf04Idx  >= 0 ? cellStr(spacemanWs, r, pf04Idx)  : "",
          planogram:     cellStr(spacemanWs, r, plogIdx),
          category:      catIdx   >= 0 ? cellStr(spacemanWs, r, catIdx)   : "",
          subcategory:   subIdx   >= 0 ? cellStr(spacemanWs, r, subIdx)   : "",
          descC:         descCIdx >= 0 ? cellStr(spacemanWs, r, descCIdx) : "",
        });
      }
      if (r % 10000 === 0)
        progress(5 + Math.floor((r / totalSpaceman) * 25), `DATA_SPACEMAN: ${r.toLocaleString()} / ${totalSpaceman.toLocaleString()} rows...`);
    }
    progress(30, `DATA_SPACEMAN: ${spacemanMap.size.toLocaleString()} barcodes`);

    // ── 2. Parse INDEX ──────────────────────────────────────────────────────
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

    // ── 3. Parse QRY_Product_by_POG_by_Position ─────────────────────────────
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

    // ── 4. Read target workbook ─────────────────────────────────────────────
    progress(57, "อ่านไฟล์ Template...");

    const wb = XLSX.read(targetBuf, { type: "array", cellStyles: true });
    progress(65, "ประมวลผล Sheet 1...");

    // ── 5. Sheet 1 ──────────────────────────────────────────────────────────

    const s1Name = wb.SheetNames.find((n) => n.trimStart().startsWith("New&Exsiting For Oder"));
    if (!s1Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New&Exsiting For Oder'");
    const ws1 = wb.Sheets[s1Name];

    const ref1 = XLSX.utils.decode_range(ws1["!ref"] ?? "A1:V7998");
    const lastRow1 = Math.min(ref1.e.r + 1, 7998);

    const colMap1 = detectCols(ws1, 5);
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

    let s1Processed = 0, s1Matched = 0;

    for (let row1 = 7; row1 <= lastRow1; row1++) {
      const r = row1 - 1;
      const isSpecialRow = row1 <= 8;

      const bcCell = ws1[XLSX.utils.encode_cell({ r, c: BARCODE_COL })];
      const bc = normalizeBarcode(bcCell?.v);

      if (!isSpecialRow && bc) {
        s1Processed++;
        const spaceman = spacemanMap.get(bc);
        const qry      = qryMap.get(bc);
        const planogram = spaceman?.planogram ?? "";
        if (spaceman || qry) s1Matched++;

        const setCellS = (c: number | undefined, v: string) => {
          if (c === undefined || !v) return;
          ws1[XLSX.utils.encode_cell({ r, c })] = { t: "s", v };
        };
        const setCellN = (c: number | undefined, v: unknown) => {
          if (c === undefined || v === "" || v == null) return;
          const n = Number(v);
          if (!isNaN(n)) ws1[XLSX.utils.encode_cell({ r, c })] = { t: "n", v: n };
        };

        setCellS(DIVISION_COL, spaceman?.planofolder02 ?? "");
        setCellS(PF03_COL,     spaceman?.planofolder03 ?? "");
        setCellS(PF04_COL,     spaceman?.planofolder04 ?? "");

        const indexEntry = planogram ? indexMap.get(planogram) : undefined;
        setCellS(STATUS_COL, indexEntry?.status ?? "");
        setCellS(STORE_COL,  indexEntry?.store  ?? "");

        if (FIXTYPE_COL    !== undefined) ws1[XLSX.utils.encode_cell({ r, c: FIXTYPE_COL })]    = { t: "n", v: 0 };
        if (W_COL          !== undefined) ws1[XLSX.utils.encode_cell({ r, c: W_COL })]          = { t: "n", v: 2 };
        if (H_COL          !== undefined) ws1[XLSX.utils.encode_cell({ r, c: H_COL })]          = { t: "n", v: 1 };
        if (D_COL          !== undefined) ws1[XLSX.utils.encode_cell({ r, c: D_COL })]          = { t: "n", v: 1 };
        if (NEWFIXTURE_COL !== undefined) ws1[XLSX.utils.encode_cell({ r, c: NEWFIXTURE_COL })] = { t: "s", v: "" };

        setCellS(NOBAY_COL, qry?.segment    ?? "");
        setCellS(SEQ_COL,   qry?.locationId ?? "");
        setCellN(SHELFSTOCK_COL, qry?.totalUnits);
      }

      if (!isSpecialRow || bc) {
        const spaceman = bc ? spacemanMap.get(bc) : undefined;
        const pct = getOrderingPct(
          exceptionConfig,
          spaceman?.category    ?? "",
          spaceman?.subcategory ?? "",
          spaceman?.descC       ?? ""
        );
        ws1[XLSX.utils.encode_cell({ r, c: PCT_COL })] = { t: "n", v: pct, z: "0%" };
        const tAddr = XLSX.utils.encode_cell({ r, c: SHELFSTOCK_COL });
        const uAddr = XLSX.utils.encode_cell({ r, c: PCT_COL });
        ws1[XLSX.utils.encode_cell({ r, c: NETCAP_COL })] = { t: "n", f: `${tAddr}*${uAddr}`, v: 0 };
      }

      if (row1 % 1000 === 0)
        progress(65 + Math.floor(((row1 - 7) / (lastRow1 - 7)) * 18), `Sheet 1: ${row1.toLocaleString()} rows...`);
    }

    ref1.e.c = Math.max(ref1.e.c, NETCAP_COL);
    ref1.e.r = Math.max(ref1.e.r, lastRow1 - 1);
    ws1["!ref"] = XLSX.utils.encode_range(ref1);
    progress(83, "ประมวลผล Sheet 2...");

    // ── 6. Sheet 2 ──────────────────────────────────────────────────────────

    const s2Name = wb.SheetNames.find((n) => n.trim().startsWith("New for Link_IM"));
    if (!s2Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New for Link_IM'");
    const ws2 = wb.Sheets[s2Name];

    const ref2 = XLSX.utils.decode_range(ws2["!ref"] ?? "A1:H15522");
    const lastRow2 = Math.min(ref2.e.r + 1, 15522);

    const colMap2 = detectCols(ws2, 5);
    const BC2_COL   = colMap2.get("BARCODE")    ?? 4;
    const DIV2_COL  = colMap2.get("DIVISION")   ?? 1;
    const DEPT2_COL = colMap2.get("DEPARTMENT") ?? 2;

    let s2Processed = 0, s2Matched = 0;
    for (let row2 = 7; row2 <= lastRow2; row2++) {
      const r = row2 - 1;
      s2Processed++;
      const bc = normalizeBarcode(ws2[XLSX.utils.encode_cell({ r, c: BC2_COL })]?.v);
      const spaceman = bc ? spacemanMap.get(bc) : undefined;
      if (spaceman) {
        s2Matched++;
        if (spaceman.planofolder02)
          ws2[XLSX.utils.encode_cell({ r, c: DIV2_COL })]  = { t: "s", v: spaceman.planofolder02 };
        if (spaceman.planofolder03)
          ws2[XLSX.utils.encode_cell({ r, c: DEPT2_COL })] = { t: "s", v: spaceman.planofolder03 };
      }
    }
    progress(92, "เขียนไฟล์ผลลัพธ์...");

    // ── 7. Write output ─────────────────────────────────────────────────────

    const output = XLSX.write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
    progress(100, "เสร็จสิ้น!");

    ctx.postMessage(
      {
        type: "done",
        buffer: output,
        stats: {
          sheet1: { processed: s1Processed, matched: s1Matched },
          sheet2: { processed: s2Processed, matched: s2Matched },
        },
      } satisfies OutMsg,
      [output]
    );

  } catch (err) {
    ctx.postMessage({ type: "error", message: String(err) } satisfies OutMsg);
  }
});
