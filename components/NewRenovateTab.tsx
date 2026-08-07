"use client";

import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import DropZone from "./DropZone";
import { AlertTriangle, Download, FileSpreadsheet, CheckCircle, Loader2, XCircle } from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";

// ─── Types ────────────────────────────────────────────────────────────────────

interface SpacemanEntry {
  planofolder02: string; // DIVISION
  planofolder03: string;
  planofolder04: string;
  planogram: string;
  category: string;
  subcategory: string;
  descC: string;
}

interface QryEntry {
  segment: string;    // → No.Bay
  locationId: string; // → SEQ
  totalUnits: string; // → SHELF STOCK FOR ORDER (Piece)
}

interface Stats {
  sheet1: { processed: number; matched: number; unmatched: number };
  sheet2: { processed: number; matched: number; unmatched: number };
}

type LoadStatus = "loading" | "ready" | "error" | "no-file";
type ProcStatus = "idle" | "processing" | "done" | "error";

interface Props {
  exceptionConfig?: ExceptionConfig[];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function normalizeBarcode(val: unknown): string {
  if (val === null || val === undefined || val === "") return "";
  const s = String(val).trim();
  if (s === "") return "";
  const n = Number(s);
  if (!isNaN(n) && n > 0) return String(Math.round(n));
  return s;
}

const yield_ = () => new Promise<void>((r) => setTimeout(r, 0));

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

// Config Rules % Ordering — replicates processor.ts matchesConfig logic
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
  return 1.0; // default 100%
}

// ─── Component ────────────────────────────────────────────────────────────────

export default function NewRenovateTab({ exceptionConfig = [] }: Props) {
  // ── User file uploads ──
  const [targetFile, setTargetFile] = useState<File | null>(null);
  const [masterFile, setMasterFile] = useState<File | null>(null);
  const [qryFile, setQryFile]       = useState<File | null>(null);
  const [indexFile, setIndexFile]   = useState<File | null>(null);

  // ── Auto-fetch from Drive ──
  const [spacemanStatus, setSpacemanStatus] = useState<LoadStatus>("loading");
  const [spacemanInfo, setSpacemanInfo]     = useState<string>("");
  const [fixtureStatus, setFixtureStatus]   = useState<LoadStatus>("loading");
  const [fixtureInfo, setFixtureInfo]       = useState<string>("");

  // Pre-parsed data (built on mount)
  const spacemanMapRef = useRef<Map<string, SpacemanEntry>>(new Map());
  const fixtureMapRef  = useRef<Map<string, string>>(new Map()); // "SEG|POG" → Code Fixture

  // ── Processing state ──
  const [status, setStatus]       = useState<ProcStatus>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct]             = useState(0);
  const [stats, setStats]         = useState<Stats | null>(null);
  const [errorMsg, setErrorMsg]   = useState("");
  const outputRef = useRef<number[] | null>(null);

  // ── Auto-fetch DATA_SPACEMAN from Drive ───────────────────────────────────
  useEffect(() => {
    (async () => {
      try {
        const latestRes = await fetch("/api/spaceman/latest");
        const latestData = await latestRes.json();
        const file = latestData.file;
        if (!file) { setSpacemanStatus("no-file"); return; }

        setSpacemanInfo(file.name);
        const fileRes = await fetch(`/api/spaceman/file?id=${file.id}`);
        if (!fileRes.ok) throw new Error("ดาวน์โหลด DATA_SPACEMAN ไม่สำเร็จ");
        const buf = await fileRes.arrayBuffer();

        await yield_();
        const wb = XLSX.read(buf, { type: "array", cellText: false, cellHTML: false });
        const ws = wb.Sheets["QRY_Product_by_POG"];
        if (!ws) throw new Error('ไม่พบ sheet "QRY_Product_by_POG"');

        const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
        // Detect header positions
        const hdrs: string[] = [];
        for (let c = 0; c <= range.e.c; c++) hdrs.push(cellStr(ws, 0, c));

        const upcIdx  = hdrs.indexOf("UPC");
        const pf02Idx = hdrs.indexOf("PLANOFOLDER02");
        const pf03Idx = hdrs.indexOf("PLANOFOLDER03");
        const pf04Idx = hdrs.indexOf("PLANOFOLDER04");
        const plogIdx = hdrs.indexOf("PLANOGRAM") >= 0 ? hdrs.indexOf("PLANOGRAM") : 3;
        const catIdx  = hdrs.indexOf("CATEGORY");
        const subIdx  = hdrs.indexOf("SUBCATEGORY");
        const descCIdx= hdrs.indexOf("DESC_C");

        const map = new Map<string, SpacemanEntry>();
        for (let r = 1; r <= range.e.r; r++) {
          const upc = upcIdx >= 0 ? normalizeBarcode(cellStr(ws, r, upcIdx)) : "";
          if (!upc) continue;
          if (!map.has(upc)) {
            map.set(upc, {
              planofolder02: pf02Idx >= 0 ? cellStr(ws, r, pf02Idx) : "",
              planofolder03: pf03Idx >= 0 ? cellStr(ws, r, pf03Idx) : "",
              planofolder04: pf04Idx >= 0 ? cellStr(ws, r, pf04Idx) : "",
              planogram:     cellStr(ws, r, plogIdx),
              category:      catIdx  >= 0 ? cellStr(ws, r, catIdx)  : "",
              subcategory:   subIdx  >= 0 ? cellStr(ws, r, subIdx)  : "",
              descC:         descCIdx >= 0 ? cellStr(ws, r, descCIdx) : "",
            });
          }
          if (r % 10000 === 0) await yield_();
        }
        spacemanMapRef.current = map;
        setSpacemanStatus("ready");
      } catch (e) {
        console.error("SpacemanFetch:", e);
        setSpacemanStatus("error");
        setSpacemanInfo(String(e));
      }
    })();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // ── Auto-fetch Fixture Index from Drive ───────────────────────────────────
  useEffect(() => {
    (async () => {
      try {
        const latestRes = await fetch("/api/master/latest");
        const latestData = await latestRes.json();
        const file = latestData.file;
        if (!file) { setFixtureStatus("no-file"); return; }

        setFixtureInfo(file.name);
        const fileRes = await fetch(`/api/master/file?id=${file.id}`);
        if (!fileRes.ok) throw new Error("ดาวน์โหลด Fixture Index ไม่สำเร็จ");
        const buf = await fileRes.arrayBuffer();

        await yield_();
        const wb = XLSX.read(buf, { type: "array" });
        const targetSheet =
          wb.SheetNames.find((n) => n === "Fixture_2026") ??
          wb.SheetNames.find((n) => n.startsWith("Fixture")) ??
          wb.SheetNames[0];
        if (!targetSheet) throw new Error("ไม่พบ sheet Fixture ในไฟล์");

        const ws = wb.Sheets[targetSheet];
        // range:1 = skip remark row, row 2 becomes headers
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "", range: 1 });

        const map = new Map<string, string>();
        for (const row of rows) {
          const seg = String(row["SEG"] ?? "").trim();
          const pog = String(row["POG"] ?? "").trim();
          const code = String(row["Code Fixture"] ?? "").trim();
          if (seg && pog && code) map.set(`${seg}|${pog}`, code);
        }
        fixtureMapRef.current = map;
        setFixtureStatus("ready");
      } catch (e) {
        console.error("FixtureFetch:", e);
        setFixtureStatus("error");
        setFixtureInfo(String(e));
      }
    })();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // ─────────────────────────────────────────────────────────────────────────
  const canProcess =
    targetFile !== null &&
    masterFile !== null &&
    qryFile !== null &&
    indexFile !== null &&
    spacemanStatus === "ready" &&
    fixtureStatus === "ready" &&
    status !== "processing";

  // ── Main processing ───────────────────────────────────────────────────────
  const handleProcess = async () => {
    if (!targetFile || !masterFile || !qryFile || !indexFile) return;
    setStatus("processing");
    setErrorMsg("");
    setStats(null);
    outputRef.current = null;

    try {
      // ── 1. Master Assortment map ──────────────────────────────────────────
      setStatusMsg("อ่านไฟล์ Master Assortment..."); setPct(5); await yield_();

      const masterBuf = await masterFile.arrayBuffer();
      const masterWb = XLSX.read(masterBuf, { type: "array" });
      const masterWs = masterWb.Sheets["Sheet1"];
      if (!masterWs) throw new Error("ไม่พบ Sheet1 ใน Master Assortment");

      type MasterRow = Record<string, unknown>;
      const masterRows = XLSX.utils.sheet_to_json<MasterRow>(masterWs, { defval: "" });

      const masterMap = new Map<string, {
        barSingle: unknown; skuPack: unknown; extraInfo: string;
      }>();
      for (const row of masterRows) {
        const bc = normalizeBarcode(row["BARCODE"]);
        if (!bc) continue;
        masterMap.set(bc, {
          barSingle: row["BAR_SINGLE"],
          skuPack:   row["SKU_PACK"],
          extraInfo: String(row["EXTRA_INFO"] ?? ""),
        });
      }
      setStatusMsg(`Master: ${masterMap.size.toLocaleString()} barcodes`); setPct(18); await yield_();

      // ── 2. QRY_Product_by_POG_by_Position map ────────────────────────────
      setStatusMsg("อ่านไฟล์ QRY_Product_by_POG_by_Position..."); await yield_();

      const qryBuf = await qryFile.arrayBuffer();
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
      setStatusMsg(`QRY Position: ${qryMap.size.toLocaleString()} barcodes`); setPct(32); await yield_();

      // ── 3. INDEX.xlsx map ─────────────────────────────────────────────────
      setStatusMsg("อ่านไฟล์ INDEX.xlsx..."); await yield_();

      const indexBuf = await indexFile.arrayBuffer();
      const indexWb = XLSX.read(indexBuf, { type: "array" });
      const indexWs = indexWb.Sheets[indexWb.SheetNames[0]];
      if (!indexWs) throw new Error("ไม่พบ sheet ใน INDEX.xlsx");

      type IndexRow = Record<string, unknown>;
      const indexRows = XLSX.utils.sheet_to_json<IndexRow>(indexWs, { defval: "" });

      // Try common header names for PLANOGRAM key
      const indexMap = new Map<string, { status: string; store: string }>();
      for (const row of indexRows) {
        const plog = String(
          row["PLANOGRAM"] ?? row["Planogram"] ?? row["POG"] ?? ""
        ).trim();
        if (!plog) continue;
        if (!indexMap.has(plog)) {
          indexMap.set(plog, {
            status: String(row["Status"] ?? row["STATUS"] ?? row["สถานะ"] ?? ""),
            store:  String(row["Store"]  ?? row["STORE"]  ?? row["สาขา"]  ?? ""),
          });
        }
      }
      setStatusMsg(`INDEX: ${indexMap.size.toLocaleString()} planograms`); setPct(45); await yield_();

      // ── 4. Read target file ───────────────────────────────────────────────
      setStatusMsg("อ่านไฟล์ Target..."); await yield_();

      const targetBuf = await targetFile.arrayBuffer();
      const wb = XLSX.read(targetBuf, { type: "array", cellStyles: true });
      setPct(55); await yield_();

      // Convenience aliases
      const spacemanMap = spacemanMapRef.current;
      const fixtureMap  = fixtureMapRef.current;

      // ── 5. Sheet 1 ────────────────────────────────────────────────────────
      setStatusMsg("ประมวลผล Sheet 1..."); await yield_();

      const s1Name = wb.SheetNames.find((n) => n.trimStart().startsWith("New&Exsiting For Oder"));
      if (!s1Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New&Exsiting For Oder'");
      const ws1 = wb.Sheets[s1Name];

      const ref1 = XLSX.utils.decode_range(ws1["!ref"] ?? "A1:V7998");
      const lastRow1 = Math.min(ref1.e.r + 1, 7998);

      // Auto-detect column positions from header row (row 6 = index 5)
      const colMap1 = detectCols(ws1, 5);

      // Known fallbacks (original spec column indices)
      const c1 = (name: string, fallback: number) => colMap1.get(name) ?? fallback;
      const BARCODE_COL  = c1("BARCODE", 5);           // F
      const DIVISION_COL = c1("DIVISION", 1);           // B
      const PF03_COL     = colMap1.get("PLANOFOLDER03") ?? 2; // C
      const PF04_COL     = colMap1.get("PLANOFOLDER04") ?? 3; // D
      const SALEPACK_COL = colMap1.get("SALE PACK CODE") ?? 7; // H
      const PACKSIZE_COL = colMap1.get("Pack Size") ?? 8;       // I
      const EXTRA_COL    = colMap1.get("Extra info") ?? 9;       // J
      // Fuzzy-find new columns
      const findCol = (needle: string) => {
        for (const [k, v] of colMap1) if (k.toLowerCase().includes(needle.toLowerCase())) return v;
        return undefined;
      };
      const STATUS_COL    = findCol("Status");
      const STORE_COL     = findCol("Store");
      const FIXTYPE_COL   = findCol("Fixture Type");
      const W_COL         = findCol("\" W\"") ?? findCol("W");
      const H_COL         = findCol("\" H\"") ?? findCol("H");
      const D_COL         = findCol("\" D\"") ?? findCol("D");
      const NEWFIXTURE_COL= findCol("New Fixture");
      const NOBAY_COL     = findCol("No.Bay");
      const SEQ_COL       = findCol("SEQ");
      const SHELFSTOCK_COL= findCol("SHELF STOCK") ?? c1("SHELF STOCK FOR ORDER (Piece)", 19);
      const PCT_COL       = findCol("% Ordering") ?? 20;  // U
      const NETCAP_COL    = findCol("Net Capacity") ?? 21; // V

      let s1Processed = 0, s1Matched = 0;

      for (let row1 = 7; row1 <= lastRow1; row1++) {
        const r = row1 - 1;
        const isSpecialRow = row1 <= 8; // Non POG1/Non POG2

        const bcCell = ws1[XLSX.utils.encode_cell({ r, c: BARCODE_COL })];
        const bc = normalizeBarcode(bcCell?.v);

        if (!isSpecialRow && bc) {
          s1Processed++;
          const spaceman = spacemanMap.get(bc);
          const master   = masterMap.get(bc);
          const qry      = qryMap.get(bc);

          const planogram = spaceman?.planogram ?? "";
          const segment   = qry?.segment ?? "";

          if (spaceman || master || qry) s1Matched++;

          const setCellS = (c: number | undefined, v: string) => {
            if (c === undefined || !v) return;
            ws1[XLSX.utils.encode_cell({ r, c })] = { t: "s", v };
          };
          const setCellN = (c: number | undefined, v: unknown) => {
            if (c === undefined || v === "" || v === null || v === undefined) return;
            const n = Number(v);
            if (!isNaN(n)) ws1[XLSX.utils.encode_cell({ r, c })] = { t: "n", v: n };
          };

          // DIVISION, PLANOFOLDER03, PLANOFOLDER04 from DATA_SPACEMAN
          setCellS(DIVISION_COL, spaceman?.planofolder02 ?? "");
          setCellS(PF03_COL,     spaceman?.planofolder03 ?? "");
          setCellS(PF04_COL,     spaceman?.planofolder04 ?? "");

          // SALE PACK CODE (BAR_SINGLE), Pack Size (SKU_PACK), Extra info
          if (master) {
            setCellN(SALEPACK_COL, master.barSingle);
            setCellN(PACKSIZE_COL, master.skuPack);
            setCellS(EXTRA_COL,    master.extraInfo);
          }

          // Status, Store from INDEX.xlsx (join by PLANOGRAM)
          const indexEntry = planogram ? indexMap.get(planogram) : undefined;
          setCellS(STATUS_COL, indexEntry?.status ?? "");
          setCellS(STORE_COL,  indexEntry?.store  ?? "");

          // Fixture Type = 0, W = 2, H = 1, D = 1 (constants)
          if (FIXTYPE_COL !== undefined) ws1[XLSX.utils.encode_cell({ r, c: FIXTYPE_COL })] = { t: "n", v: 0 };
          if (W_COL !== undefined)       ws1[XLSX.utils.encode_cell({ r, c: W_COL })]       = { t: "n", v: 2 };
          if (H_COL !== undefined)       ws1[XLSX.utils.encode_cell({ r, c: H_COL })]       = { t: "n", v: 1 };
          if (D_COL !== undefined)       ws1[XLSX.utils.encode_cell({ r, c: D_COL })]       = { t: "n", v: 1 };

          // New Fixture from Fixture Index (join by PLANOGRAM+SEGMENT)
          if (planogram && segment) {
            const fixtureCode = fixtureMap.get(`${segment}|${planogram}`);
            setCellS(NEWFIXTURE_COL, fixtureCode ?? "");
          }

          // No.Bay = SEGMENT, SEQ = LOCATION_ID, SHELF STOCK = TOTAL_UNITS
          setCellS(NOBAY_COL, segment);
          setCellS(SEQ_COL,   qry?.locationId ?? "");
          setCellN(SHELFSTOCK_COL, qry?.totalUnits);
        }

        // % Ordering (with Config Rules), Net Capacity formula — ALL rows including special rows
        const spaceman = bc ? spacemanMap.get(bc) : undefined;
        const pct = getOrderingPct(
          exceptionConfig,
          spaceman?.category    ?? "",
          spaceman?.subcategory ?? "",
          spaceman?.descC       ?? ""
        );
        ws1[XLSX.utils.encode_cell({ r, c: PCT_COL })]    = { t: "n", v: pct, z: "0%" };
        const tAddr = XLSX.utils.encode_cell({ r, c: SHELFSTOCK_COL });
        const uAddr = XLSX.utils.encode_cell({ r, c: PCT_COL });
        ws1[XLSX.utils.encode_cell({ r, c: NETCAP_COL })] = { t: "n", f: `${tAddr}*${uAddr}`, v: 0 };

        if (row1 % 500 === 0) await yield_();
      }

      ref1.e.c = Math.max(ref1.e.c, NETCAP_COL);
      ref1.e.r = Math.max(ref1.e.r, lastRow1 - 1);
      ws1["!ref"] = XLSX.utils.encode_range(ref1);
      setPct(78); await yield_();

      // ── 6. Sheet 2 ────────────────────────────────────────────────────────
      setStatusMsg("ประมวลผล Sheet 2..."); await yield_();

      const s2Name = wb.SheetNames.find((n) => n.trim().startsWith("New for Link_IM"));
      if (!s2Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New for Link_IM'");
      const ws2 = wb.Sheets[s2Name];

      const ref2 = XLSX.utils.decode_range(ws2["!ref"] ?? "A1:H15522");
      const lastRow2 = Math.min(ref2.e.r + 1, 15522);

      const colMap2    = detectCols(ws2, 5);
      const BC2_COL    = colMap2.get("BARCODE") ?? 4;       // E
      const DIV2_COL   = colMap2.get("DIVISION") ?? 1;      // B
      const DEPT2_COL  = colMap2.get("DEPARTMENT") ?? 2;    // C

      let s2Processed = 0, s2Matched = 0;

      for (let row2 = 7; row2 <= lastRow2; row2++) {
        const r = row2 - 1;
        s2Processed++;
        const bcCell = ws2[XLSX.utils.encode_cell({ r, c: BC2_COL })];
        const bc = normalizeBarcode(bcCell?.v);
        const spaceman = bc ? spacemanMapRef.current.get(bc) : undefined;

        if (spaceman) {
          s2Matched++;
          if (spaceman.planofolder02)
            ws2[XLSX.utils.encode_cell({ r, c: DIV2_COL })] = { t: "s", v: spaceman.planofolder02 };
          if (spaceman.planofolder03)
            ws2[XLSX.utils.encode_cell({ r, c: DEPT2_COL })] = { t: "s", v: spaceman.planofolder03 };
        }

        if (row2 % 500 === 0) await yield_();
      }

      // ── 7. Write output ───────────────────────────────────────────────────
      setStatusMsg("เขียนไฟล์ผลลัพธ์..."); setPct(92); await yield_();
      const output = XLSX.write(wb, { bookType: "xlsx", type: "array" }) as number[];
      outputRef.current = output;

      setStats({
        sheet1: { processed: s1Processed, matched: s1Matched, unmatched: s1Processed - s1Matched },
        sheet2: { processed: s2Processed, matched: s2Matched, unmatched: s2Processed - s2Matched },
      });
      setPct(100);
      setStatus("done");
      setStatusMsg("เสร็จสิ้น!");

    } catch (e) {
      setStatus("error");
      setErrorMsg(String(e));
    }
  };

  const handleDownload = () => {
    if (!outputRef.current || !targetFile) return;
    const blob = new Blob([new Uint8Array(outputRef.current)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const base = targetFile.name.replace(/\.xlsx?$/i, "");
    a.href = url;
    a.download = `${base}_filled.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleReset = () => {
    setTargetFile(null); setMasterFile(null); setQryFile(null); setIndexFile(null);
    setStatus("idle"); setStatusMsg(""); setPct(0); setStats(null); setErrorMsg("");
    outputRef.current = null;
  };

  // ── Render ────────────────────────────────────────────────────────────────
  const AutoFetchBadge = ({ status: s, label, info }: { status: LoadStatus; label: string; info: string }) => {
    const map = {
      loading: { icon: <Loader2 className="w-3 h-3 animate-spin" />, cls: "text-slate-400", text: "กำลังโหลด..." },
      ready:   { icon: <CheckCircle className="w-3 h-3" />,          cls: "text-emerald-600", text: info || "พร้อม" },
      error:   { icon: <XCircle className="w-3 h-3" />,              cls: "text-red-500", text: "โหลดไม่สำเร็จ" },
      "no-file": { icon: <AlertTriangle className="w-3 h-3" />,      cls: "text-amber-500", text: "ยังไม่มีไฟล์ใน Drive" },
    };
    const { icon, cls, text } = map[s];
    return (
      <div className={`flex items-center gap-1.5 text-xs ${cls}`}>
        {icon}
        <span className="font-medium">{label}:</span>
        <span className="truncate max-w-[200px]">{text}</span>
      </div>
    );
  };

  return (
    <div className="max-w-5xl mx-auto p-6 space-y-6">

      {/* Header */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div className="flex items-start gap-4">
          <div className="bg-gradient-to-br from-pink-50 to-orange-50 rounded-xl p-3 flex-shrink-0">
            <FileSpreadsheet className="w-8 h-8 text-[#E91E8C]" />
          </div>
          <div className="flex-1">
            <h2 className="text-lg font-bold text-slate-800">TO BE Mini New&amp;Renovate Report Filler</h2>
            <p className="text-sm text-slate-500 mt-1">
              เติมข้อมูลทุกคอลัมน์จาก DATA_SPACEMAN · Master Assortment · QRY Position · INDEX · Fixture Index
            </p>
            <div className="mt-3 space-y-1.5 bg-slate-50 rounded-xl p-3">
              <p className="text-xs font-semibold text-slate-500 mb-1">Auto-fetch จาก Google Drive:</p>
              <AutoFetchBadge status={spacemanStatus} label="DATA_SPACEMAN" info={spacemanInfo} />
              <AutoFetchBadge status={fixtureStatus}  label="Fixture Index"  info={fixtureInfo} />
            </div>
          </div>
        </div>
      </div>

      {/* Upload zones */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <p className="text-sm font-semibold text-slate-700 mb-3">📋 ไฟล์ Target</p>
          <DropZone label="TO BE Mini New&Renovate Report.xlsx" accept=".xlsx"
            files={targetFile ? [targetFile] : []}
            onFiles={(fs) => { setTargetFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="ไฟล์ที่มี sheet New&Exsiting / New for Link_IM" />
        </div>
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <p className="text-sm font-semibold text-slate-700 mb-3">📦 Master Assortment</p>
          <DropZone label="Master_Assortment_Orderable_*.xlsx" accept=".xlsx"
            files={masterFile ? [masterFile] : []}
            onFiles={(fs) => { setMasterFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="SALE PACK CODE · Pack Size · Extra info" />
        </div>
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <p className="text-sm font-semibold text-slate-700 mb-3">📊 QRY_Product_by_POG_by_Position</p>
          <DropZone label="QRY_Product_by_POG_by_Position.xlsx" accept=".xlsx"
            files={qryFile ? [qryFile] : []}
            onFiles={(fs) => { setQryFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="SEGMENT · LOCATION_ID · TOTAL_UNITS — join by BARCODE" />
        </div>
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <p className="text-sm font-semibold text-slate-700 mb-3">📁 INDEX.xlsx</p>
          <DropZone label="INDEX.xlsx" accept=".xlsx"
            files={indexFile ? [indexFile] : []}
            onFiles={(fs) => { setIndexFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="Status · Store — join by PLANOGRAM" />
        </div>
      </div>

      {/* Drive not ready warning */}
      {(spacemanStatus === "no-file" || spacemanStatus === "error" ||
        fixtureStatus === "no-file" || fixtureStatus === "error") && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex gap-3">
          <AlertTriangle className="w-4 h-4 text-amber-500 mt-0.5 flex-shrink-0" />
          <div className="text-xs text-amber-700 space-y-1">
            {(spacemanStatus === "no-file" || spacemanStatus === "error") && (
              <p>⚠ DATA_SPACEMAN: ไปอัปโหลดไฟล์ใน Tab DATA_SPACEMAN ก่อน เพื่อให้ระบบดึงข้อมูล DIVISION / PLANOFOLDER ได้</p>
            )}
            {(fixtureStatus === "no-file" || fixtureStatus === "error") && (
              <p>⚠ Fixture Index: ไปอัปโหลดไฟล์ใน Tab Master ก่อน เพื่อให้ระบบดึง New Fixture Code ได้</p>
            )}
          </div>
        </div>
      )}

      {/* Action buttons */}
      {status !== "processing" && (
        <div className="flex gap-3">
          <button onClick={handleProcess} disabled={!canProcess}
            className="flex items-center gap-2 px-6 py-3 rounded-xl font-semibold text-sm
              bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white shadow-sm hover:shadow-md
              disabled:opacity-40 disabled:cursor-not-allowed transition-all">
            ⚡ ประมวลผลและเติมข้อมูล
          </button>
          {(status === "done" || status === "error") && (
            <button onClick={handleReset}
              className="px-6 py-3 rounded-xl font-semibold text-sm border border-pink-200 text-[#d41679] hover:bg-pink-50 transition-all">
              เริ่มใหม่
            </button>
          )}
        </div>
      )}

      {/* Progress */}
      {status === "processing" && (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-4">
          <div className="flex items-center gap-3">
            <div className="animate-spin rounded-full h-5 w-5 border-2 border-pink-200 border-t-[#E91E8C] flex-shrink-0" />
            <p className="text-slate-600 font-medium text-sm">{statusMsg}</p>
          </div>
          <div className="w-full bg-slate-200 rounded-full h-2.5 overflow-hidden">
            <div className="h-2.5 rounded-full transition-all duration-500 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]"
              style={{ width: `${pct}%` }} />
          </div>
          <p className="text-xs text-slate-400 text-right">{pct}%</p>
        </div>
      )}

      {/* Error */}
      {status === "error" && (
        <div className="bg-red-50 border border-red-200 rounded-xl p-4 text-red-700 text-sm">
          ❌ {errorMsg}
        </div>
      )}

      {/* Results + Download */}
      {status === "done" && stats && (
        <div className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
              <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">Sheet 1 — New&amp;Exsiting For Oder</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between"><span className="text-slate-600">แถวที่ประมวลผล</span><span className="font-semibold">{stats.sheet1.processed.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-emerald-600">Matched</span><span className="font-semibold text-emerald-600">{stats.sheet1.matched.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">ไม่พบข้อมูล</span><span className="font-semibold text-slate-400">{stats.sheet1.unmatched.toLocaleString()}</span></div>
                <div className="pt-2 border-t border-slate-100 text-xs text-slate-400">
                  ✓ ใช้ Config Rules ({exceptionConfig.filter(e=>e.status==="active").length} rules active) กำหนด % Ordering
                </div>
              </div>
            </div>
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
              <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">Sheet 2 — New for Link_IM</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between"><span className="text-slate-600">แถวที่ประมวลผล</span><span className="font-semibold">{stats.sheet2.processed.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-emerald-600">Matched</span><span className="font-semibold text-emerald-600">{stats.sheet2.matched.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">ไม่พบข้อมูล</span><span className="font-semibold text-slate-400">{stats.sheet2.unmatched.toLocaleString()}</span></div>
              </div>
            </div>
          </div>

          <button onClick={handleDownload}
            className="w-full flex items-center justify-center gap-2 px-6 py-4 rounded-2xl font-bold text-base
              bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-md hover:shadow-lg hover:scale-[1.01] transition-all">
            <Download className="w-5 h-5" />
            ดาวน์โหลดไฟล์ที่เติมข้อมูลแล้ว
          </button>
        </div>
      )}
    </div>
  );
}
