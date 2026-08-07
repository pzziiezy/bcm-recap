"use client";

import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import DropZone from "./DropZone";
import { Download, FileSpreadsheet } from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";

// ─── Types ────────────────────────────────────────────────────────────────────

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

interface Stats {
  sheet1: { processed: number; matched: number };
  sheet2: { processed: number; matched: number };
}

type ProcStatus = "idle" | "processing" | "done" | "error";

interface Props {
  exceptionConfig?: ExceptionConfig[];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function normalizeBarcode(val: unknown): string {
  if (val === null || val === undefined || val === "") return "";
  const s = String(val).trim();
  if (!s) return "";
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

// ─── Component ────────────────────────────────────────────────────────────────

export default function NewRenovateTab({ exceptionConfig = [] }: Props) {
  const [targetFile, setTargetFile]     = useState<File | null>(null);
  const [spacemanFile, setSpacemanFile] = useState<File | null>(null);
  const [indexFile, setIndexFile]       = useState<File | null>(null);
  const [qryFile, setQryFile]           = useState<File | null>(null);

  const [status, setStatus]       = useState<ProcStatus>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct]             = useState(0);
  const [stats, setStats]         = useState<Stats | null>(null);
  const [errorMsg, setErrorMsg]   = useState("");
  const outputRef = useRef<number[] | null>(null);

  const canProcess =
    !!targetFile && !!spacemanFile && !!indexFile && !!qryFile &&
    status !== "processing";

  // ── Process ───────────────────────────────────────────────────────────────

  const handleProcess = async () => {
    if (!targetFile || !spacemanFile || !indexFile || !qryFile) return;
    setStatus("processing");
    setErrorMsg("");
    setStats(null);
    outputRef.current = null;

    try {
      // ── 1. Parse DATA_SPACEMAN ────────────────────────────────────────────
      setStatusMsg("อ่านไฟล์ DATA_SPACEMAN..."); setPct(5); await yield_();

      const spacemanBuf = await spacemanFile.arrayBuffer();
      const spacemanWb = XLSX.read(spacemanBuf, { type: "array", cellText: false, cellHTML: false });
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
      for (let r = 1; r <= range.e.r; r++) {
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
        if (r % 10000 === 0) await yield_();
      }
      setStatusMsg(`DATA_SPACEMAN: ${spacemanMap.size.toLocaleString()} barcodes`); setPct(30); await yield_();

      // ── 2. Parse INDEX ────────────────────────────────────────────────────
      setStatusMsg("อ่านไฟล์ INDEX..."); await yield_();

      const indexBuf = await indexFile.arrayBuffer();
      const indexWb  = XLSX.read(indexBuf, { type: "array" });
      const indexWs  = indexWb.Sheets[indexWb.SheetNames[0]];
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
      setStatusMsg(`INDEX: ${indexMap.size.toLocaleString()} planograms`); setPct(45); await yield_();

      // ── 3. Parse QRY_Product_by_POG_by_Position ───────────────────────────
      setStatusMsg("อ่านไฟล์ QRY_Product_by_POG_by_Position..."); await yield_();

      const qryBuf = await qryFile.arrayBuffer();
      const qryWb  = XLSX.read(qryBuf, { type: "array" });
      const qryWs  = qryWb.Sheets[qryWb.SheetNames[0]];
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
      setStatusMsg(`QRY Position: ${qryMap.size.toLocaleString()} barcodes`); setPct(55); await yield_();

      // ── 4. Read target workbook ───────────────────────────────────────────
      setStatusMsg("อ่านไฟล์ Template..."); await yield_();

      const targetBuf = await targetFile.arrayBuffer();
      const wb = XLSX.read(targetBuf, { type: "array", cellStyles: true });
      setPct(65); await yield_();

      // ── 5. Sheet 1 ────────────────────────────────────────────────────────
      setStatusMsg("ประมวลผล Sheet 1..."); await yield_();

      const s1Name = wb.SheetNames.find((n) => n.trimStart().startsWith("New&Exsiting For Oder"));
      if (!s1Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New&Exsiting For Oder'");
      const ws1 = wb.Sheets[s1Name];

      const ref1 = XLSX.utils.decode_range(ws1["!ref"] ?? "A1:V7998");
      const lastRow1 = Math.min(ref1.e.r + 1, 7998);

      const colMap1 = detectCols(ws1, 5);
      const c1 = (name: string, fallback: number) => colMap1.get(name) ?? fallback;
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

        // % Ordering + Net Capacity — all data rows
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

        if (row1 % 500 === 0) await yield_();
      }

      ref1.e.c = Math.max(ref1.e.c, NETCAP_COL);
      ref1.e.r = Math.max(ref1.e.r, lastRow1 - 1);
      ws1["!ref"] = XLSX.utils.encode_range(ref1);
      setPct(82); await yield_();

      // ── 6. Sheet 2 ────────────────────────────────────────────────────────
      setStatusMsg("ประมวลผล Sheet 2..."); await yield_();

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
        if (row2 % 500 === 0) await yield_();
      }

      // ── 7. Write output ───────────────────────────────────────────────────
      setStatusMsg("เขียนไฟล์ผลลัพธ์..."); setPct(95); await yield_();
      const output = XLSX.write(wb, { bookType: "xlsx", type: "array" }) as number[];
      outputRef.current = output;

      setStats({
        sheet1: { processed: s1Processed, matched: s1Matched },
        sheet2: { processed: s2Processed, matched: s2Matched },
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
    a.href = url;
    a.download = targetFile.name.replace(/\.xlsx?$/i, "") + "_filled.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleReset = () => {
    setTargetFile(null); setSpacemanFile(null); setIndexFile(null); setQryFile(null);
    setStatus("idle"); setStatusMsg(""); setPct(0); setStats(null); setErrorMsg("");
    outputRef.current = null;
  };

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className="max-w-4xl mx-auto p-6 space-y-6">

      {/* Header */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex items-start gap-4">
        <div className="bg-gradient-to-br from-pink-50 to-orange-50 rounded-xl p-3 flex-shrink-0">
          <FileSpreadsheet className="w-8 h-8 text-[#E91E8C]" />
        </div>
        <div>
          <h2 className="text-lg font-bold text-slate-800">TO BE Mini New&amp;Renovate Report Filler</h2>
          <p className="text-sm text-slate-500 mt-1">
            อัปโหลด 4 ไฟล์ตามลำดับ แล้วกด Build เพื่อรับไฟล์ที่เติมข้อมูลครบแล้ว
          </p>
        </div>
      </div>

      {/* Upload zones */}
      <div className="space-y-3">
        {/* 1 — Template */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-3">
            <span className="w-6 h-6 rounded-full bg-[#E91E8C] text-white text-xs font-bold flex items-center justify-center flex-shrink-0">1</span>
            <p className="text-sm font-semibold text-slate-700">Template New&amp;Renovate Report.xlsx</p>
          </div>
          <DropZone label="Template_New&Renovate_Report.xlsx" accept=".xlsx"
            files={targetFile ? [targetFile] : []}
            onFiles={(fs) => { setTargetFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="ไฟล์ที่ต้องการเติมข้อมูล — มี sheet New&Exsiting For Oder / New for Link_IM" />
        </div>

        {/* 2 — DATA_SPACEMAN */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-3">
            <span className="w-6 h-6 rounded-full bg-[#E91E8C] text-white text-xs font-bold flex items-center justify-center flex-shrink-0">2</span>
            <p className="text-sm font-semibold text-slate-700">DATA_SPACEMAN</p>
          </div>
          <DropZone label="DATA_SPACEMAN.xlsx / .xlsb" accept=".xlsx,.xlsb,.xls"
            files={spacemanFile ? [spacemanFile] : []}
            onFiles={(fs) => { setSpacemanFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="ไฟล์เดียวกับที่ใช้ใน Menu DATA_SPACEMAN — ต้องมี sheet QRY_Product_by_POG" />
        </div>

        {/* 3 — INDEX */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-3">
            <span className="w-6 h-6 rounded-full bg-[#E91E8C] text-white text-xs font-bold flex items-center justify-center flex-shrink-0">3</span>
            <p className="text-sm font-semibold text-slate-700">FILE INDEX</p>
          </div>
          <DropZone label="INDEX.xlsx" accept=".xlsx,.xls"
            files={indexFile ? [indexFile] : []}
            onFiles={(fs) => { setIndexFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="Status · Store — join by PLANOGRAM" />
        </div>

        {/* 4 — QRY */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-3">
            <span className="w-6 h-6 rounded-full bg-[#E91E8C] text-white text-xs font-bold flex items-center justify-center flex-shrink-0">4</span>
            <p className="text-sm font-semibold text-slate-700">QRY_Product by POG by Position</p>
          </div>
          <DropZone label="QRY_Product_by_POG_by_Position.xlsx" accept=".xlsx,.xls"
            files={qryFile ? [qryFile] : []}
            onFiles={(fs) => { setQryFile(fs[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="SEGMENT · LOCATION_ID · TOTAL_UNITS — join by BARCODE" />
        </div>
      </div>

      {/* Action */}
      {status !== "processing" && (
        <div className="flex gap-3">
          <button onClick={handleProcess} disabled={!canProcess}
            className="flex items-center gap-2 px-8 py-3 rounded-xl font-bold text-sm
              bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-sm hover:shadow-md
              disabled:opacity-40 disabled:cursor-not-allowed transition-all">
            ⚡ Build New&amp;Renovate Report
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
          <div className="grid grid-cols-2 gap-4">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
              <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">Sheet 1 — New&amp;Exsiting</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between"><span className="text-slate-600">แถวที่ประมวลผล</span><span className="font-semibold">{stats.sheet1.processed.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-emerald-600">Matched</span><span className="font-semibold text-emerald-600">{stats.sheet1.matched.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">ไม่พบ barcode</span><span className="font-semibold text-slate-400">{(stats.sheet1.processed - stats.sheet1.matched).toLocaleString()}</span></div>
                <div className="pt-2 border-t border-slate-100 text-xs text-slate-400">
                  Config Rules active: {exceptionConfig.filter(e => e.status === "active").length} rules
                </div>
              </div>
            </div>
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
              <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">Sheet 2 — New for Link_IM</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between"><span className="text-slate-600">แถวที่ประมวลผล</span><span className="font-semibold">{stats.sheet2.processed.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-emerald-600">Matched</span><span className="font-semibold text-emerald-600">{stats.sheet2.matched.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">ไม่พบ barcode</span><span className="font-semibold text-slate-400">{(stats.sheet2.processed - stats.sheet2.matched).toLocaleString()}</span></div>
              </div>
            </div>
          </div>

          <button onClick={handleDownload}
            className="w-full flex items-center justify-center gap-2 px-6 py-4 rounded-2xl font-bold text-base
              bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-md hover:shadow-lg hover:scale-[1.01] transition-all">
            <Download className="w-5 h-5" />
            ดาวน์โหลด New&amp;Renovate Report ที่เติมข้อมูลแล้ว
          </button>
        </div>
      )}
    </div>
  );
}
