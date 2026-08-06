"use client";

import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import DropZone from "./DropZone";
import { AlertTriangle, Download, FileSpreadsheet } from "lucide-react";

interface ProcessStats {
  sheet1: { processed: number; matched: number; unmatched: number };
  sheet2: { processed: number; matched: number; unmatched: number };
}

type Status = "idle" | "processing" | "done" | "error";

function normalizeBarcode(val: unknown): string {
  if (val === null || val === undefined || val === "") return "";
  const s = String(val).trim();
  if (s === "") return "";
  // Handle float representation e.g. "8850096730518.0"
  const n = Number(s);
  if (!isNaN(n) && n > 0) return String(Math.round(n));
  return s;
}

const yield_ = () => new Promise<void>((r) => setTimeout(r, 0));

export default function NewRenovateTab() {
  const [targetFile, setTargetFile] = useState<File | null>(null);
  const [masterFile, setMasterFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct] = useState(0);
  const [stats, setStats] = useState<ProcessStats | null>(null);
  const [errorMsg, setErrorMsg] = useState("");
  const outputRef = useRef<number[] | null>(null);

  const canProcess = targetFile !== null && masterFile !== null && status !== "processing";

  const handleProcess = async () => {
    if (!targetFile || !masterFile) return;
    setStatus("processing");
    setErrorMsg("");
    setStats(null);
    outputRef.current = null;

    try {
      // ── Step 1: Build Master Assortment lookup ──────────────────────────
      setStatusMsg("กำลังอ่านไฟล์ Master Assortment...");
      setPct(5);
      await yield_();

      const masterBuf = await masterFile.arrayBuffer();
      const masterWb = XLSX.read(masterBuf, { type: "array" });
      const masterSheet = masterWb.Sheets["Sheet1"];
      if (!masterSheet) throw new Error("ไม่พบ Sheet1 ใน Master Assortment file");

      setStatusMsg("สร้าง Barcode index...");
      setPct(15);
      await yield_();

      type MasterRow = Record<string, unknown>;
      const masterRows = XLSX.utils.sheet_to_json<MasterRow>(masterSheet, { defval: "" });

      const masterMap = new Map<string, {
        DIVISION: string; DEPARTMENT: string; SUB_DEPARTMENT: string;
        CLASS: string; EXTRA_INFO: string; SKU_PACK: unknown; BAR_SINGLE: unknown;
      }>();

      for (const row of masterRows) {
        const bc = normalizeBarcode(row["BARCODE"]);
        if (!bc) continue;
        masterMap.set(bc, {
          DIVISION:       String(row["DIVISION"]       ?? ""),
          DEPARTMENT:     String(row["DEPARTMENT"]     ?? ""),
          SUB_DEPARTMENT: String(row["SUB_DEPARTMENT"] ?? ""),
          CLASS:          String(row["CLASS"]           ?? ""),
          EXTRA_INFO:     String(row["EXTRA_INFO"]     ?? ""),
          SKU_PACK:       row["SKU_PACK"],
          BAR_SINGLE:     row["BAR_SINGLE"],
        });
      }

      setStatusMsg(`Index สร้างเสร็จ — ${masterMap.size.toLocaleString()} barcodes`);
      setPct(25);
      await yield_();

      // ── Step 2: Read Target file ────────────────────────────────────────
      setStatusMsg("กำลังอ่านไฟล์ TO BE Mini New&Renovate Report...");
      const targetBuf = await targetFile.arrayBuffer();
      const wb = XLSX.read(targetBuf, { type: "array", cellStyles: true });
      setPct(35);
      await yield_();

      // ── Sheet 1: New&Exsiting For Oder_SCM+MIS ─────────────────────────
      setStatusMsg("ประมวลผล Sheet 1: New&Exsiting For Oder_SCM+MIS...");
      const s1Name = wb.SheetNames.find((n) => n.trimStart().startsWith("New&Exsiting For Oder"));
      if (!s1Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New&Exsiting For Oder'");
      const ws1 = wb.Sheets[s1Name];

      const ref1 = XLSX.utils.decode_range(ws1["!ref"] ?? "A1:V7998");
      const lastRow1 = Math.min(ref1.e.r + 1, 7998); // 1-indexed

      let s1Processed = 0, s1Matched = 0;

      for (let row1 = 7; row1 <= lastRow1; row1++) {
        const r = row1 - 1; // xlsx 0-indexed row
        const isSpecialRow = row1 <= 8; // rows 7-8 = Non POG1/Non POG2 — skip barcode lookup

        if (!isSpecialRow) {
          s1Processed++;
          const bcCell = ws1[XLSX.utils.encode_cell({ r, c: 5 })]; // col F = BARCODE
          const bc = normalizeBarcode(bcCell?.v);
          const m = bc ? masterMap.get(bc) : undefined;

          if (m) {
            s1Matched++;
            ws1[XLSX.utils.encode_cell({ r, c: 1  })] = { t: "s", v: m.DIVISION };       // B DIVISION
            ws1[XLSX.utils.encode_cell({ r, c: 2  })] = { t: "s", v: m.SUB_DEPARTMENT }; // C PLANOFOLDER03
            ws1[XLSX.utils.encode_cell({ r, c: 3  })] = { t: "s", v: m.CLASS };           // D PLANOFOLDER04
            // H = SALE PACK CODE = BAR_SINGLE (only when numeric and non-null)
            const bs = m.BAR_SINGLE;
            if (bs !== "" && bs !== null && bs !== undefined) {
              const bsNum = Number(bs);
              if (!isNaN(bsNum)) ws1[XLSX.utils.encode_cell({ r, c: 7 })] = { t: "n", v: bsNum };
            }
            // I = Pack Size = SKU_PACK
            const sp = Number(m.SKU_PACK);
            if (!isNaN(sp)) ws1[XLSX.utils.encode_cell({ r, c: 8 })] = { t: "n", v: sp };
            // J = Extra info
            ws1[XLSX.utils.encode_cell({ r, c: 9 })] = { t: "s", v: m.EXTRA_INFO };
          }
        }

        // U = % Ordering = 1.0 (100%), col index 20, format 0%
        ws1[XLSX.utils.encode_cell({ r, c: 20 })] = { t: "n", v: 1.0, z: "0%" };
        // V = Net Capacity = T*U formula, col index 21
        const tAddr = XLSX.utils.encode_cell({ r, c: 19 }); // T
        const uAddr = XLSX.utils.encode_cell({ r, c: 20 }); // U
        ws1[XLSX.utils.encode_cell({ r, c: 21 })] = { t: "n", f: `${tAddr}*${uAddr}`, v: 0 };
      }

      // Expand !ref to cover new cols U (20) and V (21)
      ref1.e.c = Math.max(ref1.e.c, 21);
      ref1.e.r = Math.max(ref1.e.r, lastRow1 - 1);
      ws1["!ref"] = XLSX.utils.encode_range(ref1);

      setPct(65);
      await yield_();

      // ── Sheet 2: New for Link_IM ────────────────────────────────────────
      setStatusMsg("ประมวลผล Sheet 2: New for Link_IM...");
      // Sheet name has trailing space — match by trimmed prefix
      const s2Name = wb.SheetNames.find((n) => n.trim().startsWith("New for Link_IM"));
      if (!s2Name) throw new Error("ไม่พบ sheet ที่ขึ้นต้นด้วย 'New for Link_IM'");
      const ws2 = wb.Sheets[s2Name];

      const ref2 = XLSX.utils.decode_range(ws2["!ref"] ?? "A1:H15522");
      const lastRow2 = Math.min(ref2.e.r + 1, 15522);

      let s2Processed = 0, s2Matched = 0;

      for (let row2 = 7; row2 <= lastRow2; row2++) {
        const r = row2 - 1;
        s2Processed++;
        const bcCell = ws2[XLSX.utils.encode_cell({ r, c: 4 })]; // col E = BARCODE
        const bc = normalizeBarcode(bcCell?.v);
        const m = bc ? masterMap.get(bc) : undefined;

        if (m) {
          s2Matched++;
          ws2[XLSX.utils.encode_cell({ r, c: 1 })] = { t: "s", v: m.DIVISION };   // B
          ws2[XLSX.utils.encode_cell({ r, c: 2 })] = { t: "s", v: m.DEPARTMENT }; // C
          // D = POG CATE: open issue — source unclear, skip
        }
      }

      // ── Step 3: Write output ────────────────────────────────────────────
      setStatusMsg("เขียนไฟล์ผลลัพธ์...");
      setPct(88);
      await yield_();

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
    setTargetFile(null);
    setMasterFile(null);
    setStatus("idle");
    setStatusMsg("");
    setPct(0);
    setStats(null);
    setErrorMsg("");
    outputRef.current = null;
  };

  return (
    <div className="max-w-4xl mx-auto p-6 space-y-6">

      {/* Header */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div className="flex items-start gap-4">
          <div className="bg-gradient-to-br from-pink-50 to-orange-50 rounded-xl p-3 flex-shrink-0">
            <FileSpreadsheet className="w-8 h-8 text-[#E91E8C]" />
          </div>
          <div>
            <h2 className="text-lg font-bold text-slate-800">TO BE Mini New&amp;Renovate Report Filler</h2>
            <p className="text-sm text-slate-500 mt-1">
              เติมข้อมูล DIVISION / SUB_DEPT / CLASS / Pack / Extra Info จาก Master Assortment
              พร้อม % Ordering (100%) และ Net Capacity formula อัตโนมัติ
            </p>
            <div className="mt-3 flex flex-wrap gap-2">
              {[
                "New&Exsiting For Oder → 8 คอลัมน์",
                "New for Link_IM → 2 คอลัมน์",
                "% Ordering = 100%",
                "=T×U formula",
              ].map((t) => (
                <span key={t} className="px-2 py-0.5 rounded-full bg-pink-50 text-[#E91E8C] text-xs font-medium border border-pink-100">
                  {t}
                </span>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Upload zones */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <p className="text-sm font-semibold text-slate-700 mb-3">📋 ไฟล์ Target (ที่ต้องเติมข้อมูล)</p>
          <DropZone
            label="TO BE Mini New&Renovate Report.xlsx"
            accept=".xlsx"
            files={targetFile ? [targetFile] : []}
            onFiles={(files) => { setTargetFile(files[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="ไฟล์ที่มี sheet New&Exsiting / New for Link_IM"
          />
        </div>
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
          <p className="text-sm font-semibold text-slate-700 mb-3">📦 Master Assortment</p>
          <DropZone
            label="Master_Assortment_Orderable_*.xlsx"
            accept=".xlsx"
            files={masterFile ? [masterFile] : []}
            onFiles={(files) => { setMasterFile(files[0] ?? null); if (status !== "idle") handleReset(); }}
            hint="~236k rows — join ด้วย BARCODE"
          />
        </div>
      </div>

      {/* Open issues notice */}
      <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex gap-3">
        <AlertTriangle className="w-4 h-4 text-amber-500 mt-0.5 flex-shrink-0" />
        <div className="text-xs text-amber-700 space-y-1">
          <p className="font-semibold">Open Issues — รอข้อมูลเพิ่มเติม (ข้ามไว้ก่อน):</p>
          <p>• <b>col D — POG CATE (POG 04,03)</b> ใน Sheet 2: แหล่งข้อมูลยังไม่ชัดเจน ไม่ได้มาจาก Master Assortment</p>
          <p>• <b>LOCATION_ID ↔ Store No.</b>: ยังไม่มี crosswalk จึงยังไม่รองรับข้อมูลจาก QRY_Product_by_POG</p>
        </div>
      </div>

      {/* Action buttons */}
      {status !== "processing" && (
        <div className="flex gap-3">
          <button
            onClick={handleProcess}
            disabled={!canProcess}
            className="flex items-center gap-2 px-6 py-3 rounded-xl font-semibold text-sm
              bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white shadow-sm hover:shadow-md
              disabled:opacity-40 disabled:cursor-not-allowed transition-all"
          >
            ⚡ ประมวลผลและเติมข้อมูล
          </button>
          {(status === "done" || status === "error") && (
            <button
              onClick={handleReset}
              className="px-6 py-3 rounded-xl font-semibold text-sm border border-pink-200 text-[#d41679] hover:bg-pink-50 transition-all"
            >
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
            <div
              className="h-2.5 rounded-full transition-all duration-500 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]"
              style={{ width: `${pct}%` }}
            />
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
            {/* Sheet 1 stats */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
              <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">Sheet 1 — New&amp;Exsiting For Oder_SCM+MIS</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-slate-600">แถวที่ประมวลผล</span>
                  <span className="font-semibold">{stats.sheet1.processed.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-emerald-600">Matched</span>
                  <span className="font-semibold text-emerald-600">{stats.sheet1.matched.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">ไม่พบใน Master</span>
                  <span className="font-semibold text-slate-400">{stats.sheet1.unmatched.toLocaleString()}</span>
                </div>
                <div className="pt-2 border-t border-slate-100 text-xs text-slate-400">
                  ✓ % Ordering (col U) = 100% และ Net Capacity formula (col V) เติมทุกแถว
                </div>
              </div>
            </div>
            {/* Sheet 2 stats */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
              <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">Sheet 2 — New for Link_IM</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-slate-600">แถวที่ประมวลผล</span>
                  <span className="font-semibold">{stats.sheet2.processed.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-emerald-600">Matched</span>
                  <span className="font-semibold text-emerald-600">{stats.sheet2.matched.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">ไม่พบใน Master</span>
                  <span className="font-semibold text-slate-400">{stats.sheet2.unmatched.toLocaleString()}</span>
                </div>
                <div className="pt-2 border-t border-slate-100 text-xs text-amber-500">
                  ⚠ col D (POG CATE) ข้ามไว้ — รอข้อมูลแหล่งที่มา
                </div>
              </div>
            </div>
          </div>

          <button
            onClick={handleDownload}
            className="w-full flex items-center justify-center gap-2 px-6 py-4 rounded-2xl font-bold text-base
              bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-md hover:shadow-lg hover:scale-[1.01] transition-all"
          >
            <Download className="w-5 h-5" />
            ดาวน์โหลดไฟล์ที่เติมข้อมูลแล้ว
          </button>
        </div>
      )}
    </div>
  );
}
