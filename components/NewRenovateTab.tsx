"use client";

import { useState, useRef, useEffect } from "react";
import DropZone from "./DropZone";
import { Download, FileSpreadsheet } from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";

// ─── Types ────────────────────────────────────────────────────────────────────

interface Stats {
  sheet1: { written: number; matchedSpaceman: number; matchedIndex: number };
  sheet2: { written: number; matched: number };
}

type ProcStatus = "idle" | "processing" | "done" | "error";

interface Props {
  exceptionConfig?: ExceptionConfig[];
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

  const outputRef = useRef<ArrayBuffer | null>(null);
  const workerRef = useRef<Worker | null>(null);

  // Terminate worker on unmount
  useEffect(() => () => { workerRef.current?.terminate(); }, []);

  const canProcess =
    !!targetFile && !!spacemanFile && !!indexFile && !!qryFile &&
    status !== "processing";

  // ── Process ───────────────────────────────────────────────────────────────

  const handleProcess = async () => {
    if (!targetFile || !spacemanFile || !indexFile || !qryFile) return;

    setStatus("processing");
    setStatusMsg("กำลังโหลดไฟล์...");
    setPct(2);
    setErrorMsg("");
    setStats(null);
    outputRef.current = null;

    // Read all files to ArrayBuffers on main thread (fast I/O, not CPU)
    const [targetBuf, spacemanBuf, indexBuf, qryBuf] = await Promise.all([
      targetFile.arrayBuffer(),
      spacemanFile.arrayBuffer(),
      indexFile.arrayBuffer(),
      qryFile.arrayBuffer(),
    ]);

    // Terminate any previous worker
    workerRef.current?.terminate();

    const worker = new Worker(
      new URL("../lib/newrenovate.worker.ts", import.meta.url)
    );
    workerRef.current = worker;

    worker.onmessage = (e: MessageEvent) => {
      const msg = e.data as
        | { type: "progress"; pct: number; msg: string }
        | { type: "done"; buffer: ArrayBuffer; stats: Stats }
        | { type: "error"; message: string };

      if (msg.type === "progress") {
        setPct(msg.pct);
        setStatusMsg(msg.msg);
      } else if (msg.type === "done") {
        outputRef.current = msg.buffer;
        setStats(msg.stats);
        setStatus("done");
        setStatusMsg("เสร็จสิ้น!");
        setPct(100);
        worker.terminate();
        workerRef.current = null;
      } else if (msg.type === "error") {
        setStatus("error");
        setErrorMsg(msg.message);
        worker.terminate();
        workerRef.current = null;
      }
    };

    worker.onerror = (e: ErrorEvent) => {
      setStatus("error");
      setErrorMsg(e.message ?? "Worker crashed");
      workerRef.current = null;
    };

    // Transfer all buffers to worker (zero-copy)
    worker.postMessage(
      { type: "run", targetBuf, spacemanBuf, indexBuf, qryBuf, exceptionConfig },
      [targetBuf, spacemanBuf, indexBuf, qryBuf]
    );
  };

  const handleDownload = () => {
    if (!outputRef.current || !targetFile) return;
    const blob = new Blob([outputRef.current], {
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
    workerRef.current?.terminate();
    workerRef.current = null;
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
            อัปโหลด 4 ไฟล์ตามลำดับ แล้วกด Build — ประมวลผลบน Background Thread (UI ไม่กระตุก)
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
                <div className="flex justify-between"><span className="text-slate-600">แถวที่เขียน (จาก QRY)</span><span className="font-semibold">{stats.sheet1.written.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-emerald-600">Matched DATA_SPACEMAN</span><span className="font-semibold text-emerald-600">{stats.sheet1.matchedSpaceman.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-emerald-600">Matched INDEX</span><span className="font-semibold text-emerald-600">{stats.sheet1.matchedIndex.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">ไม่พบใน DATA_SPACEMAN</span><span className="font-semibold text-slate-400">{(stats.sheet1.written - stats.sheet1.matchedSpaceman).toLocaleString()}</span></div>
                <div className="pt-2 border-t border-slate-100 text-xs text-slate-400">
                  Config Rules active: {exceptionConfig.filter(e => e.status === "active").length} rules
                </div>
              </div>
            </div>
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
              <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">Sheet 2 — New for Link_IM</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between"><span className="text-slate-600">แถวที่เขียน (จาก QRY)</span><span className="font-semibold">{stats.sheet2.written.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-emerald-600">Matched DATA_SPACEMAN</span><span className="font-semibold text-emerald-600">{stats.sheet2.matched.toLocaleString()}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">ไม่พบใน DATA_SPACEMAN</span><span className="font-semibold text-slate-400">{(stats.sheet2.written - stats.sheet2.matched).toLocaleString()}</span></div>
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
