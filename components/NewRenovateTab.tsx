"use client";

import { useState, useRef, useEffect } from "react";
import { Download, FileSpreadsheet, X } from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";

// ─── Types ────────────────────────────────────────────────────────────────────

interface Stats {
  total: number;
  matchedSpaceman: number;
  matchedMaster: number;
  matchedIndex: number;
  matchedFixture: number;
}

type ProcStatus = "idle" | "processing" | "done" | "error";

interface Props {
  exceptionConfig?: ExceptionConfig[];
  fixtureRows?: Record<string, string>[];
}

// ─── Compact upload slot ──────────────────────────────────────────────────────

function CompactUploadSlot({
  num, title, hint, file, onFile, accept,
}: {
  num: number; title: string; hint: string;
  file: File | null; onFile: (f: File | null) => void; accept: string;
}) {
  const ref = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  const pick = (list: FileList | null) => { if (list?.[0]) onFile(list[0]); };
  return (
    <div className={`rounded-xl border flex items-center gap-3 px-4 py-3 transition-all ${
      file  ? "bg-green-50 border-green-200" :
      drag  ? "bg-pink-50 border-[#E91E8C]"  :
              "bg-white border-slate-200 hover:border-pink-200"
    }`}>
      <input ref={ref} type="file" accept={accept} className="hidden"
        onChange={e => pick(e.target.files)} />
      <span className={`w-5 h-5 rounded-full text-white text-[10px] font-bold flex items-center justify-center flex-shrink-0 ${
        file ? "bg-[#72BF44]" : "bg-[#E91E8C]"
      }`}>{num}</span>
      <div
        className="flex-1 min-w-0 cursor-pointer"
        onClick={() => { if (!file) ref.current?.click(); }}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); pick(e.dataTransfer.files); }}
      >
        <p className="text-sm font-semibold text-slate-700 truncate">{title}</p>
        {file
          ? <p className="text-xs text-green-600 font-medium truncate">✓ {file.name}</p>
          : <p className="text-xs text-slate-400 truncate">{hint}</p>}
      </div>
      {file ? (
        <button onClick={() => onFile(null)} className="flex-shrink-0 text-slate-300 hover:text-red-400 transition-colors">
          <X className="w-4 h-4" />
        </button>
      ) : (
        <button onClick={() => ref.current?.click()}
          className="flex-shrink-0 text-xs px-3 py-1.5 rounded-lg border border-pink-200 text-pink-400 hover:border-[#E91E8C] hover:text-[#E91E8C] hover:bg-pink-50 transition-all font-medium whitespace-nowrap">
          เลือกไฟล์
        </button>
      )}
    </div>
  );
}

// ─── Component ────────────────────────────────────────────────────────────────

export default function NewRenovateTab({ exceptionConfig = [], fixtureRows = [] }: Props) {
  const [targetFile,  setTargetFile]  = useState<File | null>(null); // 1 Template
  const [qryFile,     setQryFile]     = useState<File | null>(null); // 2 QRY (primary)
  const [spacemanFile, setSpacemanFile] = useState<File | null>(null); // 3 DATA_SPACEMAN
  const [masterFile,  setMasterFile]  = useState<File | null>(null); // 4 Master Assortment
  const [indexFile,   setIndexFile]   = useState<File | null>(null); // 5 INDEX

  const [status,    setStatus]    = useState<ProcStatus>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct,       setPct]       = useState(0);
  const [stats,     setStats]     = useState<Stats | null>(null);
  const [errorMsg,  setErrorMsg]  = useState("");

  const outputRef  = useRef<ArrayBuffer | null>(null);
  const workerRef  = useRef<Worker | null>(null);

  useEffect(() => () => { workerRef.current?.terminate(); }, []);

  const canProcess =
    !!targetFile && !!qryFile && !!spacemanFile &&
    !!masterFile && !!indexFile &&
    status !== "processing";

  // ── Reset ─────────────────────────────────────────────────────────────────

  const handleReset = () => {
    workerRef.current?.terminate();
    workerRef.current = null;
    setTargetFile(null); setQryFile(null); setSpacemanFile(null);
    setMasterFile(null); setIndexFile(null);
    setStatus("idle"); setStatusMsg(""); setPct(0);
    setStats(null); setErrorMsg("");
    outputRef.current = null;
  };

  // ── Process ───────────────────────────────────────────────────────────────

  const handleProcess = async () => {
    if (!targetFile || !qryFile || !spacemanFile || !masterFile || !indexFile) return;

    setStatus("processing");
    setStatusMsg("กำลังโหลดไฟล์...");
    setPct(2);
    setErrorMsg("");
    setStats(null);
    outputRef.current = null;

    const [targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf] = await Promise.all([
      targetFile.arrayBuffer(),
      qryFile.arrayBuffer(),
      spacemanFile.arrayBuffer(),
      masterFile.arrayBuffer(),
      indexFile.arrayBuffer(),
    ]);

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

    worker.postMessage(
      { type: "run", targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf, fixtureRows, exceptionConfig },
      [targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf]
    );
  };

  // ── Download ──────────────────────────────────────────────────────────────

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

  // ── Render ────────────────────────────────────────────────────────────────

  const badge = (n: number, label: string, color: "pink" | "emerald" | "slate") => {
    const cls: Record<typeof color, string> = {
      pink:    "bg-pink-50 text-[#E91E8C] border border-pink-100",
      emerald: "bg-emerald-50 text-emerald-700 border border-emerald-100",
      slate:   "bg-slate-50 text-slate-500 border border-slate-200",
    };
    return (
      <div className={`rounded-lg px-3 py-2 ${cls[color]}`}>
        <div className="text-lg font-bold tabular-nums">{n.toLocaleString()}</div>
        <div className="text-[10px] font-medium mt-0.5 leading-tight">{label}</div>
      </div>
    );
  };

  const slot = (num: number, title: string, hint: string, file: File | null, setter: (f: File | null) => void, accept: string) => (
    <CompactUploadSlot key={num} num={num} title={title} hint={hint} accept={accept}
      file={file} onFile={(f) => { setter(f); if (status !== "idle") handleReset(); }} />
  );

  return (
    <div className="max-w-2xl mx-auto p-6 space-y-4">

      {/* Header */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 flex items-center gap-3">
        <div className="bg-gradient-to-br from-pink-50 to-orange-50 rounded-lg p-2 flex-shrink-0">
          <FileSpreadsheet className="w-6 h-6 text-[#E91E8C]" />
        </div>
        <div>
          <h2 className="text-base font-bold text-slate-800">TO BE Mini New&amp;Renovate Report Filler</h2>
          <p className="text-xs text-slate-500 mt-0.5">
            อัปโหลด 5 ไฟล์ · ข้อมูล Fixture Index โหลดอัตโนมัติจาก Tab Fixture Index
          </p>
        </div>
      </div>

      {/* Upload zones */}
      <div className="space-y-2">
        {slot(1, "Template New&Renovate Report.xlsx", "ไฟล์ output — มี sheet New&Exsiting For Oder / New for Link_IM", targetFile,  setTargetFile,  ".xlsx")}
        {slot(2, "QRY_Product by POG by Position",   "ข้อมูลตั้งต้น — BARCODE · SEGMENT · LOCATION_ID · TOTAL_UNITS",            qryFile,      setQryFile,      ".xlsx,.xls")}
        {slot(3, "DATA_SPACEMAN",                    "lookup DIVISION / PF03 / PF04 / PLANOGRAM — ต้องมี sheet QRY_Product_by_POG", spacemanFile, setSpacemanFile, ".xlsx,.xlsb,.xls")}
        {slot(4, "Master Assortment Orderable",      "lookup SALE PACK CODE · Pack Size · Extra info",                              masterFile,   setMasterFile,   ".xlsx,.xls")}
        {slot(5, "FILE INDEX",                       "lookup Status · Store — join by PLANOGRAM",                                   indexFile,    setIndexFile,    ".xlsx,.xls")}
        <div className={`rounded-xl border flex items-center gap-3 px-4 py-3 ${
          fixtureRows.length > 0 ? "bg-green-50 border-green-200" : "bg-amber-50 border-amber-200"
        }`}>
          <span className={`w-5 h-5 rounded-full text-white text-[10px] font-bold flex items-center justify-center flex-shrink-0 ${
            fixtureRows.length > 0 ? "bg-[#72BF44]" : "bg-amber-400"
          }`}>6</span>
          <div className="flex-1 min-w-0">
            <p className="text-sm font-semibold text-slate-700">Fixture Index</p>
            {fixtureRows.length > 0
              ? <p className="text-xs text-green-600 font-medium">✓ โหลดแล้ว {fixtureRows.length.toLocaleString()} rows — อ่านจาก Tab Fixture Index</p>
              : <p className="text-xs text-amber-600">⚠ ยังไม่มีข้อมูล — เปิด Tab Fixture Index เพื่อโหลด</p>
            }
          </div>
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
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
            <p className="text-xs font-bold text-slate-500 mb-4 uppercase tracking-wide">ผลการ Lookup</p>
            <div className="grid grid-cols-5 gap-3">
              {badge(stats.total,           "QRY rows (ตั้งต้น)",            "pink")}
              {badge(stats.matchedSpaceman, "Matched DATA_SPACEMAN",          "emerald")}
              {badge(stats.matchedMaster,   "Matched Master Assortment",      "emerald")}
              {badge(stats.matchedIndex,    "Matched INDEX",                  "emerald")}
              {badge(stats.matchedFixture,  "Matched Fixture Index",          "emerald")}
            </div>
            <div className="mt-3 pt-3 border-t border-slate-100 grid grid-cols-2 gap-2 text-xs text-slate-400">
              <span>ไม่พบใน DATA_SPACEMAN: {(stats.total - stats.matchedSpaceman).toLocaleString()}</span>
              <span>ไม่พบใน Master Assortment: {(stats.total - stats.matchedMaster).toLocaleString()}</span>
              <span>ไม่พบใน INDEX: {(stats.total - stats.matchedIndex).toLocaleString()}</span>
              <span>ไม่พบใน Fixture Index: {(stats.total - stats.matchedFixture).toLocaleString()}</span>
            </div>
            <div className="mt-2 text-xs text-slate-400">
              Config Rules active: {exceptionConfig.filter(e => e.status === "active").length} rules
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
