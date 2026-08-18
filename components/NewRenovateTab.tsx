"use client";

import { useState, useRef, useEffect, useMemo } from "react";
import { Download, FileSpreadsheet, X } from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";

// ─── Types ────────────────────────────────────────────────────────────────────

interface Stats {
  total:           number;
  matchedSpaceman: number;
  matchedMaster:   number;
  matchedIndex:    number;
  matchedFixture:  number;
  masterMapSize?:  number;
  masterSheetName?: string;
}

interface PreviewRow {
  barcode:    string;
  planograms: string[];
  storesA:    string[];
  storesB:    string[];
  status:     "EXISTING" | "NEW EXPAND" | "DELETE";
}

interface PreviewData {
  rows:                PreviewRow[];
  unmatchedPlanograms: string[];
}

type ProcStatus   = "idle" | "processing" | "preview" | "done" | "error";
type StatusFilter = "ALL" | "EXISTING" | "NEW EXPAND" | "DELETE";

interface Props {
  exceptionConfig?: ExceptionConfig[];
  fixtureRows?:     Record<string, string>[];
}

// ─── Small helpers ────────────────────────────────────────────────────────────

const STATUS_STYLE: Record<PreviewRow["status"], string> = {
  EXISTING:     "bg-emerald-100 text-emerald-700 border border-emerald-200",
  "NEW EXPAND": "bg-blue-100 text-blue-700 border border-blue-200",
  DELETE:       "bg-red-100 text-red-600 border border-red-200",
};

function StatusPill({ status }: { status: PreviewRow["status"] }) {
  return (
    <span className={`inline-flex px-2 py-0.5 rounded-full text-[10px] font-bold tracking-wide whitespace-nowrap ${STATUS_STYLE[status]}`}>
      {status}
    </span>
  );
}

function StoreList({ stores }: { stores: string[] }) {
  if (stores.length === 0) return <span className="text-slate-300">—</span>;
  const visible = stores.slice(0, 3);
  const extra   = stores.length - 3;
  return (
    <span className="text-slate-600 tabular-nums">
      {visible.join(", ")}
      {extra > 0 && <span className="ml-1 text-slate-400 font-medium">+{extra}</span>}
    </span>
  );
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
  const [targetFile,   setTargetFile]   = useState<File | null>(null);
  const [qryFile,      setQryFile]      = useState<File | null>(null);
  const [spacemanFile, setSpacemanFile] = useState<File | null>(null);
  const [masterFile,   setMasterFile]   = useState<File | null>(null);
  const [indexFile,    setIndexFile]    = useState<File | null>(null);

  const [status,     setStatus]     = useState<ProcStatus>("idle");
  const [statusMsg,  setStatusMsg]  = useState("");
  const [pct,        setPct]        = useState(0);
  const [stats,      setStats]      = useState<Stats | null>(null);
  const [errorMsg,   setErrorMsg]   = useState("");

  const [previewData,  setPreviewData]  = useState<PreviewData | null>(null);
  const [previewTab,   setPreviewTab]   = useState<"compare" | "unmatched">("compare");
  const [statusFilter, setStatusFilter] = useState<StatusFilter>("ALL");

  const outputRef = useRef<ArrayBuffer | null>(null);
  const workerRef = useRef<Worker | null>(null);

  useEffect(() => () => { workerRef.current?.terminate(); }, []);

  const canProcess =
    !!targetFile && !!qryFile && !!spacemanFile &&
    !!masterFile && !!indexFile &&
    status !== "processing";

  const filteredRows = useMemo<PreviewRow[]>(() => {
    if (!previewData) return [];
    if (statusFilter === "ALL") return previewData.rows;
    return previewData.rows.filter(r => r.status === statusFilter);
  }, [previewData, statusFilter]);

  const countByStatus = useMemo(() => {
    if (!previewData) return { EXISTING: 0, "NEW EXPAND": 0, DELETE: 0 };
    return previewData.rows.reduce(
      (acc, r) => { acc[r.status]++; return acc; },
      { EXISTING: 0, "NEW EXPAND": 0, DELETE: 0 } as Record<PreviewRow["status"], number>,
    );
  }, [previewData]);

  // ── Reset ─────────────────────────────────────────────────────────────────

  const handleReset = () => {
    workerRef.current?.terminate();
    workerRef.current = null;
    setTargetFile(null); setQryFile(null); setSpacemanFile(null);
    setMasterFile(null); setIndexFile(null);
    setStatus("idle"); setStatusMsg(""); setPct(0);
    setStats(null); setErrorMsg(""); setPreviewData(null);
    setPreviewTab("compare"); setStatusFilter("ALL");
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
    setPreviewData(null);
    setPreviewTab("compare");
    setStatusFilter("ALL");
    outputRef.current = null;

    const [targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf] = await Promise.all([
      targetFile.arrayBuffer(),
      qryFile.arrayBuffer(),
      spacemanFile.arrayBuffer(),
      masterFile.arrayBuffer(),
      indexFile.arrayBuffer(),
    ]);

    workerRef.current?.terminate();

    const worker = new Worker(new URL("../lib/newrenovate.worker.ts", import.meta.url));
    workerRef.current = worker;

    worker.onmessage = (e: MessageEvent) => {
      const msg = e.data as
        | { type: "progress"; pct: number; msg: string }
        | { type: "done"; buffer: ArrayBuffer; stats: Stats; preview?: PreviewData }
        | { type: "error"; message: string };

      if (msg.type === "progress") {
        setPct(msg.pct);
        setStatusMsg(msg.msg);
      } else if (msg.type === "done") {
        outputRef.current = msg.buffer;
        setStats(msg.stats);
        setPreviewData(msg.preview ?? null);
        setStatus(msg.preview ? "preview" : "done");
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
      const details = [e.message, e.filename ? `${e.filename}:${e.lineno}` : ""].filter(Boolean).join(" @ ");
      setErrorMsg(details || "Worker crashed (ดู Console สำหรับ error details)");
      workerRef.current = null;
    };

    worker.postMessage(
      { type: "run", targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf, fixtureRows, exceptionConfig },
      [targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf],
    );
  };

  // ── Download ──────────────────────────────────────────────────────────────

  const handleDownload = () => {
    if (!outputRef.current || !targetFile) return;
    const blob = new Blob([outputRef.current], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a   = document.createElement("a");
    a.href     = url;
    a.download = targetFile.name.replace(/\.xlsx?$/i, "") + "_filled.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // ── Confirm (preview → done + trigger download) ───────────────────────────

  const handleConfirm = () => {
    setStatus("done");
    handleDownload();
  };

  // ── Render helpers ────────────────────────────────────────────────────────

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

  const slot = (
    num: number, title: string, hint: string,
    file: File | null, setter: (f: File | null) => void, accept: string,
  ) => (
    <CompactUploadSlot key={num} num={num} title={title} hint={hint} accept={accept}
      file={file} onFile={(f) => { setter(f); if (status !== "idle") handleReset(); }} />
  );

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className="w-full">

      <div className="max-w-2xl mx-auto px-6 pt-6 space-y-4">

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
          {(status === "done" || status === "error" || status === "preview") && (
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

      </div>

      {/* ── Preview Panel — full viewport width ───────────────────────────────── */}
      {status === "preview" && previewData && (
        <div className="px-6 pt-4 pb-6">
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">

          {/* Header */}
          <div className="px-5 pt-5 pb-4 border-b border-slate-100">
            <h3 className="text-sm font-bold text-slate-700">ตรวจสอบข้อมูลก่อนดาวน์โหลด</h3>
            <p className="text-xs text-slate-400 mt-0.5">
              ตรวจสอบ comparison ระหว่าง AS IS และ TO BE แล้วกดยืนยันเพื่อ download Report
            </p>

            {/* Summary chips */}
            <div className="flex flex-wrap gap-2 mt-3">
              {(
                [
                  ["EXISTING",  countByStatus.EXISTING,               "emerald"] ,
                  ["NEW EXPAND",countByStatus["NEW EXPAND"],          "blue"]    ,
                  ["DELETE",    countByStatus.DELETE,                  "red"]     ,
                  ["ไม่พบใน INDEX", previewData.unmatchedPlanograms.length, "amber"],
                ] as const
              ).map(([label, count, color]) => {
                const colorCls = {
                  emerald: "bg-emerald-50 text-emerald-700 border border-emerald-200",
                  blue:    "bg-blue-50 text-blue-700 border border-blue-200",
                  red:     "bg-red-50 text-red-600 border border-red-200",
                  amber:   "bg-amber-50 text-amber-700 border border-amber-200",
                }[color];
                return (
                  <span key={label} className={`text-[11px] font-semibold px-2.5 py-1 rounded-full ${colorCls}`}>
                    {label}: <span className="tabular-nums">{count.toLocaleString()}</span>
                  </span>
                );
              })}
            </div>
          </div>

          {/* Tabs */}
          <div className="flex border-b border-slate-100 bg-slate-50/50">
            {(
              [
                { key: "compare",   label: `Comparison (${previewData.rows.length.toLocaleString()} barcodes)` },
                { key: "unmatched", label: `Planogram ไม่พบใน INDEX (${previewData.unmatchedPlanograms.length})` },
              ] as const
            ).map(({ key, label }) => (
              <button
                key={key}
                onClick={() => setPreviewTab(key)}
                className={`px-5 py-3 text-xs font-semibold transition-colors border-b-2 -mb-px ${
                  previewTab === key
                    ? "border-[#E91E8C] text-[#E91E8C] bg-white"
                    : "border-transparent text-slate-400 hover:text-slate-600"
                }`}
              >
                {label}
              </button>
            ))}
          </div>

          {/* ── Tab: Comparison ─────────────────────────────────────────────── */}
          {previewTab === "compare" && (
            <>
              {/* Status filter */}
              <div className="flex gap-1.5 px-5 py-3 border-b border-slate-100 bg-slate-50/30">
                <span className="text-[10px] text-slate-400 font-medium self-center mr-1">Filter:</span>
                {(["ALL", "EXISTING", "NEW EXPAND", "DELETE"] as StatusFilter[]).map(f => {
                  const active = statusFilter === f;
                  const activeCls: Record<StatusFilter, string> = {
                    ALL:          "bg-slate-700 text-white border-slate-700",
                    EXISTING:     "bg-emerald-600 text-white border-emerald-600",
                    "NEW EXPAND": "bg-blue-600 text-white border-blue-600",
                    DELETE:       "bg-red-500 text-white border-red-500",
                  };
                  return (
                    <button
                      key={f}
                      onClick={() => setStatusFilter(f)}
                      className={`text-[10px] px-2.5 py-1 rounded-full border font-semibold transition-all ${
                        active ? activeCls[f] : "bg-white text-slate-400 border-slate-200 hover:border-slate-300 hover:text-slate-600"
                      }`}
                    >
                      {f}
                    </button>
                  );
                })}
                {filteredRows.length !== previewData.rows.length && (
                  <span className="ml-auto text-[10px] text-slate-400 self-center tabular-nums">
                    {filteredRows.length.toLocaleString()} rows
                  </span>
                )}
              </div>

              {/* Table */}
              <div className="overflow-auto" style={{ maxHeight: "60vh" }}>
                <table className="w-full text-xs border-collapse">
                  <thead>
                    <tr className="bg-slate-50 text-slate-500 font-semibold text-[11px] sticky top-0 z-10">
                      <th className="px-4 py-2.5 text-left border-b border-slate-200 w-36">BARCODE</th>
                      <th className="px-4 py-2.5 text-left border-b border-slate-200">PLANOGRAM NAME</th>
                      <th className="px-4 py-2.5 text-left border-b border-slate-200 w-36">AS IS Stores</th>
                      <th className="px-4 py-2.5 text-left border-b border-slate-200 w-36">TO BE Stores</th>
                      <th className="px-4 py-2.5 text-left border-b border-slate-200 w-28">STATUS</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.length === 0 && (
                      <tr>
                        <td colSpan={5} className="px-4 py-8 text-center text-slate-400 text-xs">
                          ไม่มีข้อมูล
                        </td>
                      </tr>
                    )}
                    {filteredRows.slice(0, 1000).map((row, idx) => (
                      <tr
                        key={row.barcode}
                        className={`border-t border-slate-100 hover:bg-slate-50/80 transition-colors ${
                          idx % 2 === 0 ? "" : "bg-slate-50/40"
                        }`}
                      >
                        <td className="px-4 py-2 font-mono text-slate-700 text-[11px] whitespace-nowrap">{row.barcode}</td>
                        <td className="px-4 py-2 text-slate-600 max-w-[180px]">
                          <span className="block truncate" title={row.planograms.join(", ")}>
                            {row.planograms.join(", ")}
                          </span>
                        </td>
                        <td className="px-4 py-2"><StoreList stores={row.storesA} /></td>
                        <td className="px-4 py-2"><StoreList stores={row.storesB} /></td>
                        <td className="px-4 py-2"><StatusPill status={row.status} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {filteredRows.length > 1000 && (
                  <div className="px-5 py-3 text-center text-[11px] text-slate-400 border-t border-slate-100">
                    แสดง 1,000 จาก {filteredRows.length.toLocaleString()} rows
                  </div>
                )}
              </div>
            </>
          )}

          {/* ── Tab: Unmatched Planograms ────────────────────────────────────── */}
          {previewTab === "unmatched" && (
            <div className="p-5" style={{ maxHeight: 460, overflow: "auto" }}>
              {previewData.unmatchedPlanograms.length === 0 ? (
                <div className="flex items-center gap-2 text-emerald-600 text-sm font-medium">
                  <span className="text-base">✓</span>
                  PLANOGRAM ทุกตัวพบในไฟล์ INDEX
                </div>
              ) : (
                <>
                  <p className="text-xs text-slate-500 mb-3">
                    PLANOGRAM NAME เหล่านี้ไม่พบในไฟล์ INDEX — ข้อมูลจาก QRY row เหล่านั้นจะไม่ถูก fill ใน report
                  </p>
                  <div className="grid grid-cols-3 gap-1.5">
                    {previewData.unmatchedPlanograms.map(p => (
                      <span
                        key={p}
                        className="text-xs px-3 py-1.5 bg-amber-50 border border-amber-200 text-amber-700 rounded-lg truncate"
                        title={p}
                      >
                        {p}
                      </span>
                    ))}
                  </div>
                </>
              )}
            </div>
          )}

          {/* Confirm button */}
          <div className="px-5 py-4 border-t border-slate-100 bg-slate-50/30">
            <button
              onClick={handleConfirm}
              className="w-full flex items-center justify-center gap-2 px-6 py-4 rounded-xl font-bold text-base
                bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-md hover:shadow-lg
                hover:scale-[1.01] transition-all"
            >
              <Download className="w-5 h-5" />
              ยืนยันและดาวน์โหลด New&amp;Renovate Report
            </button>
          </div>
        </div>
        </div>
      )}

      {/* Results + Download */}
      {status === "done" && stats && (
        <div className="max-w-2xl mx-auto px-6 pb-6 space-y-4">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
            <p className="text-xs font-bold text-slate-500 mb-4 uppercase tracking-wide">ผลการ Lookup</p>
            <div className="grid grid-cols-5 gap-3">
              {badge(stats.total,           "QRY rows (ตั้งต้น)",       "pink")}
              {badge(stats.matchedSpaceman, "Matched DATA_SPACEMAN",     "emerald")}
              {badge(stats.matchedMaster,   "Matched Master Assortment", "emerald")}
              {badge(stats.matchedIndex,    "Matched INDEX",             "emerald")}
              {badge(stats.matchedFixture,  "Matched Fixture Index",     "emerald")}
            </div>
            <div className="mt-3 pt-3 border-t border-slate-100 grid grid-cols-2 gap-2 text-xs text-slate-400">
              <span>ไม่พบใน DATA_SPACEMAN: {(stats.total - stats.matchedSpaceman).toLocaleString()}</span>
              <span>ไม่พบใน Master Assortment: {(stats.total - stats.matchedMaster).toLocaleString()}</span>
              <span>ไม่พบใน INDEX: {(stats.total - stats.matchedIndex).toLocaleString()}</span>
              <span>ไม่พบใน Fixture Index: {(stats.total - stats.matchedFixture).toLocaleString()}</span>
            </div>
            <div className="mt-2 text-xs text-slate-400">
              Config Rules active: {exceptionConfig.filter(e => e.status === "active").length} rules
              {stats.masterMapSize !== undefined && (
                <span className="ml-4 text-amber-500 font-semibold">
                  [Debug] Master: sheet=&quot;{stats.masterSheetName}&quot; | loaded {stats.masterMapSize.toLocaleString()} barcodes
                </span>
              )}
            </div>
          </div>

          <button onClick={handleDownload}
            className="w-full flex items-center justify-center gap-2 px-6 py-4 rounded-2xl font-bold text-base
              bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-md hover:shadow-lg
              hover:scale-[1.01] transition-all">
            <Download className="w-5 h-5" />
            ดาวน์โหลด New&amp;Renovate Report ที่เติมข้อมูลแล้ว
          </button>
        </div>
      )}
    </div>
  );
}
