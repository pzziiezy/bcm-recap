"use client";

import { useState, useRef, useEffect, useMemo, useTransition } from "react";
import { Download, FileSpreadsheet, X, CheckCircle2 } from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";
import { makeEntry, sendLog } from "@/lib/logger";

// ─── Types ────────────────────────────────────────────────────────────────────

interface Stats {
  total:            number;
  matchedSpaceman:  number;
  matchedMaster:    number;
  matchedIndex:     number;
  matchedFixture:   number;
  s1Rows?:          number;
  s2Rows?:          number;
  s3Rows?:          number;
  masterMapSize?:   number;
  masterSheetName?: string;
}

interface PreviewRow {
  barcode:    string;
  name:       string;
  planograms: string[];
  storesA:    string[];
  storesB:    string[];
  status:     "EXISTING" | "NEW EXPAND" | "DELETE";
}

interface PreviewData {
  rows:                PreviewRow[];
  unmatchedPlanograms: string[];
}

type ProcStatus = "idle" | "processing" | "preview" | "done" | "error";
type StatusKey  = "EXISTING" | "NEW EXPAND" | "DELETE";

interface Props {
  exceptionConfig?: ExceptionConfig[];
}

// ─── Constants ────────────────────────────────────────────────────────────────

const DISPLAY_LIMIT_INIT = 200;
const DISPLAY_LIMIT_STEP = 200;

// ─── Helpers ──────────────────────────────────────────────────────────────────

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

/** Outer-join storesA and storesB — one entry per unique store number. */
function mergeStores(storesA: string[], storesB: string[]): { storeA: string; storeB: string }[] {
  if (!storesA.length && !storesB.length) return [{ storeA: "", storeB: "" }];
  const setA = new Set(storesA);
  const setB = new Set(storesB);
  const ns   = (a: string, b: string) => { const na = +a, nb = +b; return isFinite(na) && isFinite(nb) ? na - nb : a.localeCompare(b); };
  return [...new Set([...storesA, ...storesB])].sort(ns).map(s => ({
    storeA: setA.has(s) ? s : "",
    storeB: setB.has(s) ? s : "",
  }));
}

// ─── Upload slot ──────────────────────────────────────────────────────────────

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
      <input ref={ref} type="file" accept={accept} className="hidden" onChange={e => pick(e.target.files)} />
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
        <button onClick={() => onFile(null)} className="flex-shrink-0 text-slate-300 hover:text-red-400">
          <X className="w-4 h-4" />
        </button>
      ) : (
        <button onClick={() => ref.current?.click()}
          className="flex-shrink-0 text-xs px-3 py-1.5 rounded-lg border border-pink-200 text-pink-400
            hover:border-[#E91E8C] hover:text-[#E91E8C] hover:bg-pink-50 font-medium whitespace-nowrap">
          เลือกไฟล์
        </button>
      )}
    </div>
  );
}

// ─── Search input ─────────────────────────────────────────────────────────────

function SearchInput({ value, onChange, placeholder }: {
  value: string; onChange: (v: string) => void; placeholder: string;
}) {
  return (
    <div className="relative mt-1">
      <input
        type="text"
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full text-[10px] px-2 py-1 pr-5 border border-slate-200 rounded bg-white
          focus:outline-none focus:border-pink-300 focus:ring-1 focus:ring-pink-100
          placeholder:text-slate-300 font-normal"
      />
      {value && (
        <button onClick={() => onChange("")}
          className="absolute right-1 top-1/2 -translate-y-1/2 text-slate-300 hover:text-slate-500">
          <X className="w-3 h-3" />
        </button>
      )}
    </div>
  );
}

// ─── Component ────────────────────────────────────────────────────────────────

export default function NewRenovateTab({ exceptionConfig = [] }: Props) {
  const [targetFile,   setTargetFile]   = useState<File | null>(null);
  const [qryFile,      setQryFile]      = useState<File | null>(null);
  const [spacemanFile, setSpacemanFile] = useState<File | null>(null);
  const [masterFile,   setMasterFile]   = useState<File | null>(null);
  const [indexFile,    setIndexFile]    = useState<File | null>(null);
  const [fixtureFile,  setFixtureFile]  = useState<File | null>(null);

  const [status,    setStatus]    = useState<ProcStatus>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const nrSessionRef   = useRef(crypto.randomUUID()); // refreshed on each process run
  const nrStartTimeRef = useRef(0);
  const [pct,       setPct]       = useState(0);
  const [stats,     setStats]     = useState<Stats | null>(null);
  const [errorMsg,  setErrorMsg]  = useState("");

  const [previewData,    setPreviewData]    = useState<PreviewData | null>(null);
  const [previewTab,     setPreviewTab]     = useState<"compare" | "unmatched">("compare");
  const [activeFilters,  setActiveFilters]  = useState<Set<StatusKey>>(new Set());
  const [downloaded,     setDownloaded]     = useState(false);

  // ── Per-column search ──────────────────────────────────────────────────────
  const [srchBarcode,   setSrchBarcode]   = useState("");
  const [srchName,      setSrchName]      = useState("");
  const [srchPlanogram, setSrchPlanogram] = useState("");
  const [srchStoreA,    setSrchStoreA]    = useState("");
  const [srchStoreB,    setSrchStoreB]    = useState("");

  // ── Display paging ─────────────────────────────────────────────────────────
  const [displayLimit, setDisplayLimit] = useState(DISPLAY_LIMIT_INIT);

  // ── Transition for non-urgent filter/search updates ────────────────────────
  const [isPending, startTransition] = useTransition();

  const outputRef = useRef<ArrayBuffer | null>(null);
  const workerRef = useRef<Worker | null>(null);

  useEffect(() => () => { workerRef.current?.terminate(); }, []);

  const canProcess =
    !!targetFile && !!qryFile && !!spacemanFile &&
    !!masterFile && !!indexFile && !!fixtureFile &&
    status !== "processing";

  // ── Filtered rows ──────────────────────────────────────────────────────────
  const filteredRows = useMemo<PreviewRow[]>(() => {
    if (!previewData) return [];
    let rows = previewData.rows;
    if (activeFilters.size > 0) {
      rows = rows.filter(r => {
        if (activeFilters.has("EXISTING")   && r.status === "EXISTING")   return true;
        if (activeFilters.has("NEW EXPAND") && r.status === "NEW EXPAND") return true;
        if (activeFilters.has("DELETE")     && r.status === "DELETE")     return true;
        return false;
      });
    }
    if (srchBarcode.trim())
      rows = rows.filter(r => r.barcode.includes(srchBarcode.trim()));
    if (srchName.trim()) {
      const q = srchName.trim().toLowerCase();
      rows = rows.filter(r => r.name.toLowerCase().includes(q));
    }
    if (srchPlanogram.trim()) {
      const q = srchPlanogram.trim().toLowerCase();
      rows = rows.filter(r => r.planograms.some(p => p.toLowerCase().includes(q)));
    }
    if (srchStoreA.trim())
      rows = rows.filter(r => r.storesA.some(s => s.includes(srchStoreA.trim())));
    if (srchStoreB.trim())
      rows = rows.filter(r => r.storesB.some(s => s.includes(srchStoreB.trim())));
    return rows;
  }, [previewData, activeFilters, srchBarcode, srchName, srchPlanogram, srchStoreA, srchStoreB]);

  // ── Pre-computed store rows for visible slice (avoids flatMap in render) ───
  const displayedStoreRows = useMemo(() => {
    return filteredRows.slice(0, displayLimit).flatMap((row) => {
      const merged = mergeStores(row.storesA, row.storesB);
      return merged.map((sr, idx) => ({
        key:   `${row.barcode}-${idx}`,
        row, sr, idx,
        total: merged.length,
        // Per-store status: NEW = only in TO BE, DEL = only in AS IS, EXISTING = both
        storeStatus: (!sr.storeA && sr.storeB) ? "new" as const
                   : ( sr.storeA && !sr.storeB) ? "del" as const
                   : "existing" as const,
      }));
    });
  }, [filteredRows, displayLimit]);

  // ── Filter/search helpers (all transitions) ────────────────────────────────
  const toggleFilter = (f: StatusKey) => {
    const next = new Set(activeFilters);
    if (next.has(f)) next.delete(f); else next.add(f);
    startTransition(() => { setActiveFilters(next); setDisplayLimit(DISPLAY_LIMIT_INIT); });
  };
  const clearFilters = () => {
    startTransition(() => { setActiveFilters(new Set()); setDisplayLimit(DISPLAY_LIMIT_INIT); });
  };
  const applySearch = (setter: (v: string) => void) => (v: string) => {
    startTransition(() => { setter(v); setDisplayLimit(DISPLAY_LIMIT_INIT); });
  };
  const clearSearch = () => {
    startTransition(() => {
      setSrchBarcode(""); setSrchName(""); setSrchPlanogram("");
      setSrchStoreA(""); setSrchStoreB(""); setDisplayLimit(DISPLAY_LIMIT_INIT);
    });
  };
  const hasSearch = !!(srchBarcode || srchName || srchPlanogram || srchStoreA || srchStoreB);

  // ── Export unmatched planograms to Excel ───────────────────────────────────
  const handleExportUnmatched = async () => {
    if (!previewData || previewData.unmatchedPlanograms.length === 0) return;
    sendLog([makeEntry(nrSessionRef.current, "NR_EXPORT_UNMATCHED", "INFO",
      `[New&Renovate] Export unmatched planograms — ${previewData.unmatchedPlanograms.length} รายการ (PLANOGRAM ใน INDEX TO BE ไม่พบใน QRY)`,
      { count: previewData.unmatchedPlanograms.length }
    )]);
    const XLSX = await import("xlsx");
    const ws = XLSX.utils.aoa_to_sheet([
      ["PLANOGRAM NAME"],
      ...previewData.unmatchedPlanograms.map(p => [p]),
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Unmatched Planograms");
    XLSX.writeFile(wb, "unmatched_planograms.xlsx");
  };

  // ── Reset ──────────────────────────────────────────────────────────────────
  const handleReset = () => {
    workerRef.current?.terminate();
    workerRef.current = null;
    setTargetFile(null); setQryFile(null); setSpacemanFile(null);
    setMasterFile(null); setIndexFile(null); setFixtureFile(null);
    setStatus("idle"); setStatusMsg(""); setPct(0);
    setStats(null); setErrorMsg(""); setPreviewData(null);
    setPreviewTab("compare"); setActiveFilters(new Set());
    setDownloaded(false); setDisplayLimit(DISPLAY_LIMIT_INIT);
    setSrchBarcode(""); setSrchName(""); setSrchPlanogram("");
    setSrchStoreA(""); setSrchStoreB("");
    outputRef.current = null;
  };

  // ── Process ────────────────────────────────────────────────────────────────
  const handleProcess = async () => {
    if (!targetFile || !qryFile || !spacemanFile || !masterFile || !indexFile || !fixtureFile) return;

    nrSessionRef.current = crypto.randomUUID();
    nrStartTimeRef.current = Date.now();
    setStatus("processing"); setStatusMsg("กำลังโหลดไฟล์..."); setPct(2);
    setErrorMsg(""); setStats(null); setPreviewData(null);
    setPreviewTab("compare"); setActiveFilters(new Set());
    setDownloaded(false); setDisplayLimit(DISPLAY_LIMIT_INIT);
    setSrchBarcode(""); setSrchName(""); setSrchPlanogram("");
    setSrchStoreA(""); setSrchStoreB("");
    outputRef.current = null;

    sendLog([makeEntry(nrSessionRef.current, "NR_PROCESS_START", "INFO",
      `[New&Renovate] เริ่ม Build — Template: ${targetFile.name} | QRY: ${qryFile.name} | Spaceman: ${spacemanFile.name} | Master: ${masterFile.name} | INDEX: ${indexFile.name} | Fixture: ${fixtureFile.name}`,
      {
        template:  targetFile.name,
        qry:       qryFile.name,
        spaceman:  spacemanFile.name,
        master:    masterFile.name,
        index:     indexFile.name,
        fixture:   fixtureFile.name,
      }
    )]);

    const [targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf, fixtureBuf] = await Promise.all([
      targetFile.arrayBuffer(), qryFile.arrayBuffer(), spacemanFile.arrayBuffer(),
      masterFile.arrayBuffer(), indexFile.arrayBuffer(), fixtureFile.arrayBuffer(),
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
        setPct(msg.pct); setStatusMsg(msg.msg);
      } else if (msg.type === "done") {
        outputRef.current = msg.buffer;
        setStats(msg.stats);
        setPreviewData(msg.preview ?? null);
        setStatus(msg.preview ? "preview" : "done");
        worker.terminate(); workerRef.current = null;
        const durSec = ((Date.now() - nrStartTimeRef.current) / 1000).toFixed(1);
        const s = msg.stats;
        sendLog([makeEntry(nrSessionRef.current, "NR_PROCESS_COMPLETE", "INFO",
          `[New&Renovate] Build เสร็จใน ${durSec}s — Sheet1: ${s.s1Rows ?? 0} rows | Sheet2 (NEW EXPAND): ${s.s2Rows ?? 0} rows | Sheet3 (DELETE): ${s.s3Rows ?? 0} rows | จับคู่ QRY: ${s.total} items (Spaceman ${s.matchedSpaceman} / Master ${s.matchedMaster} / INDEX ${s.matchedIndex} / Fixture ${s.matchedFixture})`,
          {
            durationSec:     durSec,
            sheet1Rows:      s.s1Rows,
            sheet2Rows:      s.s2Rows,
            sheet3Rows:      s.s3Rows,
            totalQryItems:   s.total,
            matchedSpaceman: s.matchedSpaceman,
            matchedMaster:   s.matchedMaster,
            matchedIndex:    s.matchedIndex,
            matchedFixture:  s.matchedFixture,
            masterMapSize:   s.masterMapSize,
          }
        )]);
      } else if (msg.type === "error") {
        setStatus("error"); setErrorMsg(msg.message);
        worker.terminate(); workerRef.current = null;
        const durSec = ((Date.now() - nrStartTimeRef.current) / 1000).toFixed(1);
        sendLog([makeEntry(nrSessionRef.current, "NR_PROCESS_ERROR", "ERROR",
          `[New&Renovate] Build ล้มเหลวหลังจาก ${durSec}s — ${msg.message}`,
          { error: msg.message, durationSec: durSec }
        )]);
      }
    };

    worker.onerror = (e: ErrorEvent) => {
      setStatus("error");
      const details = [e.message, e.filename ? `${e.filename}:${e.lineno}` : ""].filter(Boolean).join(" @ ");
      setErrorMsg(details || "Worker crashed (ดู Console สำหรับ error details)");
      workerRef.current = null;
      const durSec = ((Date.now() - nrStartTimeRef.current) / 1000).toFixed(1);
      sendLog([makeEntry(nrSessionRef.current, "NR_PROCESS_ERROR", "ERROR",
        `[New&Renovate] Worker crashed หลังจาก ${durSec}s — ${details || "Unknown error"}`,
        { error: details, durationSec: durSec }
      )]);
    };

    worker.postMessage(
      { type: "run", targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf, fixtureBuf, exceptionConfig },
      [targetBuf, qryBuf, spacemanBuf, masterBuf, indexBuf, fixtureBuf],
    );
  };

  // ── Download ───────────────────────────────────────────────────────────────
  const handleDownload = () => {
    if (!outputRef.current || !targetFile) return;
    const blob = new Blob([outputRef.current], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a   = document.createElement("a");
    const filename = targetFile.name.replace(/\.xlsx?$/i, "") + "_filled.xlsx";
    a.href = url;
    a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
    sendLog([makeEntry(nrSessionRef.current, "NR_DOWNLOAD", "INFO",
      `[New&Renovate] ดาวน์โหลด output: ${filename} (${(outputRef.current.byteLength / 1048576).toFixed(2)} MB)`,
      { filename, sizeMB: +(outputRef.current.byteLength / 1048576).toFixed(2) }
    )]);
  };

  const handleConfirm = () => {
    handleDownload();
    setDownloaded(true); // stay on preview — no setStatus("done")
  };

  // ── Stats badges ───────────────────────────────────────────────────────────
  const badge = (n: number, label: string, color: "pink" | "emerald" | "blue" | "slate") => {
    const cls: Record<typeof color, string> = {
      pink:    "bg-pink-50 text-[#E91E8C] border border-pink-100",
      emerald: "bg-emerald-50 text-emerald-700 border border-emerald-100",
      blue:    "bg-blue-50 text-blue-700 border border-blue-100",
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
      file={file} onFile={(f) => {
        setter(f);
        if (status !== "idle") handleReset();
        if (f) sendLog([makeEntry(nrSessionRef.current, "FILE_UPLOAD", "INFO",
          `[New&Renovate] อัปโหลด ${title}: ${f.name} (${(f.size / 1048576).toFixed(2)} MB)`,
          { tab: "NewRenovate", slot: num, fileType: title, name: f.name, sizeMB: +(f.size / 1048576).toFixed(2) }
        )]);
      }} />
  );

  // ─── Render ────────────────────────────────────────────────────────────────

  const isPreview = status === "preview" && !!previewData;

  return (
    <div className="w-full">
      <div className="flex gap-4 px-4 pt-2 pb-6 items-stretch">

      {/* ── Upload panel — left sidebar ── */}
      <div className="w-96 flex-shrink-0 bg-white rounded-2xl border border-slate-200 shadow-sm p-4 flex flex-col gap-3">

        {/* Header */}
        <div className="flex items-center gap-3 pb-3 border-b border-slate-100">
          <div className="bg-gradient-to-br from-pink-50 to-orange-50 rounded-lg p-2 flex-shrink-0">
            <FileSpreadsheet className="w-6 h-6 text-[#E91E8C]" />
          </div>
          <div>
            <h2 className="text-base font-bold text-slate-800">TO BE Mini New&amp;Renovate Report Filler</h2>
            <p className="text-xs text-slate-500 mt-0.5">
              อัปโหลด 6 ไฟล์เพื่อสร้าง New&amp;Renovate Report
            </p>
          </div>
        </div>

        {/* Upload zones */}
        <div className="space-y-2">
          {slot(1, "Template New&Renovate Report.xlsx", "ไฟล์ output — มี sheet New&Exsiting For Oder / New for Link_IM", targetFile,  setTargetFile,  ".xlsx")}
          {slot(2, "QRY_Product by POG by Position",   "ข้อมูลตั้งต้น — BARCODE · SEGMENT · LOCATION_ID · TOTAL_UNITS",            qryFile,      setQryFile,      ".xlsx,.xls")}
          {slot(3, "DATA_SPACEMAN",                    "lookup DIVISION / PF03 / PF04 / PLANOGRAM — ต้องมี sheet QRY_Product_by_POG", spacemanFile, setSpacemanFile, ".xlsx,.xlsb,.xls")}
          {slot(4, "Master Assortment Orderable",      "lookup SALE PACK CODE · Pack Size · Extra info",                              masterFile,   setMasterFile,   ".xlsx,.xls")}
          {slot(5, "FILE INDEX",    "lookup Status · Store — join by PLANOGRAM",                    indexFile,   setIndexFile,   ".xlsx,.xls")}
          {slot(6, "Fixture Index", "lookup Code Fixture — มี sheet Fixture_2026 หรือ Fixture*",    fixtureFile, setFixtureFile, ".xlsx,.xls")}
        </div>

        {/* Action buttons */}
        {status !== "processing" && (
          <div className="flex flex-col gap-2">
            <button onClick={handleProcess} disabled={!canProcess}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 rounded-xl font-bold text-sm
                bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-sm hover:shadow-md
                disabled:opacity-40 disabled:cursor-not-allowed transition-all">
              ⚡ Build New&amp;Renovate Report
            </button>
            {(status === "done" || status === "error" || status === "preview") && (
              <button onClick={handleReset}
                className="w-full px-4 py-3 rounded-xl font-semibold text-sm border border-pink-200
                  text-[#d41679] hover:bg-pink-50 transition-all">
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
              <div className="h-2.5 rounded-full bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]"
                style={{ width: `${pct}%`, transition: "width 0.5s ease" }} />
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

      {/* ── Right panel — always visible ── */}
      <div className="flex-1 min-w-0 flex flex-col">
        {isPreview && (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">

          {/* Panel header */}
          <div className="px-5 pt-5 pb-4 border-b border-slate-100">
            <h3 className="text-sm font-bold text-slate-700">ตรวจสอบข้อมูลก่อนดาวน์โหลด</h3>
            <p className="text-xs text-slate-400 mt-0.5">
              ตรวจสอบ comparison ระหว่าง AS IS และ TO BE แล้วกดยืนยันเพื่อ download Report
            </p>
          </div>

          {/* Tabs */}
          <div className="flex border-b border-slate-100 bg-slate-50/50">
            {([
              { key: "compare",   label: `Comparison (${previewData.rows.length.toLocaleString()} barcodes)` },
              { key: "unmatched", label: `Planogram ใน INDEX TO BE ไม่พบใน QRY (${previewData.unmatchedPlanograms.length})` },
            ] as const).map(({ key, label }) => (
              <button
                key={key}
                onClick={() => { setPreviewTab(key); clearSearch(); clearFilters(); }}
                className={`px-5 py-3 text-xs font-semibold border-b-2 -mb-px ${
                  previewTab === key
                    ? "border-[#E91E8C] text-[#E91E8C] bg-white"
                    : "border-transparent text-slate-400 hover:text-slate-600"
                }`}
              >
                {label}
              </button>
            ))}
          </div>

          {/* ── Comparison tab ──────────────────────────────────────────────────── */}
          {previewTab === "compare" && (
            <>
              {/* Status filter bar — multi-select */}
              <div className="flex items-center gap-1.5 px-5 py-3 border-b border-slate-100 bg-slate-50/30">
                <span className="text-[10px] text-slate-400 font-medium mr-1 whitespace-nowrap">Filter:</span>
                {/* ALL — clears all selections */}
                <button
                  onClick={clearFilters}
                  className={`text-[10px] px-2.5 py-1 rounded-full border font-semibold whitespace-nowrap ${
                    activeFilters.size === 0
                      ? "bg-slate-700 text-white border-slate-700"
                      : "bg-white text-slate-400 border-slate-200 hover:border-slate-300 hover:text-slate-600"
                  }`}
                >
                  ALL
                </button>
                {([
                  { f: "EXISTING"   as StatusKey, activeCls: "bg-emerald-600 text-white border-emerald-600" },
                  { f: "NEW EXPAND" as StatusKey, activeCls: "bg-blue-600 text-white border-blue-600" },
                  { f: "DELETE"     as StatusKey, activeCls: "bg-red-500 text-white border-red-500" },
                ]).map(({ f, activeCls }) => {
                  const active = activeFilters.has(f);
                  return (
                    <button
                      key={f}
                      onClick={() => toggleFilter(f)}
                      className={`text-[10px] px-2.5 py-1 rounded-full border font-semibold whitespace-nowrap flex items-center gap-1 ${
                        active ? activeCls : "bg-white text-slate-400 border-slate-200 hover:border-slate-300 hover:text-slate-600"
                      }`}
                    >
                      {active && <span className="text-[9px] leading-none">✓</span>}
                      {f}
                    </button>
                  );
                })}
                <span className="ml-auto flex items-center gap-2 text-[10px] text-slate-400 tabular-nums whitespace-nowrap">
                  {isPending && <span className="inline-block w-3 h-3 rounded-full border-2 border-pink-200 border-t-pink-400 animate-spin" />}
                  {filteredRows.length.toLocaleString()} barcodes
                  {hasSearch && (
                    <button onClick={clearSearch} className="text-pink-400 hover:text-pink-600 font-semibold">
                      ล้าง search
                    </button>
                  )}
                </span>
              </div>

              {/* Table */}
              <div className="overflow-auto" style={{ maxHeight: "62vh" }}>
                <table className="w-full text-xs border-collapse">
                  <thead>
                    {/* Column labels — sticky row 1 */}
                    <tr className="bg-slate-50 text-slate-500 font-semibold text-[11px]">
                      <th className="px-3 py-2.5 text-left border-b border-slate-200 sticky top-0 bg-slate-50 z-20 w-36">BARCODE</th>
                      <th className="px-3 py-2.5 text-left border-b border-slate-200 sticky top-0 bg-slate-50 z-20 w-64">ชื่อสินค้า (NAME)</th>
                      <th className="px-3 py-2.5 text-left border-b border-slate-200 sticky top-0 bg-slate-50 z-20 w-56">PLANOGRAM NAME</th>
                      <th className="px-3 py-2.5 text-left border-b border-slate-200 sticky top-0 bg-slate-50 z-20 w-32">AS IS Store</th>
                      <th className="px-3 py-2.5 text-left border-b border-slate-200 sticky top-0 bg-slate-50 z-20 w-32">TO BE Store</th>
                      <th className="px-3 py-2.5 text-left border-b border-slate-200 sticky top-0 bg-slate-50 z-20 w-28">STATUS</th>
                    </tr>
                    {/* Search inputs — sticky row 2 */}
                    <tr className="bg-white border-b border-slate-100" style={{ top: 37 }}>
                      <th className="px-3 pb-2 pt-1 sticky bg-white z-20" style={{ top: 37 }}>
                        <SearchInput value={srchBarcode}   onChange={applySearch(setSrchBarcode)}   placeholder="ค้นหา barcode…" />
                      </th>
                      <th className="px-3 pb-2 pt-1 sticky bg-white z-20" style={{ top: 37 }}>
                        <SearchInput value={srchName}      onChange={applySearch(setSrchName)}      placeholder="ค้นหาชื่อสินค้า…" />
                      </th>
                      <th className="px-3 pb-2 pt-1 sticky bg-white z-20" style={{ top: 37 }}>
                        <SearchInput value={srchPlanogram} onChange={applySearch(setSrchPlanogram)} placeholder="ค้นหา planogram…" />
                      </th>
                      <th className="px-3 pb-2 pt-1 sticky bg-white z-20" style={{ top: 37 }}>
                        <SearchInput value={srchStoreA}    onChange={applySearch(setSrchStoreA)}    placeholder="ค้นหา store…" />
                      </th>
                      <th className="px-3 pb-2 pt-1 sticky bg-white z-20" style={{ top: 37 }}>
                        <SearchInput value={srchStoreB}    onChange={applySearch(setSrchStoreB)}    placeholder="ค้นหา store…" />
                      </th>
                      <th className="px-3 pb-2 pt-1 sticky bg-white z-20" style={{ top: 37 }} />
                    </tr>
                  </thead>
                  <tbody className={isPending ? "opacity-60" : undefined}>
                    {displayedStoreRows.length === 0 && (
                      <tr>
                        <td colSpan={6} className="px-4 py-10 text-center text-slate-400 text-xs">
                          ไม่พบข้อมูลที่ตรงกับ filter / search
                        </td>
                      </tr>
                    )}
                    {displayedStoreRows.map(({ key, row, sr, idx, total, storeStatus }) => (
                      <tr
                        key={key}
                        className={`hover:bg-slate-100 ${idx === 0 ? "border-t-2 border-slate-200" : "border-t border-slate-100"}`}
                      >
                        {/* BARCODE — rowspan */}
                        {idx === 0 && (
                          <td rowSpan={total} className="px-3 py-2 font-mono text-slate-700 text-xs whitespace-nowrap align-top">
                            {row.barcode}
                          </td>
                        )}
                        {/* NAME — rowspan */}
                        {idx === 0 && (
                          <td rowSpan={total} className="px-3 py-2 text-slate-600 align-top text-xs leading-snug">
                            {row.name || <span className="text-slate-200">—</span>}
                          </td>
                        )}
                        {/* PLANOGRAM NAME — rowspan */}
                        {idx === 0 && (
                          <td rowSpan={total} className="px-3 py-2 text-slate-600 align-top">
                            <span className="block text-xs leading-snug" title={row.planograms.join(", ")}>
                              {row.planograms.join(", ")}
                            </span>
                          </td>
                        )}
                        {/* AS IS store — red tint if this store is being deleted */}
                        <td className={`px-3 py-1.5 tabular-nums text-xs ${storeStatus === "del" ? "bg-red-50" : ""}`}>
                          {sr.storeA
                            ? <span className={`font-medium ${storeStatus === "del" ? "text-red-500" : "text-slate-700"}`}>{sr.storeA}</span>
                            : <span className="text-slate-200">—</span>}
                        </td>
                        {/* TO BE store — blue tint if this store is new */}
                        <td className={`px-3 py-1.5 tabular-nums text-xs ${storeStatus === "new" ? "bg-blue-50" : ""}`}>
                          {sr.storeB
                            ? <span className={`font-medium ${storeStatus === "new" ? "text-blue-600" : "text-slate-700"}`}>
                                {sr.storeB}
                                {storeStatus === "new" && (
                                  <span className="ml-1.5 text-[9px] font-bold text-blue-500 bg-blue-100 px-1 py-0.5 rounded">NEW</span>
                                )}
                              </span>
                            : <span className="text-slate-200">—</span>}
                        </td>
                        {/* STATUS — rowspan */}
                        {idx === 0 && (
                          <td rowSpan={total} className="px-3 py-2 align-top">
                            <StatusPill status={row.status} />
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>

                {/* Load more / total indicator */}
                {filteredRows.length > displayLimit && (
                  <div className="px-5 py-3 flex items-center justify-between border-t border-slate-100 bg-slate-50/50">
                    <span className="text-[11px] text-slate-400 tabular-nums">
                      แสดง {displayLimit.toLocaleString()} จาก {filteredRows.length.toLocaleString()} barcodes
                    </span>
                    <button
                      onClick={() => setDisplayLimit(l => l + DISPLAY_LIMIT_STEP)}
                      className="text-[11px] px-3 py-1.5 rounded-lg border border-pink-200 text-pink-500 hover:bg-pink-50 font-semibold"
                    >
                      โหลดเพิ่ม {DISPLAY_LIMIT_STEP} รายการ
                    </button>
                  </div>
                )}
                {filteredRows.length > 0 && filteredRows.length <= displayLimit && (
                  <div className="px-5 py-2 text-center text-[11px] text-slate-300 border-t border-slate-100">
                    แสดงครบทุก {filteredRows.length.toLocaleString()} barcodes
                  </div>
                )}
              </div>
            </>
          )}

          {/* ── Unmatched tab ──────────────────────────────────────────────────── */}
          {previewTab === "unmatched" && (
            <div className="p-5 overflow-auto" style={{ maxHeight: "60vh" }}>
              {previewData.unmatchedPlanograms.length === 0 ? (
                <div className="flex items-center gap-2 text-emerald-600 text-sm font-medium">
                  <span>✓</span> PLANOGRAM ทุกตัวใน INDEX TO BE พบในไฟล์ QRY ครบถ้วน
                </div>
              ) : (
                <>
                  <div className="flex items-center justify-between mb-3">
                    <p className="text-xs text-slate-500">
                      PLANOGRAM NAME เหล่านี้มี TO BE Store ใน INDEX แต่ไม่พบข้อมูลใน QRY — สินค้าจะไม่ถูก fill ใน report สำหรับ planogram เหล่านี้
                    </p>
                    <button
                      onClick={handleExportUnmatched}
                      className="flex-shrink-0 ml-4 flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[#217346] text-[#217346] hover:bg-green-50 text-xs font-medium whitespace-nowrap"
                    >
                      <svg viewBox="0 0 20 20" className="w-3.5 h-3.5 flex-shrink-0" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <rect width="20" height="20" rx="3" fill="#217346"/>
                        <path d="M4 5.5L7.5 10L4 14.5H6.5L9 10.75L11.5 14.5H14L10.5 10L14 5.5H11.5L9 9.25L6.5 5.5H4Z" fill="white"/>
                        <rect x="14.5" y="5.5" width="1.5" height="9" fill="white"/>
                      </svg>
                      Export Excel
                    </button>
                  </div>
                  <div className="grid grid-cols-3 gap-1.5">
                    {previewData.unmatchedPlanograms.map(p => (
                      <span key={p} className="text-xs px-3 py-1.5 bg-amber-50 border border-amber-200 text-amber-700 rounded-lg truncate" title={p}>
                        {p}
                      </span>
                    ))}
                  </div>
                </>
              )}
            </div>
          )}

          {/* Confirm / downloaded footer */}
          <div className="px-5 py-4 border-t border-slate-100 bg-slate-50/30">
            {downloaded ? (
              <div className="flex items-center gap-3">
                <div className="flex-1 flex items-center gap-2 px-5 py-3.5 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-700 font-semibold text-sm">
                  <CheckCircle2 className="w-5 h-5 flex-shrink-0" />
                  ดาวน์โหลดเรียบร้อยแล้ว
                </div>
                <button
                  onClick={handleDownload}
                  className="flex items-center gap-1.5 px-4 py-3.5 rounded-xl border border-slate-200 text-slate-500 hover:bg-slate-50 text-sm font-medium whitespace-nowrap"
                >
                  <Download className="w-4 h-4" />
                  ดาวน์โหลดอีกครั้ง
                </button>
                <button
                  onClick={handleReset}
                  className="px-4 py-3.5 rounded-xl border border-pink-200 text-[#d41679] hover:bg-pink-50 text-sm font-semibold whitespace-nowrap"
                >
                  เริ่มใหม่
                </button>
              </div>
            ) : (
              <button
                onClick={handleConfirm}
                className="w-full flex items-center justify-center gap-2 px-6 py-4 rounded-xl font-bold text-base
                  bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-md hover:shadow-lg hover:scale-[1.01] transition-all"
              >
                <Download className="w-5 h-5" />
                ยืนยันและดาวน์โหลด New&amp;Renovate Report
              </button>
            )}
          </div>
        </div>
        )}

        {/* Done */}
        {status === "done" && stats && (
          <div className="space-y-4">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
            <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wide">ผลการ Lookup & Report</p>

            {/* Lookup counts */}
            <div className="grid grid-cols-5 gap-2 mb-3">
              {badge(stats.total,           "QRY rows (ตั้งต้น)",       "pink")}
              {badge(stats.matchedSpaceman, "Matched DATA_SPACEMAN",     "emerald")}
              {badge(stats.matchedMaster,   "Matched Master Assortment", "emerald")}
              {badge(stats.matchedIndex,    "Matched INDEX",             "emerald")}
              {badge(stats.matchedFixture,  "Matched Fixture Index",     "emerald")}
            </div>

            {/* Sheet row counts */}
            {(stats.s1Rows !== undefined || stats.s2Rows !== undefined || stats.s3Rows !== undefined) && (
              <div className="grid grid-cols-3 gap-2 mb-3">
                {stats.s1Rows !== undefined && badge(stats.s1Rows, "Rows → Sheet 1 (New&Existing)", "blue")}
                {stats.s2Rows !== undefined && badge(stats.s2Rows, "Rows → Sheet 2 (New for Link)", "blue")}
                {stats.s3Rows !== undefined && badge(stats.s3Rows, "Rows → Sheet 3 (Delete)", "blue")}
              </div>
            )}

            <div className="pt-3 border-t border-slate-100 grid grid-cols-2 gap-2 text-xs text-slate-400">
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
              bg-gradient-to-r from-[#E91E8C] to-[#F15A22] text-white shadow-md hover:shadow-lg hover:scale-[1.01] transition-all">
            <Download className="w-5 h-5" />
            ดาวน์โหลด New&amp;Renovate Report ที่เติมข้อมูลแล้ว
          </button>
          </div>
        )}

        {/* Placeholder when idle / processing / error */}
        {!isPreview && status !== "done" && (
          <div className="flex-1 min-h-[440px] flex flex-col items-center justify-center rounded-2xl bg-white border border-slate-200 shadow-sm text-slate-300 select-none gap-2">
            <FileSpreadsheet className="w-10 h-10 opacity-20" />
            <p className="text-sm font-medium">อัปโหลดไฟล์แล้วกด ⚡ Build เพื่อดู Preview</p>
          </div>
        )}

      </div>{/* end right panel */}
      </div>{/* end flex container */}
    </div>
  );
}
