"use client";

import { useRef, useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Download,
  RotateCcw,
  FileSpreadsheet,
  Zap,
  CheckCircle,
  XCircle,
  Clock,
  CloudOff,
  Trash2,
  Square,
  Loader2,
  ListOrdered,
  Plus,
  MinusCircle,
} from "lucide-react";

import StepIndicator from "@/components/StepIndicator";
import DropZone from "@/components/DropZone";
import ResultsTable from "@/components/ResultsTable";
import SpacemanMaster, {
  DriveFileInfo,
  formatDateTime,
} from "@/components/SpacemanMaster";
import { toDownloadRows, type DownloadRow } from "@/lib/download";
import {
  parseMissingRows,
  parseXlsbFiles,
  buildStructureLookup,
  parsePlanogramLookup,
  processRows,
  extractExistingValues,
} from "@/lib/processor";
import type { ProcessedRow } from "@/lib/types";

// ─── Types ─────────────────────────────────────────────────────────────────

const STEPS = [
  { id: 1, label: "RECAP" },
  { id: 2, label: "100 ช่อง" },
  { id: 3, label: "ตรวจสอบ" },
  { id: 4, label: "ดาวน์โหลด" },
];

const MAX_CONCURRENT = 2;

type Status = "idle" | "processing" | "done" | "error";
type AppView = "main" | "spaceman";
type JobStatus = "queued" | "processing" | "done" | "failed" | "terminated" | "downloaded";

interface BuildJob {
  id: string;
  label: string;
  status: JobStatus;
  createdAt: Date;
  startedAt?: Date;
  completedAt?: Date;
  progress: number;
  error?: string;
  buffer?: ArrayBuffer;
}

// ─── Home ──────────────────────────────────────────────────────────────────

export default function Home() {
  const [view, setView] = useState<AppView>("main");
  // Lazy-mount flag: SpacemanMaster is only added to the DOM on first visit
  const [spacemanMounted, setSpacemanMounted] = useState(false);
  const [step, setStep] = useState(1);
  const [status, setStatus] = useState<Status>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct] = useState(0);

  const [recapFiles, setRecapFiles] = useState<File[]>([]);
  const [xlsbFiles, setXlsbFiles] = useState<File[]>([]);
  const [driveFileInfo, setDriveFileInfo] = useState<DriveFileInfo | null>(null);
  const [driveLoading, setDriveLoading] = useState(true);
  const [results, setResults] = useState<ProcessedRow[]>([]);
  const [recapSuggestions, setRecapSuggestions] = useState<Partial<Record<string, string[]>>>({});

  // Queue state (display only — heavy data lives in refs)
  const [jobs, setJobs] = useState<BuildJob[]>([]);

  // Refs — not in React state to avoid re-render overhead and serialization issues
  const recapBufRef = useRef<ArrayBuffer | null>(null);
  const workersRef = useRef<Map<string, Worker>>(new Map());
  const jobDataRef = useRef<Map<string, { recapBuf: ArrayBuffer; rows: DownloadRow[] }>>(new Map());
  const jobCounterRef = useRef(0);

  // Terminate all workers on unmount
  useEffect(() => {
    return () => { workersRef.current.forEach((w) => w.terminate()); };
  }, []);

  // Fetch latest GDrive file on mount
  useEffect(() => {
    fetch("/api/spaceman/latest")
      .then((r) => r.json())
      .then((data) => setDriveFileInfo(data.file ?? null))
      .catch(() => setDriveFileInfo(null))
      .finally(() => setDriveLoading(false));
  }, []);

  // ── Core job starter (stable — only touches refs + functional setJobs) ──

  const startJobFn = useCallback((id: string, recapBuf: ArrayBuffer, rows: DownloadRow[]) => {
    setJobs((prev) =>
      prev.map((j) =>
        j.id === id ? { ...j, status: "processing", startedAt: new Date(), progress: 5 } : j
      )
    );

    const worker = new Worker(new URL("../lib/download.worker.ts", import.meta.url));
    workersRef.current.set(id, worker);

    worker.onmessage = (e: MessageEvent) => {
      const msg = e.data as { type: string; pct?: number; buffer?: ArrayBuffer; message?: string };
      switch (msg.type) {
        case "init_ok":
          worker.postMessage({ type: "build", rows });
          break;
        case "progress":
          setJobs((prev) =>
            prev.map((j) => (j.id === id ? { ...j, progress: msg.pct ?? j.progress } : j))
          );
          break;
        case "done":
          workersRef.current.delete(id);
          jobDataRef.current.delete(id);
          setJobs((prev) =>
            prev.map((j) =>
              j.id === id
                ? { ...j, status: "done", progress: 100, completedAt: new Date(), buffer: msg.buffer }
                : j
            )
          );
          break;
        case "error":
          workersRef.current.delete(id);
          setJobs((prev) =>
            prev.map((j) =>
              j.id === id
                ? { ...j, status: "failed", error: msg.message ?? "Worker error", completedAt: new Date() }
                : j
            )
          );
          break;
      }
    };

    worker.onerror = (e: ErrorEvent) => {
      workersRef.current.delete(id);
      setJobs((prev) =>
        prev.map((j) =>
          j.id === id
            ? { ...j, status: "failed", error: e.message ?? "Worker crashed", completedAt: new Date() }
            : j
        )
      );
    };

    // Transfer buffer to avoid a full copy (slice first to preserve original)
    const buf = recapBuf.slice(0);
    worker.postMessage({ type: "init", buffer: buf }, [buf]);
  }, []); // stable — no external deps

  // ── Auto-start queued jobs when a slot opens ────────────────────────────

  useEffect(() => {
    const running = jobs.filter((j) => j.status === "processing").length;
    if (running >= MAX_CONCURRENT) return;
    const next = jobs.find((j) => j.status === "queued");
    if (!next) return;
    const data = jobDataRef.current.get(next.id);
    if (!data) return;
    startJobFn(next.id, data.recapBuf, data.rows);
  }, [jobs, startJobFn]);

  // ── Queue actions ───────────────────────────────────────────────────────

  const enqueueJob = () => {
    if (!recapBufRef.current || results.length === 0) return;
    const id = crypto.randomUUID();
    const num = ++jobCounterRef.current;
    const baseName = recapFiles[0]?.name.replace(/\.[^.]+$/, "") ?? "RECAP";

    jobDataRef.current.set(id, {
      recapBuf: recapBufRef.current.slice(0),
      rows: toDownloadRows(results),
    });

    setJobs((prev) => [
      ...prev,
      { id, label: `${baseName}_filled_#${num}.xlsx`, status: "queued", createdAt: new Date(), progress: 0 },
    ]);
    setStep(4);
  };

  const terminateJob = (id: string) => {
    workersRef.current.get(id)?.terminate();
    workersRef.current.delete(id);
    setJobs((prev) =>
      prev.map((j) =>
        j.id === id && (j.status === "processing" || j.status === "queued")
          ? { ...j, status: "terminated", completedAt: new Date() }
          : j
      )
    );
  };

  const removeJob = (id: string) => {
    workersRef.current.get(id)?.terminate();
    workersRef.current.delete(id);
    jobDataRef.current.delete(id);
    setJobs((prev) => prev.filter((j) => j.id !== id));
  };

  const downloadJob = (id: string, label: string, buffer: ArrayBuffer) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = label;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    // Clear buffer to free memory; mark as downloaded
    setJobs((prev) =>
      prev.map((j) => (j.id === id ? { ...j, status: "downloaded", buffer: undefined } : j))
    );
  };

  // ── Main processing flow ────────────────────────────────────────────────

  const canProcess = () =>
    recapFiles.length === 1 && xlsbFiles.length > 0 && driveFileInfo !== null;

  const handleProcess = async () => {
    if (!recapFiles[0] || xlsbFiles.length === 0 || !driveFileInfo) return;

    setStatus("processing");
    setStep(3);
    setPct(0);

    try {
      setStatusMsg("อ่านไฟล์ RECAP...");
      setPct(10);
      const recapBuf = await recapFiles[0].arrayBuffer();
      recapBufRef.current = recapBuf.slice(0);
      const wb = XLSX.read(recapBuf, { type: "array" });
      const missing = parseMissingRows(wb);
      setRecapSuggestions(extractExistingValues(wb));

      setStatusMsg(`พบ ${missing.length} รายการที่ต้องเติมข้อมูล — กำลังค้นหาในไฟล์ 100 ช่อง...`);
      setPct(25);

      const [barcodeMap, structureMap] = await Promise.all([
        parseXlsbFiles(xlsbFiles),
        buildStructureLookup(xlsbFiles),
      ]);

      setStatusMsg("กำลังตรวจสอบข้อมูลจาก DATA_SPACEMAN...");
      setPct(45);
      const res = await fetch(`/api/spaceman/file?id=${driveFileInfo.id}`);
      if (!res.ok) throw new Error("ไม่สามารถดาวน์โหลดไฟล์ DATA_SPACEMAN จาก Google Drive ได้");
      const buf = await res.arrayBuffer();
      const spacemanFile = new File([buf], driveFileInfo.name, {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      setStatusMsg("อ่าน DATA_SPACEMAN เพื่อหา PLANOGRAM...");
      setPct(50);
      const planogramMap = await parsePlanogramLookup(spacemanFile, (p) =>
        setPct(50 + p * 0.35)
      );

      setStatusMsg("ประมวลผลข้อมูล...");
      setPct(90);
      const processed = processRows(missing, barcodeMap, structureMap, planogramMap);
      setResults(processed);

      setPct(100);
      setStatusMsg("เสร็จสิ้น!");
      setStatus("done");
    } catch (err) {
      setStatus("error");
      setStatusMsg(String(err));
    }
  };

  const handleResultsChange = (updated: ProcessedRow[]) => {
    setResults(updated);
  };

  const reset = () => {
    // Jobs persist across resets — do NOT clear them
    recapBufRef.current = null;
    setStep(1);
    setStatus("idle");
    setStatusMsg("");
    setPct(0);
    setRecapFiles([]);
    setXlsbFiles([]);
    setResults([]);
    setRecapSuggestions({});
  };

  const confirmed = results.filter((r) => r.confidence === "confirmed").length;
  const inferred = results.filter((r) => r.confidence === "inferred").length;
  const notFound = results.filter((r) => r.confidence === "not_found").length;

  // ── Render ──────────────────────────────────────────────────────────────

  return (
    <main className="min-h-screen bg-slate-50">
      {/* Header */}
      <div className="bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100] text-white px-6 py-4 shadow-lg">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="bg-white rounded-xl px-3 py-2 shadow-sm flex items-center">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src="/mini-bigc-logo.png" alt="Mini BigC" className="h-9 w-auto object-contain" />
            </div>
            <div className="border-l-2 border-white/40 pl-4">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-5 h-5 text-white/90" />
                <h1 className="text-lg font-bold tracking-tight">RECAP Auto-Filler</h1>
              </div>
              <p className="text-white/80 text-xs mt-0.5">
                เติมข้อมูล DIVISION / DEPT / SUB-DEPT / Class / PLANOGRAM อัตโนมัติ
              </p>
            </div>
          </div>

          {view === "main" && step > 1 && (
            <button
              onClick={reset}
              className="flex items-center gap-2 px-4 py-2 bg-white/20 hover:bg-white/30 rounded-xl text-sm font-medium transition-colors border border-white/30"
            >
              <RotateCcw className="w-4 h-4" />
              เริ่มใหม่
            </button>
          )}
        </div>
      </div>

      {/* Tab Navigation */}
      <div className="bg-white border-b border-slate-200 shadow-sm px-6">
        <div className="flex gap-0">
          <TabBtn active={view === "main"} onClick={() => setView("main")}>
            <FileSpreadsheet className="w-4 h-4" />
            อัปโหลดข้อมูล
          </TabBtn>
          <TabBtn active={view === "spaceman"} onClick={() => { setView("spaceman"); setSpacemanMounted(true); }}>
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
              <ellipse cx="12" cy="5" rx="9" ry="3" />
              <path d="M3 5v14c0 1.66 4.03 3 9 3s9-1.34 9-3V5" />
              <path d="M3 12c0 1.66 4.03 3 9 3s9-1.34 9-3" />
            </svg>
            DATA_SPACEMAN
          </TabBtn>
        </div>
      </div>

      {/* Main layout — flex with sticky queue panel on the right */}
      <div className="px-6 py-8">
        <div className="flex gap-6 items-start">

          {/* ── Content area ─────────────────────────────────────────── */}
          <div className="flex-1 min-w-0 space-y-8">

            {/* DATA_SPACEMAN — lazy-mount on first visit, then kept in DOM (hidden) to preserve parsed data */}
            {spacemanMounted && (
              <div className={view === "spaceman" ? "" : "hidden"}>
                <SpacemanMaster
                  onFileInfoChange={(info) => {
                    setDriveFileInfo(info);
                    setDriveLoading(false);
                  }}
                />
              </div>
            )}

            {/* Main upload flow */}
            {view === "main" && (
              <>
                <StepIndicator steps={STEPS} current={step} />

                {/* Step 1 — Upload RECAP */}
                {step === 1 && (
                  <Card title="Step 1 - อัปโหลดไฟล์ RECAP">
                    <DropZone
                      label="ไฟล์ RECAP.xlsx"
                      accept=".xlsx,.xls"
                      files={recapFiles}
                      onFiles={setRecapFiles}
                      hint="ไฟล์ที่ต้องการเติมข้อมูลในคอลัมน์ F-J ของ Sheet 'NEW SCM'"
                    />
                    <NavBtn onClick={() => setStep(2)} disabled={recapFiles.length !== 1}>
                      ถัดไป →
                    </NavBtn>
                  </Card>
                )}

                {/* Step 2 — Upload 100 ช่อง */}
                {step === 2 && (
                  <Card title="Step 2 - อัปโหลดไฟล์ 100 ช่อง (.xlsb)">
                    <DropZone
                      label="ไฟล์ 100 ช่อง (เลือกได้หลายไฟล์)"
                      accept=".xlsb,.xlsx,.xls"
                      multiple
                      files={xlsbFiles}
                      onFiles={setXlsbFiles}
                      hint="7_2_10_SNACKS, 7_2_50_CONFECTIONARY, 7_2_60_BISCUITS, 7_2_60_WINE ฯลฯ"
                    />

                    {/* DATA_SPACEMAN GDrive status */}
                    <div
                      className={`rounded-xl border px-4 py-3 flex items-center gap-3 text-sm ${
                        driveLoading
                          ? "bg-slate-50 border-slate-200 text-slate-500"
                          : driveFileInfo
                            ? "bg-green-50 border-green-200 text-green-800"
                            : "bg-amber-50 border-amber-200 text-amber-800"
                      }`}
                    >
                      {driveLoading ? (
                        <>
                          <div className="animate-spin rounded-full h-4 w-4 border-2 border-slate-300 border-t-slate-500 flex-shrink-0" />
                          <span>กำลังตรวจสอบ DATA_SPACEMAN ใน Google Drive...</span>
                        </>
                      ) : driveFileInfo ? (
                        <>
                          <CheckCircle className="w-4 h-4 text-green-500 flex-shrink-0" />
                          <div>
                            <span className="font-medium">DATA_SPACEMAN พร้อมใช้งาน</span>
                            <span className="text-green-600 ml-2">—</span>
                            <span className="text-green-700 ml-2 flex items-center gap-1 inline-flex">
                              <Clock className="w-3 h-3" />
                              อัปโหลดล่าสุด: <strong>{formatDateTime(driveFileInfo.createdTime)}</strong>
                            </span>
                          </div>
                        </>
                      ) : (
                        <>
                          <CloudOff className="w-4 h-4 text-amber-500 flex-shrink-0" />
                          <div>
                            <span className="font-medium">ไม่พบไฟล์ DATA_SPACEMAN ใน Google Drive</span>
                            <span className="text-amber-600 ml-2 text-xs">
                              — กรุณาไปที่แท็บ &quot;DATA_SPACEMAN&quot; เพื่ออัปโหลดไฟล์ก่อน
                            </span>
                          </div>
                        </>
                      )}
                    </div>

                    <div className="flex gap-3">
                      <NavBtn variant="outline" onClick={() => setStep(1)}>← ย้อนกลับ</NavBtn>
                      <NavBtn onClick={handleProcess} disabled={!canProcess()}>
                        <Zap className="w-4 h-4" />
                        ประมวลผลทันที
                      </NavBtn>
                    </div>
                  </Card>
                )}

                {/* Step 3 — Review & enqueue */}
                {step === 3 && (
                  <Card title="Step 3 - ตรวจสอบผลลัพธ์">
                    {status === "processing" && (
                      <div className="space-y-4 py-8">
                        <div className="flex items-center justify-center">
                          <div className="animate-spin rounded-full h-10 w-10 border-4 border-pink-200 border-t-[#E91E8C]" />
                        </div>
                        <p className="text-center text-slate-600 font-medium">{statusMsg}</p>
                        <div className="w-full bg-slate-200 rounded-full h-3 overflow-hidden">
                          <div
                            className="h-3 rounded-full transition-all duration-500 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]"
                            style={{ width: `${pct}%` }}
                          />
                        </div>
                        <p className="text-center text-sm text-slate-400">{pct}%</p>
                      </div>
                    )}

                    {status === "error" && (
                      <div className="bg-red-50 border border-red-200 rounded-xl p-5 text-red-700">
                        ❌ เกิดข้อผิดพลาด: {statusMsg}
                      </div>
                    )}

                    {status === "done" && (
                      <>
                        <div className="grid grid-cols-3 gap-4 mb-6">
                          <StatCard label="ยืนยันแล้ว" value={confirmed} color="green" />
                          <StatCard label="อนุมาน" value={inferred} color="amber" />
                          <StatCard label="ไม่พบ / กรอกเอง" value={notFound} color="red" />
                        </div>
                        <ResultsTable rows={results} onChange={handleResultsChange} externalSuggestions={recapSuggestions} />
                        <div className="flex gap-3 pt-4 border-t border-slate-100">
                          <NavBtn onClick={enqueueJob}>
                            <Plus className="w-4 h-4" />
                            เพิ่มเข้าคิว Build File
                          </NavBtn>
                        </div>
                      </>
                    )}
                  </Card>
                )}

                {/* Step 4 — Queued confirmation */}
                {step === 4 && (
                  <Card title="เพิ่มเข้าคิวสำเร็จ!">
                    <div className="text-center py-10 space-y-4">
                      <div className="flex justify-center">
                        <div className="rounded-full bg-green-100 p-4">
                          <CheckCircle className="w-12 h-12 text-green-500" />
                        </div>
                      </div>
                      <p className="text-xl font-semibold text-slate-700">เพิ่มเข้าคิว Build เรียบร้อย</p>
                      <p className="text-slate-500 text-sm">
                        ไฟล์กำลังถูก Build อยู่เบื้องหลัง
                        <br />
                        ตรวจสอบสถานะและดาวน์โหลดได้ที่{" "}
                        <span className="font-semibold text-[#E91E8C]">แผงคิวด้านขวามือ</span>
                      </p>
                      <button
                        onClick={reset}
                        className="mt-4 px-6 py-3 text-white rounded-xl font-semibold transition-all shadow-md hover:shadow-lg hover:scale-[1.02] bg-gradient-to-r from-[#E91E8C] to-[#F15A22]"
                      >
                        เริ่มใหม่อีกครั้ง
                      </button>
                    </div>
                  </Card>
                )}
              </>
            )}
          </div>

          {/* ── Queue panel (sticky, right side) — main tab only ────── */}
          {view === "main" && jobs.length > 0 && (
            <div className="w-72 flex-shrink-0 sticky top-4">
              <JobQueuePanel
                jobs={jobs}
                onTerminate={terminateJob}
                onRemove={removeJob}
                onDownload={downloadJob}
              />
            </div>
          )}
        </div>
      </div>
    </main>
  );
}

// ─── Job Queue Panel ───────────────────────────────────────────────────────

function JobQueuePanel({
  jobs,
  onTerminate,
  onRemove,
  onDownload,
}: {
  jobs: BuildJob[];
  onTerminate: (id: string) => void;
  onRemove: (id: string) => void;
  onDownload: (id: string, label: string, buffer: ArrayBuffer) => void;
}) {
  const activeCount = jobs.filter(
    (j) => j.status === "queued" || j.status === "processing"
  ).length;

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
      <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]" />
      <div className="px-4 py-3 border-b border-pink-50 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <ListOrdered className="w-4 h-4 text-[#E91E8C]" />
          <h3 className="font-bold text-slate-800 text-sm">คิว Build ไฟล์</h3>
        </div>
        {activeCount > 0 && (
          <span className="bg-pink-100 text-[#E91E8C] text-xs font-bold px-2 py-0.5 rounded-full">
            {activeCount} รายการ
          </span>
        )}
      </div>
      <div className="divide-y divide-slate-100 max-h-[70vh] overflow-y-auto">
        {jobs.map((job) => (
          <JobItem
            key={job.id}
            job={job}
            onTerminate={onTerminate}
            onRemove={onRemove}
            onDownload={onDownload}
          />
        ))}
      </div>
    </div>
  );
}

// ─── Job Item ──────────────────────────────────────────────────────────────

function JobItem({
  job,
  onTerminate,
  onRemove,
  onDownload,
}: {
  job: BuildJob;
  onTerminate: (id: string) => void;
  onRemove: (id: string) => void;
  onDownload: (id: string, label: string, buffer: ArrayBuffer) => void;
}) {
  const fmtTime = (d: Date) =>
    d.toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", second: "2-digit" });

  const statusCfg: Record<JobStatus, { label: string; cls: string; icon: React.ReactNode }> = {
    queued:     { label: "รอคิว",           cls: "text-slate-600 bg-slate-100",  icon: <Clock className="w-3 h-3" /> },
    processing: { label: "กำลัง Build",      cls: "text-blue-700 bg-blue-100",   icon: <Loader2 className="w-3 h-3 animate-spin" /> },
    done:       { label: "พร้อมโหลด",       cls: "text-green-700 bg-green-100", icon: <CheckCircle className="w-3 h-3" /> },
    downloaded: { label: "โหลดแล้ว",        cls: "text-slate-400 bg-slate-100", icon: <MinusCircle className="w-3 h-3" /> },
    failed:     { label: "ล้มเหลว",          cls: "text-red-700 bg-red-100",     icon: <XCircle className="w-3 h-3" /> },
    terminated: { label: "ยกเลิกแล้ว",       cls: "text-slate-400 bg-slate-100", icon: <MinusCircle className="w-3 h-3" /> },
  };

  const cfg = statusCfg[job.status];
  const canStop = job.status === "queued" || job.status === "processing";
  const canRemove = !canStop;

  return (
    <div className="px-4 py-3 space-y-2">
      {/* Label + status badge */}
      <div className="flex items-start justify-between gap-2">
        <p className="text-xs font-semibold text-slate-700 truncate flex-1" title={job.label}>
          {job.label}
        </p>
        <span
          className={`flex items-center gap-1 text-xs font-medium px-1.5 py-0.5 rounded-full flex-shrink-0 ${cfg.cls}`}
        >
          {cfg.icon}
          {cfg.label}
        </span>
      </div>

      {/* Progress bar (processing only) */}
      {job.status === "processing" && (
        <div className="space-y-1">
          <div className="w-full bg-slate-200 rounded-full h-1.5 overflow-hidden">
            <div
              className="h-1.5 rounded-full transition-all duration-500 bg-gradient-to-r from-[#E91E8C] to-[#F15A22]"
              style={{ width: `${job.progress}%` }}
            />
          </div>
          <p className="text-xs text-slate-400">{job.progress}%</p>
        </div>
      )}

      {/* Error message */}
      {job.status === "failed" && job.error && (
        <p className="text-xs text-red-500 bg-red-50 rounded-lg p-2 break-words leading-relaxed">
          {job.error}
        </p>
      )}

      {/* Time info + action buttons */}
      <div className="flex items-end justify-between gap-2">
        <div className="text-xs text-slate-400 space-y-0.5">
          {job.startedAt ? (
            <div>เริ่ม {fmtTime(job.startedAt)}</div>
          ) : (
            <div>สร้างเมื่อ {fmtTime(job.createdAt)}</div>
          )}
          {job.completedAt && (
            <div>เสร็จ {fmtTime(job.completedAt)}</div>
          )}
        </div>

        <div className="flex gap-1 flex-shrink-0">
          {/* Download — only when done and buffer exists */}
          {job.status === "done" && job.buffer && (
            <button
              onClick={() => onDownload(job.id, job.label, job.buffer!)}
              className="flex items-center gap-1 px-2 py-1 rounded-lg text-xs font-medium bg-green-600 text-white hover:bg-green-700 transition-colors"
            >
              <Download className="w-3 h-3" />
              โหลด
            </button>
          )}

          {/* Stop — queued or processing */}
          {canStop && (
            <button
              onClick={() => onTerminate(job.id)}
              title="หยุด / ยกเลิก"
              className="flex items-center gap-1 px-2 py-1 rounded-lg text-xs font-medium bg-red-100 text-red-700 hover:bg-red-200 transition-colors"
            >
              <Square className="w-3 h-3" />
              หยุด
            </button>
          )}

          {/* Remove — finished states */}
          {canRemove && (
            <button
              onClick={() => onRemove(job.id)}
              title="ลบออกจากรายการ"
              className="flex items-center gap-1 px-2 py-1 rounded-lg text-xs font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition-colors"
            >
              <Trash2 className="w-3 h-3" />
              ลบ
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Shared sub-components ─────────────────────────────────────────────────

function TabBtn({
  children,
  active,
  onClick,
}: {
  children: React.ReactNode;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className={`
        flex items-center gap-2 px-5 py-3.5 text-sm font-semibold border-b-2 transition-all
        ${
          active
            ? "border-[#E91E8C] text-[#E91E8C]"
            : "border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300"
        }
      `}
    >
      {children}
    </button>
  );
}

function Card({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
      <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#00A6E2] via-[#FFD100] via-[#F15A22] to-[#72BF44]" />
      <div className="px-6 py-4 border-b border-pink-50 flex items-center gap-3">
        <div className="w-1 h-5 rounded-full bg-gradient-to-b from-[#E91E8C] to-[#F15A22]" />
        <h2 className="font-bold text-slate-800 text-lg">{title}</h2>
      </div>
      <div className="p-6 space-y-6">{children}</div>
    </div>
  );
}

function NavBtn({
  children,
  onClick,
  disabled,
  variant = "primary",
}: {
  children: React.ReactNode;
  onClick: () => void;
  disabled?: boolean;
  variant?: "primary" | "outline";
}) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`
        flex items-center gap-2 px-6 py-3 rounded-xl font-semibold text-sm
        transition-all duration-200 disabled:opacity-40 disabled:cursor-not-allowed
        ${
          variant === "primary"
            ? "bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white shadow-sm hover:shadow-md hover:from-[#d41679] hover:to-[#be185d]"
            : "border border-pink-200 text-[#d41679] hover:bg-pink-50"
        }
      `}
    >
      {children}
    </button>
  );
}

function StatCard({
  label,
  value,
  color,
}: {
  label: string;
  value: number;
  color: "green" | "amber" | "red";
}) {
  const colors = {
    green: "bg-green-50 border-green-200 text-green-700",
    amber: "bg-amber-50 border-amber-200 text-amber-700",
    red: "bg-red-50 border-red-200 text-red-700",
  };
  return (
    <div className={`${colors[color]} border rounded-xl p-4 text-center`}>
      <div className="text-3xl font-bold">{value}</div>
      <div className="text-sm font-medium mt-1">{label}</div>
    </div>
  );
}
