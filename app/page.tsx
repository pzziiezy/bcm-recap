"use client";

import { useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  CheckCircle,
  Download,
  FileSpreadsheet,
  RotateCcw,
  XCircle,
  Zap,
} from "lucide-react";

import DropZone from "@/components/DropZone";
import ResultsTable from "@/components/ResultsTable";
import StepIndicator from "@/components/StepIndicator";
import { toDownloadRows } from "@/lib/download";
import {
  buildXlsbLookups,
  parseMissingRows,
  parsePlanogramLookup,
  processRows,
} from "@/lib/processor";
import type { ProcessedRow } from "@/lib/types";

const STEPS = [
  { id: 1, label: "RECAP" },
  { id: 2, label: "100 ช่อง" },
  { id: 3, label: "DATA_SPACEMAN" },
  { id: 4, label: "ตรวจสอบ" },
  { id: 5, label: "ดาวน์โหลด" },
];

type Status = "idle" | "processing" | "done" | "error";
type ModalState =
  | { type: "hidden" }
  | { type: "loading" }
  | { type: "success" }
  | { type: "error"; message: string };

type PrebuildWorkerMessage =
  | { type: "init"; ok: boolean }
  | { type: "build"; ok: boolean; jobId: number; buffer?: ArrayBuffer; error?: string };

export default function Home() {
  const [step, setStep] = useState(1);
  const [status, setStatus] = useState<Status>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct] = useState(0);

  const [recapFiles, setRecapFiles] = useState<File[]>([]);
  const [xlsbFiles, setXlsbFiles] = useState<File[]>([]);
  const [spacemanFiles, setSpacemanFiles] = useState<File[]>([]);

  const [results, setResults] = useState<ProcessedRow[]>([]);
  const [modal, setModal] = useState<ModalState>({ type: "hidden" });

  const recapBufRef = useRef<ArrayBuffer | null>(null);
  const prebuildRef = useRef<{
    worker: Worker | null;
    ready: Promise<Worker | null> | null;
    promise: Promise<Blob | null>;
    jobId: number;
  }>({
    worker: null,
    ready: null,
    promise: Promise.resolve(null),
    jobId: 0,
  });

  const ensurePrebuildWorker = async (): Promise<Worker | null> => {
    if (prebuildRef.current.ready) return prebuildRef.current.ready;
    if (!recapBufRef.current) return null;

    const worker = new Worker(
      new URL("../lib/download.worker.ts", import.meta.url)
    );
    prebuildRef.current.worker = worker;

    const ready = new Promise<Worker | null>((resolve) => {
      const handleMessage = (event: MessageEvent<PrebuildWorkerMessage>) => {
        if (event.data.type !== "init") return;

        worker.removeEventListener("message", handleMessage);
        worker.removeEventListener("error", handleError);

        if (event.data.ok) {
          resolve(worker);
          return;
        }

        worker.terminate();
        prebuildRef.current.worker = null;
        resolve(null);
      };

      const handleError = () => {
        worker.removeEventListener("message", handleMessage);
        worker.removeEventListener("error", handleError);
        worker.terminate();
        prebuildRef.current.worker = null;
        prebuildRef.current.ready = null;
        resolve(null);
      };

      worker.addEventListener("message", handleMessage);
      worker.addEventListener("error", handleError);

      const initBuffer = recapBufRef.current?.slice(0);
      if (!initBuffer) {
        handleError();
        return;
      }

      worker.postMessage({ type: "init", buffer: initBuffer }, [initBuffer]);
    });

    prebuildRef.current.ready = ready;
    return ready;
  };

  const startPrebuild = (rows: ProcessedRow[]) => {
    if (!recapBufRef.current) {
      prebuildRef.current.promise = Promise.resolve(null);
      return;
    }

    const jobId = ++prebuildRef.current.jobId;
    const downloadRows = toDownloadRows(rows);

    prebuildRef.current.promise = ensurePrebuildWorker().then((worker) => {
      if (!worker) return null;

      return new Promise<Blob | null>((resolve) => {
        const handleMessage = (event: MessageEvent<PrebuildWorkerMessage>) => {
          if (event.data.type !== "build" || event.data.jobId !== jobId) return;

          worker.removeEventListener("message", handleMessage);
          worker.removeEventListener("error", handleError);

          if (!event.data.ok || !event.data.buffer) {
            resolve(null);
            return;
          }

          resolve(
            new Blob([event.data.buffer], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            })
          );
        };

        const handleError = () => {
          worker.removeEventListener("message", handleMessage);
          worker.removeEventListener("error", handleError);
          worker.terminate();
          prebuildRef.current.worker = null;
          prebuildRef.current.ready = null;
          resolve(null);
        };

        worker.addEventListener("message", handleMessage);
        worker.addEventListener("error", handleError);
        worker.postMessage({ type: "build", jobId, rows: downloadRows });
      });
    });
  };

  const canNext = () => {
    if (step === 1) return recapFiles.length === 1;
    if (step === 2) return xlsbFiles.length > 0;
    if (step === 3) return spacemanFiles.length === 1;
    return true;
  };

  const handleProcess = async () => {
    if (!recapFiles[0] || xlsbFiles.length === 0 || !spacemanFiles[0]) return;

    setStatus("processing");
    setStep(4);
    setPct(0);

    try {
      setStatusMsg("อ่านไฟล์ RECAP...");
      setPct(10);

      const recapBuf = await recapFiles[0].arrayBuffer();
      recapBufRef.current = recapBuf.slice(0);

      const workbook = XLSX.read(recapBuf, { type: "array" });
      const missing = parseMissingRows(workbook);

      setStatusMsg(`พบ ${missing.length} รายการที่ต้องเติมข้อมูล - กำลังค้นหาในไฟล์ 100 ช่อง...`);
      setPct(25);

      const { barcodeMap, structureMap } = await buildXlsbLookups(xlsbFiles);

      setStatusMsg("อ่าน DATA_SPACEMAN เพื่อหา PLANOGRAM...");
      setPct(50);
      const planogramMap = await parsePlanogramLookup(spacemanFiles[0], (progress) =>
        setPct(50 + progress * 0.35)
      );

      setStatusMsg("ประมวลผลข้อมูล...");
      setPct(90);

      const processed = processRows(
        missing,
        barcodeMap,
        structureMap,
        planogramMap
      );

      setResults(processed);
      setPct(100);
      setStatusMsg("เสร็จสิ้น!");
      setStatus("done");
      setStep(4);

      startPrebuild(processed);
    } catch (error) {
      setStatus("error");
      setStatusMsg(String(error));
    }
  };

  const handleResultsChange = (updated: ProcessedRow[]) => {
    setResults(updated);
    startPrebuild(updated);
  };

  const handleDownload = async () => {
    setModal({ type: "loading" });

    try {
      const blob = await prebuildRef.current.promise;
      if (!blob) {
        throw new Error("ไม่สามารถสร้างไฟล์ได้ กรุณาลองใหม่");
      }

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "RECAP_filled.xlsx";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      setModal({ type: "success" });
      setStep(5);
    } catch (error) {
      setModal({ type: "error", message: String(error) });
    }
  };

  const reset = () => {
    prebuildRef.current.worker?.terminate();
    prebuildRef.current.worker = null;
    prebuildRef.current.ready = null;
    prebuildRef.current.promise = Promise.resolve(null);
    prebuildRef.current.jobId = 0;

    setStep(1);
    setStatus("idle");
    setStatusMsg("");
    setPct(0);
    setRecapFiles([]);
    setXlsbFiles([]);
    setSpacemanFiles([]);
    setResults([]);
    setModal({ type: "hidden" });

    recapBufRef.current = null;
  };

  const confirmed = results.filter((row) => row.confidence === "confirmed").length;
  const inferred = results.filter((row) => row.confidence === "inferred").length;
  const notFound = results.filter((row) => row.confidence === "not_found").length;

  return (
    <main className="min-h-screen">
      <div className="bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100] text-white px-6 py-4 shadow-lg">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="bg-white rounded-xl px-3 py-2 shadow-sm flex items-center">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img
                src="/mini-bigc-logo.png"
                alt="Mini BigC"
                className="h-9 w-auto object-contain"
              />
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
          {step > 1 && (
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

      <div className="px-6 py-8 space-y-8">
        <StepIndicator steps={STEPS} current={step} />

        {step === 1 && (
          <Card title="Step 1 - อัปโหลดไฟล์ RECAP">
            <DropZone
              label="ไฟล์ RECAP.xlsx"
              accept=".xlsx,.xls"
              files={recapFiles}
              onFiles={setRecapFiles}
              hint="ไฟล์ที่ต้องการเติมข้อมูลในคอลัมน์ F-J ของ Sheet 'NEW SCM'"
            />
            <NavBtn onClick={() => setStep(2)} disabled={!canNext()}>
              ถัดไป →
            </NavBtn>
          </Card>
        )}

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
            <div className="flex gap-3">
              <NavBtn variant="outline" onClick={() => setStep(1)}>
                ← ย้อนกลับ
              </NavBtn>
              <NavBtn onClick={() => setStep(3)} disabled={!canNext()}>
                ถัดไป →
              </NavBtn>
            </div>
          </Card>
        )}

        {step === 3 && (
          <Card title="Step 3 - อัปโหลด DATA_SPACEMAN.xlsx">
            <DropZone
              label="DATA_SPACEMAN.xlsx"
              accept=".xlsx,.xls"
              files={spacemanFiles}
              onFiles={setSpacemanFiles}
              hint="ใช้หา PLANOGRAM - อัปโหลดใหม่ทุกสัปดาห์"
            />
            <div className="flex gap-3">
              <NavBtn variant="outline" onClick={() => setStep(2)}>
                ← ย้อนกลับ
              </NavBtn>
              <NavBtn onClick={handleProcess} disabled={!canNext()}>
                <Zap className="w-4 h-4" />
                ประมวลผลทันที
              </NavBtn>
            </div>
          </Card>
        )}

        {step === 4 && (
          <Card title="Step 4 - ตรวจสอบผลลัพธ์">
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
                ผิดพลาด: {statusMsg}
              </div>
            )}

            {status === "done" && (
              <>
                <div className="grid grid-cols-3 gap-4 mb-6">
                  <StatCard label="ยืนยันแล้ว" value={confirmed} color="green" />
                  <StatCard label="อนุมาน" value={inferred} color="amber" />
                  <StatCard label="ไม่พบ / กรอกเอง" value={notFound} color="red" />
                </div>

                <ResultsTable rows={results} onChange={handleResultsChange} />

                <div className="flex gap-3 pt-4 border-t border-slate-100">
                  <NavBtn onClick={handleDownload} disabled={modal.type === "loading"}>
                    <Download className="w-4 h-4" />
                    ดาวน์โหลด RECAP_filled.xlsx
                  </NavBtn>
                </div>
              </>
            )}
          </Card>
        )}

        {step === 5 && (
          <Card title="เสร็จสิ้น!">
            <div className="text-center py-10 space-y-4">
              <div className="flex justify-center gap-2 text-5xl">
                <span className="animate-bounce" style={{ animationDelay: "0ms" }}>
                  🎉
                </span>
                <span className="animate-bounce" style={{ animationDelay: "100ms" }}>
                  ✅
                </span>
                <span className="animate-bounce" style={{ animationDelay: "200ms" }}>
                  🎊
                </span>
              </div>
              <p className="text-xl font-semibold text-slate-700">
                ดาวน์โหลดไฟล์สำเร็จแล้ว
              </p>
              <p className="text-slate-500 text-sm">
                ไฟล์ <strong>RECAP_filled.xlsx</strong> อยู่ในโฟลเดอร์ Downloads
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
      </div>

      <DownloadModal state={modal} onClose={() => setModal({ type: "hidden" })} />
    </main>
  );
}

function DownloadModal({
  state,
  onClose,
}: {
  state: ModalState;
  onClose: () => void;
}) {
  if (state.type === "hidden") return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-sm mx-4 overflow-hidden">
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]" />
        <div className="p-8 text-center space-y-5">
          {state.type === "loading" && (
            <>
              <div className="flex justify-center">
                <div className="animate-spin rounded-full h-14 w-14 border-4 border-pink-100 border-t-[#E91E8C]" />
              </div>
              <div>
                <p className="text-lg font-bold text-slate-800">กำลังสร้างไฟล์...</p>
                <p className="text-sm text-slate-500 mt-1">
                  โปรดรอสักครู่ อย่าปิดหน้าต่างนี้
                </p>
              </div>
            </>
          )}

          {state.type === "success" && (
            <>
              <div className="flex justify-center">
                <div className="rounded-full bg-green-100 p-4">
                  <CheckCircle className="w-12 h-12 text-green-500" />
                </div>
              </div>
              <div>
                <p className="text-lg font-bold text-slate-800">ดาวน์โหลดสำเร็จ!</p>
                <p className="text-sm text-slate-500 mt-1">
                  ไฟล์ <strong>RECAP_filled.xlsx</strong> อยู่ในโฟลเดอร์ Downloads
                </p>
              </div>
              <button
                onClick={onClose}
                className="w-full py-3 rounded-xl font-semibold text-white bg-gradient-to-r from-[#E91E8C] to-[#d41679] hover:from-[#d41679] hover:to-[#be185d] transition-all shadow-sm hover:shadow-md"
              >
                ปิด
              </button>
            </>
          )}

          {state.type === "error" && (
            <>
              <div className="flex justify-center">
                <div className="rounded-full bg-red-100 p-4">
                  <XCircle className="w-12 h-12 text-red-500" />
                </div>
              </div>
              <div>
                <p className="text-lg font-bold text-slate-800">เกิดข้อผิดพลาด</p>
                <p className="text-sm text-red-600 mt-2 bg-red-50 rounded-lg px-3 py-2 text-left break-words">
                  {state.message}
                </p>
              </div>
              <button
                onClick={onClose}
                className="w-full py-3 rounded-xl font-semibold text-white bg-slate-600 hover:bg-slate-700 transition-all"
              >
                ปิด
              </button>
            </>
          )}
        </div>
      </div>
    </div>
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
