"use client";

import { useRef, useState } from "react";
import { Download, FileSpreadsheet, AlertTriangle, CheckCircle2, Loader2 } from "lucide-react";
import type { PipelineSnapshot } from "@/lib/types";
import DropZone from "@/components/DropZone";
import { buildMinorReportSheets } from "@/lib/minorReport";
import { buildMinorReportFillPlan } from "@/lib/minorReportDownload";

interface Props {
  snapshot: PipelineSnapshot | null;
}

type Status = "idle" | "building" | "done" | "error";

function triggerBrowserDownload(filename: string, buffer: ArrayBuffer) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Minor Report sub-tab — a read-only consumer of RECAP Auto-Filler's finished output,
 * PLUS its own local upload for the `TO BE_Minor Report.xlsx` template. The template
 * upload lives entirely here (never touches the wizard's DropZones/state) — this is
 * exactly the kind of "Minor Report needs its own extra upload" case the architecture
 * was designed to allow without touching app/page.tsx beyond the one prop it already
 * receives (`snapshot`).
 *
 * The template's formatting is preserved byte-for-byte via the same ZIP/XML patch
 * technique RECAP uses (see lib/minorReportDownload.worker.ts) — never a SheetJS
 * read→write round-trip, which would risk losing colours/borders/column widths.
 */
export default function MinorReportTab({ snapshot }: Props) {
  const [templateFiles, setTemplateFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<Status>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct] = useState(0);
  const [error, setError] = useState("");
  const [counts, setCounts] = useState<{ newItem: number; newNotLink: number; deleteItem: number } | null>(null);

  const workerRef = useRef<Worker | null>(null);
  const templateFile = templateFiles[0] ?? null;

  const canBuild = !!snapshot && !!templateFile && status !== "building";

  const handleBuild = async () => {
    if (!snapshot || !templateFile) return;
    setStatus("building");
    setError("");
    setPct(0);
    setStatusMsg("กำลังจัดรูปแบบข้อมูล...");

    try {
      const sheets = buildMinorReportSheets(snapshot);
      setCounts({
        newItem: sheets.newItem.length,
        newNotLink: sheets.newNotLink.length,
        deleteItem: sheets.deleteItem.length,
      });
      const plan = buildMinorReportFillPlan(sheets);

      setStatusMsg("กำลังอ่านไฟล์ template...");
      const templateBuf = await templateFile.arrayBuffer();

      const worker = new Worker(new URL("../lib/minorReportDownload.worker.ts", import.meta.url));
      workerRef.current = worker;

      worker.onmessage = (e: MessageEvent) => {
        const msg = e.data as { type: string; pct?: number; buffer?: ArrayBuffer; message?: string };
        switch (msg.type) {
          case "init_ok":
            worker.postMessage({ type: "build", plan });
            break;
          case "progress":
            setPct(msg.pct ?? 0);
            setStatusMsg("กำลังเติมข้อมูลลงไฟล์ template...");
            break;
          case "done": {
            const outBuf = msg.buffer!;
            worker.terminate();
            workerRef.current = null;
            triggerBrowserDownload("Minor Report.xlsx", outBuf);
            setStatus("done");
            setStatusMsg("เสร็จสิ้น!");
            break;
          }
          case "error":
            worker.terminate();
            workerRef.current = null;
            setError(msg.message ?? "เกิดข้อผิดพลาดใน worker");
            setStatus("error");
            break;
        }
      };
      worker.onerror = (e: ErrorEvent) => {
        worker.terminate();
        workerRef.current = null;
        setError(e.message ?? "Worker crashed");
        setStatus("error");
      };

      const buf = templateBuf.slice(0);
      worker.postMessage({ type: "init", buffer: buf }, [buf]);
    } catch (err) {
      setError(String(err));
      setStatus("error");
    }
  };

  if (!snapshot) {
    return (
      <div className="max-w-2xl mx-auto mt-16 text-center space-y-4">
        <div className="flex justify-center">
          <div className="rounded-full bg-amber-100 p-4">
            <AlertTriangle className="w-10 h-10 text-amber-500" />
          </div>
        </div>
        <p className="text-lg font-semibold text-slate-700">ยังไม่มีข้อมูลสำหรับสร้าง Minor Report</p>
        <p className="text-slate-500 text-sm">
          ไปที่แท็บ &quot;อัปโหลดข้อมูล&quot; แล้วอัปโหลดไฟล์ทั้ง 4 ไฟล์ (Check Space, FILE_INDEX, RECAP, 100 ช่อง)
          พร้อม DATA_SPACEMAN แล้วกด &quot;ประมวลผล&quot; ให้เสร็จ (Step 5) ก่อน — Minor Report จะดึงผลลัพธ์ชุดเดียวกันมาสร้างรายงานให้อัตโนมัติ
          ไม่ต้องอัปโหลดซ้ำ
        </p>
      </div>
    );
  }

  return (
    <div className="max-w-3xl mx-auto space-y-6">
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div className="flex items-center gap-3 mb-4">
          <div className="bg-gradient-to-br from-[#E91E8C] to-[#F15A22] rounded-xl p-2.5">
            <FileSpreadsheet className="w-5 h-5 text-white" />
          </div>
          <div>
            <h2 className="text-lg font-bold text-slate-800">Minor Report</h2>
            <p className="text-sm text-slate-500">
              ดึงผลลัพธ์จากแท็บ &quot;อัปโหลดข้อมูล&quot; ที่ประมวลผลไว้แล้วมาเติมลงไฟล์ template — ไม่ parse ไฟล์ RECAP ซ้ำ
            </p>
          </div>
        </div>

        <DropZone
          label="TO BE_Minor Report.xlsx (ไฟล์ template)"
          hint="ไฟล์ template เปล่า 3 ชีต (Recap_New_item / Recap_New_not_link / Recap_Delete_item) — format/สี/เส้นขอบในไฟล์นี้จะถูกเก็บไว้เป๊ะๆ มีแค่ข้อมูลที่จะถูกเติมเข้าไปแทนที่แถวข้อมูลเดิม"
          accept=".xlsx,.xls"
          files={templateFiles}
          onFiles={setTemplateFiles}
        />

        <button
          onClick={handleBuild}
          disabled={!canBuild}
          className="mt-4 w-full flex items-center justify-center gap-2 px-5 py-2.5 text-white rounded-xl font-semibold transition-all shadow-md hover:shadow-lg hover:scale-[1.02] disabled:opacity-60 disabled:hover:scale-100 bg-gradient-to-r from-[#E91E8C] to-[#F15A22]"
        >
          {status === "building" ? <Loader2 className="w-4 h-4 animate-spin" /> : <FileSpreadsheet className="w-4 h-4" />}
          {status === "building" ? `${statusMsg} (${pct}%)` : "สร้างและดาวน์โหลด Minor Report"}
        </button>

        {status === "building" && (
          <div className="mt-3 h-1.5 bg-slate-100 rounded-full overflow-hidden">
            <div
              className="h-full bg-gradient-to-r from-[#E91E8C] to-[#F15A22] transition-all duration-300"
              style={{ width: `${pct}%` }}
            />
          </div>
        )}

        {error && (
          <p className="mt-3 text-sm text-red-600 flex items-center gap-1.5">
            <AlertTriangle className="w-4 h-4 flex-shrink-0" /> {error}
          </p>
        )}
      </div>

      {status === "done" && counts && (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-4">
          <div className="flex items-center gap-2 text-emerald-600 font-semibold">
            <CheckCircle2 className="w-5 h-5" />
            สร้างสำเร็จ — ดาวน์โหลดแล้ว
          </div>
          <div className="grid grid-cols-3 gap-4 text-center">
            <SummaryStat label="Recap_New_item" count={counts.newItem} />
            <SummaryStat label="Recap_New_not_link" count={counts.newNotLink} />
            <SummaryStat label="Recap_Delete_item" count={counts.deleteItem} />
          </div>
        </div>
      )}
    </div>
  );
}

function SummaryStat({ label, count }: { label: string; count: number }) {
  return (
    <div className="bg-slate-50 rounded-xl p-3 border border-slate-100">
      <div className="text-2xl font-bold text-slate-800">{count}</div>
      <div className="text-xs text-slate-500 mt-1">{label}</div>
    </div>
  );
}
