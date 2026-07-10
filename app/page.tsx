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
  X,
  ChevronRight,
  Settings2,
  BookOpen,
  Database,
  AlertTriangle,
} from "lucide-react";

import StepIndicator from "@/components/StepIndicator";
import DropZone from "@/components/DropZone";
import ResultsTable from "@/components/ResultsTable";
import SpacemanMaster, {
  DriveFileInfo,
  formatDateTime,
  type SpacemanValues,
} from "@/components/SpacemanMaster";
import ConfigMenu, { EXCEPTION_CONFIG_KEY, type SyncStatus } from "@/components/ConfigMenu";
import { toDownloadRows, type DownloadRow, type CheckSpaceFillPlan, type FillRow } from "@/lib/download";
import FillEditTable, {
  type TabColDef,
  type EditableFillRow,
  convertToEditableRows,
  convertFromEditableRows,
} from "@/components/FillEditTable";
import {
  parseMissingRows,
  parseXlsbFiles,
  buildStructureLookup,
  parsePlanogramLookup,
  processRows,
  parseCheckSpace,
  parseFileIndex,
  buildXlsbExtraInfoMap,
  fillNewDeleteIM,
  fillNewSCM,
  fillDelSCM,
} from "@/lib/processor";
import type { ProcessedRow, ExceptionConfig, FilledData } from "@/lib/types";
import { makeEntry, sendLog } from "@/lib/logger";

// ─── Types ─────────────────────────────────────────────────────────────────

const STEPS = [
  { id: 1, label: "Check Space" },
  { id: 2, label: "FILE_INDEX" },
  { id: 3, label: "RECAP" },
  { id: 4, label: "100 ช่อง" },
  { id: 5, label: "ตรวจสอบ" },
  { id: 6, label: "ดาวน์โหลด" },
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

// ─── Helpers ───────────────────────────────────────────────────────────────

function triggerBrowserDownload(label: string, buffer: ArrayBuffer) {
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
}

// ─── Check Space fill table definitions ────────────────────────────────────

const NEW_STATUS_OPTIONS = [
  "NEW ADD SOME STORE",
  "NEW ADD ALL STORE",
  "NEW DELETE SOME STORE",
  "NEW DELETE ALL STORE",
];

const DEL_STATUS_OPTIONS = [
  "DELETE SOME STORE",
  "DELETE ALL STORE",
  "DELETE ALL STORE (IN/OUT)",
  "DELETE SOME STORE (IN/OUT)",
];

const NDIM_COLDEFS: TabColDef[] = [
  { field: "seqNew",    col: 0,  label: "ลำดับ(New)",  editable: false, zone: "new" },
  { field: "dcNew",     col: 3,  label: "DC",           editable: false, zone: "new" },
  { field: "byCodeNew", col: 4,  label: "BY_CODE(New)", editable: true,  zone: "new" },
  { field: "statusNew", col: 5,  label: "Status(New)",  editable: true,  zone: "new" },
  { field: "remark",    col: 6,  label: "Remark",       editable: true,  zone: "new" },
  { field: "seqDel",   col: 8,  label: "ลำดับ(Del)",  editable: false, zone: "del" },
  { field: "dcDel",    col: 11, label: "DC(Del)",      editable: false, zone: "del" },
  { field: "byCodeDel",col: 12, label: "BY_CODE(Del)", editable: true,  zone: "del" },
  { field: "statusDel",col: 13, label: "Status(Del)",  editable: true,  zone: "del" },
  { field: "extraInfo",col: 14, label: "Extra_Info",   editable: true,  zone: "del" },
];

const NSCM_COLDEFS: TabColDef[] = [
  { field: "seq",       col: 0,  label: "ลำดับ",            editable: false },
  { field: "barcode",   col: 3,  label: "Barcode",           editable: false },
  { field: "name",      col: 4,  label: "ชื่อสินค้า",       editable: false },
  { field: "division",  col: 5,  label: "F — DIVISION",      editable: true },
  { field: "dept",      col: 6,  label: "G — DEPT",          editable: true, cascade: "division" },
  { field: "subDept",   col: 7,  label: "H — SUB-DEPT",      editable: true, cascade: "dept" },
  { field: "cls",       col: 8,  label: "I — Class",         editable: true, cascade: "subDept" },
  { field: "planogram", col: 9,  label: "J — PLANOGRAM",     editable: true },
  { field: "status",    col: 10, label: "Status",            editable: true },
  { field: "remark",    col: 11, label: "Remark",            editable: true },
  { field: "implement", col: 12, label: "POG ROUND", editable: true },
  { field: "colN",      col: 13, label: "MBC FCST",  editable: true },
  { field: "colPiece",  col: 14, label: "Piece",     editable: true },
  { field: "colO",      col: 15, label: "%",          editable: true },
  { field: "colNet",    col: 16, label: "Net",        editable: false },
];

const DSCM_COLDEFS: TabColDef[] = [
  { field: "seq",       col: 0, label: "ลำดับ",        editable: false },
  { field: "barcode",   col: 3, label: "Barcode",       editable: false },
  { field: "name",      col: 4, label: "ชื่อสินค้า",   editable: false },
  { field: "division",  col: 5, label: "Division",      editable: true },
  { field: "category",  col: 6, label: "Category",      editable: true, cascade: "division" },
  { field: "implement", col: 7, label: "POG ROUND",     editable: true },
  { field: "status",    col: 8, label: "Status",        editable: true },
  { field: "extraInfo", col: 9, label: "Extra_Info",    editable: true },
];

// ─── Home ──────────────────────────────────────────────────────────────────

export default function Home() {
  const [view, setView] = useState<AppView>("main");
  const [step, setStep] = useState(1);
  const [status, setStatus] = useState<Status>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct] = useState(0);

  const [checkSpaceFile, setCheckSpaceFile] = useState<File | null>(null);
  const [fileIndexFile, setFileIndexFile] = useState<File | null>(null);
  const [recapFiles, setRecapFiles] = useState<File[]>([]);
  const [xlsbFiles, setXlsbFiles] = useState<File[]>([]);
  const [driveFileInfo, setDriveFileInfo] = useState<DriveFileInfo | null>(null);
  const [driveLoading, setDriveLoading] = useState(true);
  const [results, setResults] = useState<ProcessedRow[]>([]);

  // Check Space fill tables (3 editable tabs in Step 5)
  interface FillTabData {
    sheetName: string;
    displayName: string;
    colDefs: TabColDef[];
    rows: EditableFillRow[];
    originalFillRows: FillRow[];
  }
  const [fillTabs, setFillTabs]     = useState<FillTabData[] | null>(null);
  const [previewTab, setPreviewTab] = useState(0);

  // Exception config — loaded from Google Sheets on mount; localStorage is fallback cache
  const [exceptionConfig, setExceptionConfig] = useState<ExceptionConfig[]>(() => {
    try {
      const raw = localStorage.getItem(EXCEPTION_CONFIG_KEY);
      return raw ? (JSON.parse(raw) as ExceptionConfig[]) : [];
    } catch {
      return [];
    }
  });
  const [showConfig, setShowConfig] = useState(false);
  const [showHelp, setShowHelp] = useState(false);
  const [configSyncStatus, setConfigSyncStatus] = useState<SyncStatus>("loading");
  const [configLastSaved, setConfigLastSaved] = useState<string | null>(null);
  const [configSyncError, setConfigSyncError] = useState("");
  // Unique values from DATA_SPACEMAN for config dropdowns
  const [spacemanValues, setSpacemanValues] = useState<SpacemanValues>({
    categories: [],
    subcategories: [],
    descAList: [],
    descBList: [],
    descCList: [],
    hierarchyMap: { divToDept: {}, deptToSub: {}, subToCls: {} },
    catToSub: {},
  });
  const [spacemanLoaded, setSpacemanLoaded] = useState(false);

  // Queue state (display only — heavy data lives in refs)
  const [jobs, setJobs] = useState<BuildJob[]>([]);
  const [queuePanelOpen, setQueuePanelOpen] = useState(false);

  // Refs — not in React state to avoid re-render overhead and serialization issues
  const recapBufRef = useRef<ArrayBuffer | null>(null);
  const checkSpacePlanRef = useRef<CheckSpaceFillPlan | null>(null);
  const sessionIdRef = useRef<string>("");
  const workersRef = useRef<Map<string, Worker>>(new Map());
  const jobDataRef = useRef<Map<string, { recapBuf: ArrayBuffer; rows: DownloadRow[]; checkSpacePlan?: CheckSpaceFillPlan }>>(new Map());
  const jobCounterRef = useRef(0);
  const autoDownloadedRef = useRef<Set<string>>(new Set());

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

  // Load exception config from Google Sheets on mount (localStorage is fallback)
  useEffect(() => {
    setConfigSyncStatus("loading");
    fetch("/api/config/load")
      .then((r) => r.json())
      .then((data) => {
        if (data.error) throw new Error(data.error);
        const loaded: ExceptionConfig[] = data.config ?? [];
        setExceptionConfig(loaded);
        setConfigLastSaved(data.lastSaved ?? null);
        localStorage.setItem(EXCEPTION_CONFIG_KEY, JSON.stringify(loaded));
        setConfigSyncStatus("idle");
      })
      .catch((e) => {
        setConfigSyncError(String(e));
        setConfigSyncStatus("error");
      });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // ── Core job starter (stable — only touches refs + functional setJobs) ──

  const startJobFn = useCallback((id: string, recapBuf: ArrayBuffer, rows: DownloadRow[], checkSpacePlan?: CheckSpaceFillPlan) => {
    const buildSid = `build-${id.slice(0, 8)}`;
    const buildStart = Date.now();

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
          worker.postMessage({ type: "build", rows, checkSpacePlan });
          break;
        case "progress":
          setJobs((prev) =>
            prev.map((j) => (j.id === id ? { ...j, progress: msg.pct ?? j.progress } : j))
          );
          break;
        case "done": {
          workersRef.current.delete(id);
          jobDataRef.current.delete(id);
          const durSec = ((Date.now() - buildStart) / 1000).toFixed(1);
          setJobs((prev) => {
            const label = prev.find((j) => j.id === id)?.label ?? id;
            sendLog([makeEntry(buildSid, "BUILD_COMPLETE", "INFO",
              `Build ${label} เสร็จในเวลา ${durSec} วินาที`,
              { jobId: id, label, durationSec: durSec }
            )]);
            return prev.map((j) =>
              j.id === id
                ? { ...j, status: "done", progress: 100, completedAt: new Date(), buffer: msg.buffer }
                : j
            );
          });
          break;
        }
        case "error": {
          workersRef.current.delete(id);
          setJobs((prev) => {
            const label = prev.find((j) => j.id === id)?.label ?? id;
            sendLog([makeEntry(buildSid, "BUILD_FAILED", "ERROR",
              `Build ${label} ล้มเหลว: ${msg.message ?? "Worker error"}`,
              { jobId: id, label, error: msg.message }
            )]);
            return prev.map((j) =>
              j.id === id
                ? { ...j, status: "failed", error: msg.message ?? "Worker error", completedAt: new Date() }
                : j
            );
          });
          break;
        }
      }
    };

    worker.onerror = (e: ErrorEvent) => {
      workersRef.current.delete(id);
      setJobs((prev) => {
        const label = prev.find((j) => j.id === id)?.label ?? id;
        sendLog([makeEntry(buildSid, "BUILD_FAILED", "ERROR",
          `Build ${label} crash: ${e.message ?? "Worker crashed"}`,
          { jobId: id, label, error: e.message }
        )]);
        return prev.map((j) =>
          j.id === id
            ? { ...j, status: "failed", error: e.message ?? "Worker crashed", completedAt: new Date() }
            : j
        );
      });
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
    startJobFn(next.id, data.recapBuf, data.rows, data.checkSpacePlan);
  }, [jobs, startJobFn]);

  // ── Queue actions ───────────────────────────────────────────────────────

  const enqueueJob = () => {
    if (!recapBufRef.current || results.length === 0) return;
    const id = crypto.randomUUID();
    const num = ++jobCounterRef.current;
    const baseName = recapFiles[0]?.name.replace(/\.[^.]+$/, "") ?? "RECAP";
    const label = `${baseName}_filled_#${num}.xlsx`;

    jobDataRef.current.set(id, {
      recapBuf: recapBufRef.current.slice(0),
      rows: toDownloadRows(results),
      checkSpacePlan: checkSpacePlanRef.current ?? undefined,
    });

    sendLog([makeEntry(
      `build-${id.slice(0, 8)}`, "BUILD_QUEUED", "INFO",
      `เพิ่มไฟล์ ${label} เข้าคิว Build (session: ${sessionIdRef.current.slice(0, 8)})`,
      { jobId: id, filename: label, processingSession: sessionIdRef.current }
    )]);

    setJobs((prev) => [
      ...prev,
      { id, label, status: "queued", createdAt: new Date(), progress: 0 },
    ]);
    setQueuePanelOpen(true);
    setStep(6);
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
    triggerBrowserDownload(label, buffer);
    // Buffer is kept so the user can re-download; status stays as-is
  };

  // Auto-download each job the moment it completes
  useEffect(() => {
    for (const job of jobs) {
      if (job.status === "done" && job.buffer && !autoDownloadedRef.current.has(job.id)) {
        autoDownloadedRef.current.add(job.id);
        triggerBrowserDownload(job.label, job.buffer);
        setJobs((prev) =>
          prev.map((j) => (j.id === job.id ? { ...j, status: "downloaded" } : j))
        );
      }
    }
  }, [jobs]);

  // ── Main processing flow ────────────────────────────────────────────────

  const canProcess = () =>
    checkSpaceFile !== null &&
    fileIndexFile !== null &&
    recapFiles.length === 1 &&
    xlsbFiles.length > 0 &&
    driveFileInfo !== null;

  const handleProcess = async () => {
    if (!checkSpaceFile || !fileIndexFile || !recapFiles[0] || xlsbFiles.length === 0 || !driveFileInfo) return;

    const sessionId = crypto.randomUUID();
    sessionIdRef.current = sessionId;
    const t0 = Date.now();

    setStatus("processing");
    setStep(5);
    setPct(0);

    sendLog([makeEntry(sessionId, "PROCESS_START", "INFO",
      `เริ่มประมวลผล: ${recapFiles[0].name} + ${xlsbFiles.length} ไฟล์ 100 ช่อง + ${driveFileInfo.name}`,
      {
        recapFile: recapFiles[0].name,
        xlsbFiles: xlsbFiles.map((f) => f.name),
        spacemanFile: driveFileInfo.name,
        spacemanFileId: driveFileInfo.id,
      }
    )]);

    try {
      setStatusMsg("อ่านไฟล์ RECAP...");
      setPct(5);
      const recapBuf = await recapFiles[0].arrayBuffer();
      recapBufRef.current = recapBuf.slice(0);

      // Pass 1: read sheet names only (fast — no cell parsing)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wbMeta = XLSX.read(recapBuf, { type: "array", bookSheets: true } as any);
      const findActualName = (target: string) => {
        const lo = target.toLowerCase().trim();
        return wbMeta.SheetNames.find((n: string) => n.toLowerCase().trim() === lo) ?? target;
      };
      const newScmActual = findActualName("NEW SCM");
      const ndimActual   = findActualName("NEW_DELETE_IM");
      const dscmActual   = findActualName("DEL SCM");

      // Pass 2: parse only the 3 sheets we need (avoids loading 165k+ cells in unused sheets)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = XLSX.read(recapBuf, { type: "array", sheets: [newScmActual, ndimActual, dscmActual] } as any);

      // ── Check Space pre-fill (runs BEFORE parseMissingRows so new rows are included) ──
      // Fill functions write to wb in-memory (needed for parseMissingRows) AND
      // return FillRow[] that the download worker uses for ZIP-patch — no XLSX.write needed.
      setStatusMsg("อ่านไฟล์ Check Space และ FILE_INDEX...");
      setPct(8);
      // Captured outside the CS try block so they're accessible after processRows
      let csNdimRows: FillRow[] = [];
      let csNewScmRows: FillRow[] = [];
      let csDscmRows: FillRow[] = [];

      checkSpacePlanRef.current = null;
      try {
        const [csItems, indexLookup, xlsbExtraInfo] = await Promise.all([
          parseCheckSpace(checkSpaceFile),
          parseFileIndex(fileIndexFile),
          buildXlsbExtraInfoMap(xlsbFiles),
        ]);
        if (csItems.length > 0) {
          setStatusMsg(`พบ ${csItems.length} รายการจาก Check Space — เตรียมข้อมูล...`);
          const newDeleteIMRows = fillNewDeleteIM(wb, csItems, indexLookup, xlsbExtraInfo);
          const newScmRows      = fillNewSCM(wb, csItems, indexLookup);
          const delScmRows      = fillDelSCM(wb, csItems, indexLookup, xlsbExtraInfo);
          csNdimRows    = newDeleteIMRows;
          csNewScmRows  = newScmRows;
          csDscmRows    = delScmRows;
          // Store plan — use ACTUAL names resolved from the RECAP file (avoids name-mismatch in worker)
          checkSpacePlanRef.current = {
            newScmRows,
            extraSheets: [
              { sheetName: ndimActual, rows: newDeleteIMRows },
              { sheetName: dscmActual, rows: delScmRows },
            ],
          };
          // recapBufRef.current stays as the ORIGINAL buffer — no XLSX.write
          const newCount = csItems.filter(i => !i.status.toUpperCase().startsWith("DELETE")).length;
          const delCount = csItems.filter(i => i.status.toUpperCase().startsWith("DELETE")).length;
          sendLog([makeEntry(sessionId, "PROCESS_START", "INFO",
            `Check Space: ${newCount} NEW / ${delCount} DELETE items`,
            { checkSpaceFile: checkSpaceFile.name, fileIndexFile: fileIndexFile.name, totalItems: csItems.length }
          )]);
          sendLog([makeEntry(sessionId, "CS_FILL_DIAG", "INFO",
            `SheetNames: [${wbMeta.SheetNames.join(", ")}] → loaded: [${wb.SheetNames.join(", ")}] | NEW_DELETE_IM(${ndimActual}): ${newDeleteIMRows.length}r | DEL SCM(${dscmActual}): ${delScmRows.length}r | NEW SCM(${newScmActual}): ${newScmRows.length}r`,
            {
              allSheetNames: wbMeta.SheetNames,
              loadedSheets:  wb.SheetNames,
              ndimActual, dscmActual, newScmActual,
              ndimRows: newDeleteIMRows.length,
              nscmRows: newScmRows.length,
              dscmRows: delScmRows.length,
            }
          )]);

          // setFillTabs is called AFTER processRows below so NEW SCM tab has planogram data
        }
      } catch (csErr) {
        sendLog([makeEntry(sessionId, "ERROR", "WARN",
          `Check Space/FILE_INDEX parse failed: ${String(csErr)}`,
          { error: String(csErr) }
        )]);
      }

      setStatusMsg("อ่านไฟล์ RECAP...");
      setPct(10);
      const { rows: missing, totalScanned, alreadyFilled } = parseMissingRows(wb);

      const recapLevel = missing.length === 0 && totalScanned > 0 ? "WARN" : "INFO";
      const recapMsg = missing.length === 0 && totalScanned === 0
        ? `ไม่พบบาร์โค้ดในชีท NEW SCM — ตรวจสอบว่าไฟล์ถูกต้อง`
        : missing.length === 0
          ? `ไม่พบแถวที่ต้องเติม — บาร์โค้ดทั้งหมด ${totalScanned} รายการมีข้อมูลคอลัมน์ F อยู่แล้ว`
          : `RECAP: พบ ${missing.length} แถวที่ต้องเติม (สแกน ${totalScanned} แถว, ข้าม ${alreadyFilled} แถวที่มีข้อมูลอยู่แล้ว)`;
      sendLog([makeEntry(sessionId, "RECAP_PARSED", recapLevel, recapMsg, {
        totalScanned, rowsMissing: missing.length, rowsAlreadyFilled: alreadyFilled,
      })]);

      setStatusMsg(`พบ ${missing.length} รายการที่ต้องเติมข้อมูล — กำลังค้นหาในไฟล์ 100 ช่อง...`);
      setPct(25);

      const [barcodeMap, structureMap] = await Promise.all([
        parseXlsbFiles(xlsbFiles),
        buildStructureLookup(xlsbFiles),
      ]);

      sendLog([makeEntry(sessionId, "XLSB_PARSED", "INFO",
        `100 ช่อง: พบบาร์โค้ด ${barcodeMap.size} รายการ, โครงสร้างสินค้า ${structureMap.size} รายการ จาก ${xlsbFiles.length} ไฟล์`,
        { files: xlsbFiles.map((f) => f.name), barcodesFound: barcodeMap.size, structureFound: structureMap.size }
      )]);

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
      const planogramResult = await parsePlanogramLookup(spacemanFile, (p) =>
        setPct(50 + p * 0.35)
      );

      sendLog([makeEntry(sessionId, "SPACEMAN_PARSED", "INFO",
        `DATA_SPACEMAN: พบ prefix ${planogramResult.byPrefix.size} รายการ, UPC ${planogramResult.byUpc.size} รายการ`,
        { prefixCount: planogramResult.byPrefix.size, upcCount: planogramResult.byUpc.size, filename: driveFileInfo.name }
      )]);

      setStatusMsg("ประมวลผลข้อมูล...");
      setPct(90);
      const processed = processRows(missing, barcodeMap, structureMap, planogramResult, exceptionConfig);

      const cCount = processed.filter((r) => r.confidence === "confirmed").length;
      const iCount = processed.filter((r) => r.confidence === "inferred").length;
      const sCount = processed.filter((r) => r.confidence === "from_spaceman").length;
      const nCount = processed.filter((r) => r.confidence === "not_found").length;
      const durSec = ((Date.now() - t0) / 1000).toFixed(1);

      sendLog([makeEntry(sessionId, "PROCESS_COMPLETE", "INFO",
        `เสร็จใน ${durSec}s — ยืนยัน ${cCount} | ไม่มี Planogram ${iCount} | จาก Spaceman ${sCount} | ไม่พบ ${nCount}`,
        { confirmed: cCount, inferred: iCount, fromSpaceman: sCount, notFound: nCount, durationSec: durSec }
      )]);

      setResults(processed);

      // Build fillTabs now that processRows has run — NEW SCM needs planogram data from `processed`
      if (csNdimRows.length > 0 || csNewScmRows.length > 0 || csDscmRows.length > 0) {
        const nscmEditableRows: EditableFillRow[] = csNewScmRows.map(fr => {
          const pr = processed.find(r => r.rowIndex === fr.rowIndex);
          const pdata = pr ? { ...pr.filled, ...pr.override } as Record<string, string> : {};
          return {
            rowIndex: fr.rowIndex,
            fields: {
              seq:       fr.cells.find(c => c.col === 0)?.value  ?? "",
              barcode:   fr.cells.find(c => c.col === 3)?.value  ?? "",
              name:      fr.cells.find(c => c.col === 4)?.value  ?? "",
              division:  pdata.division  ?? "",
              dept:      pdata.dept      ?? "",
              subDept:   pdata.subDept   ?? "",
              cls:       pdata.cls       ?? "",
              planogram: pdata.planogram ?? "",
              status:    fr.cells.find(c => c.col === 10)?.value ?? "",
              remark:    fr.cells.find(c => c.col === 11)?.value ?? "",
              implement: fr.cells.find(c => c.col === 12)?.value ?? "",
              colN:      pdata.colN      ?? "",
              colPiece:  pdata.colPiece  ?? "",
              colO:      pdata.colO      ?? "",
              colNet:    (() => {
                const pct   = parseFloat(pdata.colO    ?? "") || 0;
                const piece = parseFloat(pdata.colPiece ?? "") || 0;
                return (pct > 0 && piece > 0)
                  ? (Math.round((pct / 100) * piece * 100) / 100).toFixed(2)
                  : "";
              })(),
            },
          };
        });
        setFillTabs([
          {
            sheetName: ndimActual,
            displayName: "NEW_DELETE_IM",
            colDefs: NDIM_COLDEFS,
            rows: convertToEditableRows(csNdimRows, NDIM_COLDEFS),
            originalFillRows: csNdimRows,
          },
          {
            sheetName: newScmActual,
            displayName: "NEW SCM",
            colDefs: NSCM_COLDEFS,
            rows: nscmEditableRows,
            originalFillRows: csNewScmRows,
          },
          {
            sheetName: dscmActual,
            displayName: "DEL SCM",
            colDefs: DSCM_COLDEFS,
            rows: convertToEditableRows(csDscmRows, DSCM_COLDEFS),
            originalFillRows: csDscmRows,
          },
        ]);
        setPreviewTab(0);
      }

      setPct(100);
      setStatusMsg("เสร็จสิ้น!");
      setStatus("done");
    } catch (err) {
      sendLog([makeEntry(sessionId, "ERROR", "ERROR",
        `เกิดข้อผิดพลาดระหว่างประมวลผล: ${String(err)}`,
        { error: String(err) }
      )]);
      setStatus("error");
      setStatusMsg(String(err));
    }
  };

  const handleResultsChange = (updated: ProcessedRow[]) => {
    setResults(updated);
  };

  const handleFillTabChange = (tabIdx: number, updatedRows: EditableFillRow[]) => {
    setFillTabs(prev => {
      if (!prev) return prev;
      const next = [...prev];
      const tab = next[tabIdx];
      // For NSCM: recompute colNet whenever colPiece or colO change
      const rowsToStore = tabIdx === 1
        ? updatedRows.map(r => {
            const pct   = parseFloat(r.fields.colO    ?? "") || 0;
            const piece = parseFloat(r.fields.colPiece ?? "") || 0;
            const colNet = (pct > 0 && piece > 0)
              ? (Math.round((pct / 100) * piece * 100) / 100).toFixed(2)
              : "";
            return colNet === r.fields.colNet ? r : { ...r, fields: { ...r.fields, colNet } };
          })
        : updatedRows;
      next[tabIdx] = { ...tab, rows: rowsToStore };
      // Sync edits back to checkSpacePlanRef so the worker uses the updated data
      if (checkSpacePlanRef.current) {
        const updatedFillRows = convertFromEditableRows(updatedRows, tab.originalFillRows, tab.colDefs);
        if (tabIdx === 0) {
          checkSpacePlanRef.current = {
            ...checkSpacePlanRef.current,
            extraSheets: checkSpacePlanRef.current.extraSheets.map(
              (sf, i) => i === 0 ? { ...sf, rows: updatedFillRows } : sf
            ),
          };
        } else if (tabIdx === 1) {
          checkSpacePlanRef.current = { ...checkSpacePlanRef.current, newScmRows: updatedFillRows };
        } else if (tabIdx === 2) {
          checkSpacePlanRef.current = {
            ...checkSpacePlanRef.current,
            extraSheets: checkSpacePlanRef.current.extraSheets.map(
              (sf, i) => i === 1 ? { ...sf, rows: updatedFillRows } : sf
            ),
          };
        }
      }
      return next;
    });
    // For NEW SCM: sync F-J/N-Q edits to results so applyRows (worker) uses updated planogram values
    if (tabIdx === 1) {
      setResults(prevResults =>
        prevResults.map(pr => {
          const er = updatedRows.find(r => r.rowIndex === pr.rowIndex);
          if (!er) return pr;
          const override: Partial<FilledData> = {
            ...(pr.override ?? {}),
            division:  er.fields.division  ?? "",
            dept:      er.fields.dept      ?? "",
            subDept:   er.fields.subDept   ?? "",
            cls:       er.fields.cls       ?? "",
            planogram: er.fields.planogram ?? "",
            colN:      er.fields.colN      ?? "",
            colPiece:  er.fields.colPiece  ?? "",
            colO:      er.fields.colO      ?? "",
          };
          return { ...pr, override };
        })
      );
    }
  };

  const reset = () => {
    // Jobs persist across resets — do NOT clear them
    recapBufRef.current = null;
    checkSpacePlanRef.current = null;
    setStep(1);
    setStatus("idle");
    setStatusMsg("");
    setPct(0);
    setCheckSpaceFile(null);
    setFileIndexFile(null);
    setRecapFiles([]);
    setXlsbFiles([]);
    setResults([]);
    setFillTabs(null);
    setPreviewTab(0);
  };

  const handleConfigChange = (updated: ExceptionConfig[]) => {
    setExceptionConfig(updated);
    localStorage.setItem(EXCEPTION_CONFIG_KEY, JSON.stringify(updated));
    setConfigSyncStatus("saving");
    fetch("/api/config/save", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ entries: updated }),
    })
      .then((r) => r.json())
      .then((data) => {
        if (data.error) throw new Error(data.error);
        // Server stamped updatedAt — sync those back so localStorage stays accurate
        if (data.entries) {
          setExceptionConfig(data.entries);
          localStorage.setItem(EXCEPTION_CONFIG_KEY, JSON.stringify(data.entries));
        }
        setConfigLastSaved(data.savedAt ?? null);
        setConfigSyncStatus("saved");
      })
      .catch((e) => {
        setConfigSyncError(String(e));
        setConfigSyncStatus("error");
      });
  };

  // ─── Fill-tab summary stats ─────────────────────────────────────────────────
  const ndimRows  = fillTabs?.[0]?.rows ?? [];
  const nscmRows  = fillTabs?.[1]?.rows ?? [];
  const dscmRows  = fillTabs?.[2]?.rows ?? [];

  const ndimNewCount = new Set(ndimRows.map(r => r.fields.seqNew).filter(Boolean)).size;
  const ndimDelCount = new Set(ndimRows.map(r => r.fields.seqDel).filter(Boolean)).size;

  const nscmFilled    = nscmRows.filter(r => {
    const pr = results.find(p => p.rowIndex === r.rowIndex);
    return pr?.confidence === "confirmed" || pr?.confidence === "from_spaceman";
  }).length;
  const nscmNotFilled = nscmRows.length - nscmFilled;

  const pendingNdim  = ndimRows.filter(r =>
    (r.fields.seqNew && !r.fields.statusNew) || (r.fields.seqDel && !r.fields.statusDel)
  ).length;
  const pendingNscm  = nscmRows.filter(r => !r.fields.division).length;
  const pendingDscm  = dscmRows.filter(r => !r.fields.status).length;
  const pendingTotal = pendingNdim + pendingNscm + pendingDscm;

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
        <div className="flex items-center justify-between">
          <div className="flex gap-0">
            <TabBtn active={view === "main"} onClick={() => setView("main")}>
              <FileSpreadsheet className="w-4 h-4" />
              อัปโหลดข้อมูล
            </TabBtn>
            <TabBtn active={view === "spaceman"} onClick={() => setView("spaceman")}>
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <ellipse cx="12" cy="5" rx="9" ry="3" />
                <path d="M3 5v14c0 1.66 4.03 3 9 3s9-1.34 9-3V5" />
                <path d="M3 12c0 1.66 4.03 3 9 3s9-1.34 9-3" />
              </svg>
              DATA_SPACEMAN
            </TabBtn>
          </div>
          <div className="flex items-center gap-1">
            {/* Help button */}
            <button
              onClick={() => setShowHelp(true)}
              title="วิธีการทำงาน"
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition-colors text-slate-500 hover:bg-slate-100"
            >
              <BookOpen className="w-4 h-4" />
              วิธีการทำงาน
            </button>
            {/* Config button — badge counts only non-deleted rules */}
            {(() => {
              const activeCount = exceptionConfig.filter((e) => e.status !== "deleted").length;
              return (
                <button
                  onClick={() => setShowConfig(true)}
                  title="Config Rules"
                  className={`flex items-center gap-1.5 px-3 py-1.5 mr-1 rounded-lg text-xs font-medium transition-colors ${
                    activeCount > 0
                      ? "bg-pink-50 text-[#E91E8C] border border-pink-200 hover:bg-pink-100"
                      : "text-slate-500 hover:bg-slate-100"
                  }`}
                >
                  <Settings2 className="w-4 h-4" />
                  Config Rules
                  {activeCount > 0 && (
                    <span className="bg-[#E91E8C] text-white text-[10px] font-bold rounded-full px-1.5 py-0.5 leading-none">
                      {activeCount}
                    </span>
                  )}
                </button>
              );
            })()}
          </div>
        </div>
      </div>

      {/* Main layout — flex with sticky queue panel on the right */}
      <div className="px-6 py-8">
        <div className="flex gap-6 items-start">

          {/* ── Content area ─────────────────────────────────────────── */}
          <div className="flex-1 min-w-0 space-y-8">

            {/* DATA_SPACEMAN — always mounted so fetch+parse starts immediately in background.
                Parsing runs in a Web Worker (separate thread) so the upload tab is unaffected.
                Hidden via CSS only; JS/Worker execution continues uninterrupted. */}
            <div className={view === "spaceman" ? "" : "hidden"}>
              <SpacemanMaster
                isVisible={view === "spaceman"}
                onFileInfoChange={(info) => {
                  setDriveFileInfo(info);
                  setDriveLoading(false);
                }}
                onSpacemanValues={(v) => { setSpacemanValues(v); setSpacemanLoaded(true); }}
              />
            </div>

            {/* Main upload flow */}
            {view === "main" && (
              <>
                <StepIndicator steps={STEPS} current={step} />

                {/* Step 1 — Upload Check Space */}
                {step === 1 && (
                  <Card title="Step 1 - อัปโหลดไฟล์ Check Space">
                    <DropZone
                      label="Check Space.xlsx"
                      accept=".xlsx,.xls"
                      files={checkSpaceFile ? [checkSpaceFile] : []}
                      onFiles={(files) => setCheckSpaceFile(files[0] ?? null)}
                    />
                    <div className="flex gap-3">
                      <NavBtn onClick={() => setStep(2)} disabled={!checkSpaceFile}>
                        ถัดไป →
                      </NavBtn>
                    </div>
                  </Card>
                )}

                {/* Step 2 — Upload FILE_INDEX */}
                {step === 2 && (
                  <Card title="Step 2 - อัปโหลดไฟล์ FILE_INDEX_1">
                    <DropZone
                      label="FILE_INDEX_1.xlsx"
                      accept=".xlsx,.xls"
                      files={fileIndexFile ? [fileIndexFile] : []}
                      onFiles={(files) => setFileIndexFile(files[0] ?? null)}
                    />
                    <div className="flex gap-3">
                      <NavBtn variant="outline" onClick={() => setStep(1)}>← ย้อนกลับ</NavBtn>
                      <NavBtn onClick={() => setStep(3)} disabled={!fileIndexFile}>
                        ถัดไป →
                      </NavBtn>
                    </div>
                  </Card>
                )}

                {/* Step 3 — Upload RECAP */}
                {step === 3 && (
                  <Card title="Step 3 - อัปโหลดไฟล์ RECAP">
                    <DropZone
                      label="ไฟล์ RECAP.xlsx"
                      accept=".xlsx,.xls"
                      files={recapFiles}
                      onFiles={setRecapFiles}
                    />
                    <div className="flex gap-3">
                      <NavBtn variant="outline" onClick={() => setStep(2)}>← ย้อนกลับ</NavBtn>
                      <NavBtn onClick={() => setStep(4)} disabled={recapFiles.length !== 1}>
                        ถัดไป →
                      </NavBtn>
                    </div>
                  </Card>
                )}

                {/* Step 4 — Upload 100 ช่อง */}
                {step === 4 && (
                  <Card title="Step 4 - อัปโหลดไฟล์ 100 ช่อง (.xlsb)">
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
                      <NavBtn variant="outline" onClick={() => setStep(3)}>← ย้อนกลับ</NavBtn>
                      <NavBtn onClick={handleProcess} disabled={!canProcess()}>
                        <Zap className="w-4 h-4" />
                        ประมวลผลทันที
                      </NavBtn>
                    </div>
                  </Card>
                )}

                {/* Step 5 — Review & enqueue */}
                {step === 5 && (
                  <Card title="Step 5 - ตรวจสอบผลลัพธ์">
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
                        {fillTabs && (
                          <>
                            {/* ── KPI Summary Cards (also serve as tab nav) ─── */}
                            <div className="grid grid-cols-4 gap-3 mb-4">
                              {/* Card 0 — NEW_DELETE_IM */}
                              {(() => {
                                const active = previewTab === 0;
                                return (
                                  <button
                                    onClick={() => setPreviewTab(0)}
                                    className={`rounded-xl border-2 p-4 flex flex-col items-center gap-1.5 text-center transition-all ${
                                      active
                                        ? "border-[#E91E8C] bg-pink-50 shadow-sm"
                                        : "border-slate-200 bg-white hover:border-slate-300 hover:bg-slate-50"
                                    }`}
                                  >
                                    <span className="text-[11px] font-semibold text-slate-500 truncate w-full text-center">NEW_DELETE_IM</span>
                                    <span className={`text-3xl font-bold ${active ? "text-[#E91E8C]" : "text-slate-700"}`}>
                                      {ndimRows.length}
                                    </span>
                                    <span className="text-[10px] text-slate-400">แถวทั้งหมด</span>
                                    <div className="flex gap-1 flex-wrap justify-center mt-0.5">
                                      <span className="px-1.5 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-semibold">
                                        NEW {ndimNewCount}
                                      </span>
                                      <span className="px-1.5 py-0.5 rounded-full bg-rose-100 text-rose-700 text-[10px] font-semibold">
                                        DEL {ndimDelCount}
                                      </span>
                                    </div>
                                  </button>
                                );
                              })()}

                              {/* Card 1 — NEW SCM */}
                              {(() => {
                                const active = previewTab === 1;
                                return (
                                  <button
                                    onClick={() => setPreviewTab(1)}
                                    className={`rounded-xl border-2 p-4 flex flex-col items-center gap-1.5 text-center transition-all ${
                                      active
                                        ? "border-blue-400 bg-blue-50 shadow-sm"
                                        : "border-slate-200 bg-white hover:border-slate-300 hover:bg-slate-50"
                                    }`}
                                  >
                                    <span className="text-[11px] font-semibold text-slate-500">NEW SCM</span>
                                    <span className={`text-3xl font-bold ${active ? "text-blue-600" : "text-slate-700"}`}>
                                      {nscmRows.length}
                                    </span>
                                    <span className="text-[10px] text-slate-400">แถวทั้งหมด</span>
                                    <div className="flex gap-1 flex-wrap justify-center mt-0.5">
                                      <span className="px-1.5 py-0.5 rounded-full bg-green-100 text-green-700 text-[10px] font-semibold">
                                        เจอ {nscmFilled}
                                      </span>
                                      {nscmNotFilled > 0 && (
                                        <span className="px-1.5 py-0.5 rounded-full bg-red-100 text-red-600 text-[10px] font-semibold">
                                          กรอกเอง {nscmNotFilled}
                                        </span>
                                      )}
                                    </div>
                                  </button>
                                );
                              })()}

                              {/* Card 2 — DEL SCM */}
                              {(() => {
                                const active = previewTab === 2;
                                return (
                                  <button
                                    onClick={() => setPreviewTab(2)}
                                    className={`rounded-xl border-2 p-4 flex flex-col items-center gap-1.5 text-center transition-all ${
                                      active
                                        ? "border-orange-400 bg-orange-50 shadow-sm"
                                        : "border-slate-200 bg-white hover:border-slate-300 hover:bg-slate-50"
                                    }`}
                                  >
                                    <span className="text-[11px] font-semibold text-slate-500">DEL SCM</span>
                                    <span className={`text-3xl font-bold ${active ? "text-orange-500" : "text-slate-700"}`}>
                                      {dscmRows.length}
                                    </span>
                                    <span className="text-[10px] text-slate-400">แถวทั้งหมด</span>
                                    <div className="flex gap-1 flex-wrap justify-center mt-0.5">
                                      <span className="px-1.5 py-0.5 rounded-full bg-orange-100 text-orange-700 text-[10px] font-semibold">
                                        ลบสินค้า {dscmRows.length}
                                      </span>
                                    </div>
                                  </button>
                                );
                              })()}

                              {/* Card 3 — รอแก้ไข (info, not a tab) */}
                              <div className={`rounded-xl border-2 p-4 flex flex-col items-center gap-1.5 text-center ${
                                pendingTotal > 0
                                  ? "border-amber-300 bg-amber-50"
                                  : "border-green-300 bg-green-50"
                              }`}>
                                <span className="text-[11px] font-semibold text-slate-500">รอแก้ไข</span>
                                <span className={`text-3xl font-bold ${pendingTotal > 0 ? "text-amber-600" : "text-green-600"}`}>
                                  {pendingTotal}
                                </span>
                                <span className="text-[10px] text-slate-400">แถวที่ยังไม่สมบูรณ์</span>
                                {pendingTotal > 0 ? (
                                  <div className="flex flex-col gap-0.5 mt-0.5">
                                    {pendingNdim > 0 && (
                                      <span className="px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-700 text-[10px] font-semibold">
                                        NDIM {pendingNdim}
                                      </span>
                                    )}
                                    {pendingNscm > 0 && (
                                      <span className="px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-700 text-[10px] font-semibold">
                                        SCM {pendingNscm}
                                      </span>
                                    )}
                                    {pendingDscm > 0 && (
                                      <span className="px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-700 text-[10px] font-semibold">
                                        DEL {pendingDscm}
                                      </span>
                                    )}
                                  </div>
                                ) : (
                                  <span className="text-[10px] text-green-600 font-semibold mt-0.5">พร้อม Download ✓</span>
                                )}
                              </div>
                            </div>

                            {/* ── Table content ─────────────────────────────── */}
                            <div className="border border-slate-200 rounded-xl mb-6">
                              <div className="px-4 py-2 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
                                <span className="text-xs font-semibold text-slate-700">
                                  {fillTabs[previewTab]?.displayName}
                                </span>
                                <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-pink-100 text-[#E91E8C] font-bold">
                                  {fillTabs[previewTab]?.rows.length ?? 0} แถว
                                </span>
                              </div>
                              <div className="p-3 bg-white">
                                {(fillTabs[previewTab]?.rows.length ?? 0) === 0 ? (
                                  <p className="text-center text-amber-600 text-sm py-3 flex items-center justify-center gap-2">
                                    <AlertTriangle className="w-4 h-4" />
                                    ไม่มีข้อมูลที่จะเติมในชีทนี้
                                  </p>
                                ) : fillTabs[previewTab] ? (
                                  <FillEditTable
                                    key={previewTab}
                                    colDefs={fillTabs[previewTab].colDefs}
                                    rows={fillTabs[previewTab].rows}
                                    onChange={(updated) => handleFillTabChange(previewTab, updated)}
                                    isKeyZone={previewTab === 0
                                      ? (row, zone) =>
                                          zone === "new" ? !!row.fields.seqNew
                                          : zone === "del" ? !!row.fields.seqDel
                                          : false
                                      : undefined
                                    }
                                    isIncompleteRow={
                                      previewTab === 0
                                        ? (row) =>
                                            (!!row.fields.seqNew && !row.fields.statusNew) ||
                                            (!!row.fields.seqDel && !row.fields.statusDel)
                                        : previewTab === 1
                                        ? (row) => !row.fields.division
                                        : (row) => !row.fields.status
                                    }
                                    getOptions={(field, draft) => {
                                      const tab = fillTabs[previewTab];
                                      if (!tab) return [];
                                      const allVals = (f: string) =>
                                        [...new Set(tab.rows.map(r => r.fields[f]).filter(Boolean))];
                                      const hm = spacemanValues.hierarchyMap;
                                      if (previewTab === 1) {
                                        switch (field) {
                                          case "division":  return spacemanValues.descAList;
                                          case "dept":      return draft.division && hm.divToDept[draft.division]
                                            ? hm.divToDept[draft.division] : spacemanValues.descBList;
                                          case "subDept":   return draft.dept && hm.deptToSub[draft.dept]
                                            ? hm.deptToSub[draft.dept] : spacemanValues.descCList;
                                          case "cls":       return draft.subDept && hm.subToCls[draft.subDept]
                                            ? hm.subToCls[draft.subDept] : spacemanValues.categories;
                                          case "planogram": return allVals("planogram");
                                          case "status":    return NEW_STATUS_OPTIONS;
                                          default:          return allVals(field);
                                        }
                                      }
                                      switch (field) {
                                        case "statusNew":  return NEW_STATUS_OPTIONS;
                                        case "statusDel":  return DEL_STATUS_OPTIONS;
                                        case "status":     return previewTab === 2 ? DEL_STATUS_OPTIONS : NEW_STATUS_OPTIONS;
                                        case "division":   return spacemanValues.descAList;
                                        case "category": {
                                          const cats = allVals("category");
                                          if (draft.division) {
                                            const pfx = draft.division.split(":")[0].trim();
                                            const filtered = cats.filter(c => c.startsWith(pfx));
                                            return filtered.length > 0 ? filtered : cats;
                                          }
                                          return cats;
                                        }
                                        default: return allVals(field);
                                      }
                                    }}
                                  />
                                ) : null}
                              </div>
                            </div>
                          </>
                        )}

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

                {/* Step 6 — Queued confirmation */}
                {step === 6 && (
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

        </div>
      </div>

      {/* Fixed queue panel — overlays content, doesn't affect layout */}
      {jobs.length > 0 && (
        <FixedQueuePanel
          jobs={jobs}
          open={queuePanelOpen}
          onOpenChange={setQueuePanelOpen}
          onTerminate={terminateJob}
          onRemove={removeJob}
          onDownload={downloadJob}
        />
      )}

      {/* Help modal */}
      {showHelp && <HelpModal onClose={() => setShowHelp(false)} />}

      {/* Config modal */}
      {showConfig && (
        <ConfigMenu
          config={exceptionConfig}
          onChange={handleConfigChange}
          onClose={() => setShowConfig(false)}
          categories={spacemanValues.categories}
          subcategories={spacemanValues.subcategories}
          descCList={spacemanValues.descCList}
          descCToCategory={spacemanValues.hierarchyMap.subToCls}
          categoryToSubcategory={spacemanValues.catToSub}
          spacemanLoaded={spacemanLoaded}
          syncStatus={configSyncStatus}
          lastSaved={configLastSaved}
          syncError={configSyncError}
        />
      )}
    </main>
  );
}

// ─── Fixed Queue Panel (overlays, no layout impact) ───────────────────────

function FixedQueuePanel({
  jobs,
  open,
  onOpenChange,
  onTerminate,
  onRemove,
  onDownload,
}: {
  jobs: BuildJob[];
  open: boolean;
  onOpenChange: (v: boolean) => void;
  onTerminate: (id: string) => void;
  onRemove: (id: string) => void;
  onDownload: (id: string, label: string, buffer: ArrayBuffer) => void;
}) {
  const activeCount = jobs.filter(
    (j) => j.status === "queued" || j.status === "processing"
  ).length;
  const doneCount = jobs.filter((j) => j.status === "done").length;

  return (
    <div className="fixed right-0 top-1/3 z-50 -translate-y-1/2">
      {!open ? (
        /* ── Collapsed tab: flush against right edge, no hidden sibling taking space ── */
        <button
          onClick={() => onOpenChange(true)}
          title="เปิดคิว Build ไฟล์"
          className="flex flex-col items-center gap-1.5 px-2 py-3
            bg-white border border-r-0 border-pink-200 rounded-l-xl shadow-lg
            text-[#E91E8C] hover:bg-pink-50 transition-colors"
        >
          <ListOrdered className="w-4 h-4" />
          <span className="text-xs font-bold leading-none">
            {activeCount > 0 ? activeCount : jobs.length}
          </span>
          {activeCount > 0 ? (
            <div className="w-2 h-2 rounded-full bg-[#E91E8C] animate-pulse" />
          ) : doneCount > 0 ? (
            <div className="w-2 h-2 rounded-full bg-green-400" />
          ) : null}
        </button>
      ) : (
        /* ── Expanded panel: flush against right edge ── */
        <div
          className="w-72 bg-white rounded-l-2xl shadow-2xl border border-r-0 border-pink-100
            overflow-hidden flex flex-col"
          style={{ maxHeight: "calc(100vh - 120px)" }}
        >
          <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100] flex-shrink-0" />
          <div className="px-4 py-3 border-b border-pink-50 flex items-center justify-between flex-shrink-0">
            <div className="flex items-center gap-2">
              <ListOrdered className="w-4 h-4 text-[#E91E8C]" />
              <h3 className="font-bold text-slate-800 text-sm">คิว Build ไฟล์</h3>
              {activeCount > 0 && (
                <span className="bg-pink-100 text-[#E91E8C] text-xs font-bold px-2 py-0.5 rounded-full animate-pulse">
                  {activeCount}
                </span>
              )}
            </div>
            <button
              onClick={() => onOpenChange(false)}
              className="p-1 rounded-lg hover:bg-slate-100 text-slate-400 transition-colors"
              title="ซ่อนแผง"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
          <div className="overflow-y-auto flex-1 divide-y divide-slate-100">
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
      )}
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
          {/* Download — available after build completes; buffer kept for re-download */}
          {(job.status === "done" || job.status === "downloaded") && job.buffer && (
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

// ─── Help Modal ───────────────────────────────────────────────────────────────

function HelpModal({ onClose }: { onClose: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100] flex-shrink-0" />
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between flex-shrink-0">
          <div className="flex items-center gap-3">
            <BookOpen className="w-5 h-5 text-[#E91E8C]" />
            <h2 className="font-bold text-slate-800 text-lg">วิธีการทำงานของระบบ</h2>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-400 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Content */}
        <div className="overflow-y-auto flex-1 px-6 py-5 space-y-6 text-sm text-slate-700">

          {/* Step-by-step */}
          <section>
            <h3 className="font-bold text-slate-800 mb-3 flex items-center gap-2">
              <span className="bg-[#E91E8C] text-white text-xs font-bold rounded-full w-5 h-5 flex items-center justify-center flex-shrink-0">1</span>
              อ่านไฟล์ RECAP
            </h3>
            <p className="text-slate-600 leading-relaxed pl-7">
              ระบบสแกนหาแถวใน Sheet <span className="font-mono bg-slate-100 px-1 rounded">NEW SCM</span> ที่มีบาร์โค้ดในคอลัมน์ D แต่คอลัมน์ F (DIVISION) ยังว่างอยู่ — เหล่านี้คือรายการที่ต้องเติมข้อมูล
            </p>
          </section>

          <section>
            <h3 className="font-bold text-slate-800 mb-3 flex items-center gap-2">
              <span className="bg-[#F15A22] text-white text-xs font-bold rounded-full w-5 h-5 flex items-center justify-center flex-shrink-0">2</span>
              ค้นหาในไฟล์ 100 ช่อง
            </h3>
            <p className="text-slate-600 leading-relaxed pl-7">
              เอาบาร์โค้ดแต่ละตัวไปค้นใน Sheet <span className="font-mono bg-slate-100 px-1 rounded">Base</span> หรือ <span className="font-mono bg-slate-100 px-1 rounded">Input</span> ถ้าพบ → ได้ Sub-Class Code → นำไป look up โครงสร้างสินค้าใน Sheet <span className="font-mono bg-slate-100 px-1 rounded">Sh_ProdStructure</span> เพื่อเอา F/G/H/I และได้ค่า N (MBC Forecast) จากคอลัมน์ DF
            </p>
          </section>

          <section>
            <h3 className="font-bold text-slate-800 mb-3 flex items-center gap-2">
              <span className="bg-[#FFD100] text-white text-xs font-bold rounded-full w-5 h-5 flex items-center justify-center flex-shrink-0">3</span>
              ค้นหาใน DATA_SPACEMAN
            </h3>
            <p className="text-slate-600 leading-relaxed pl-7">
              ใช้ Sub-Class Code prefix (6 หลักแรก) หา PLANOGRAM ที่พบบ่อยที่สุด → ใส่คอลัมน์ J และเอาบาร์โค้ดไปหา TOTAL_UNITS → ใส่คอลัมน์ O (Piece 100%)
            </p>
          </section>

          {/* Priority table */}
          <section>
            <h3 className="font-bold text-slate-800 mb-3">Priority การหาข้อมูล F/G/H/I</h3>
            <div className="rounded-xl border border-slate-200 overflow-hidden">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-600">
                    <th className="px-3 py-2 text-left font-semibold w-12">ลำดับ</th>
                    <th className="px-3 py-2 text-left font-semibold">เงื่อนไข</th>
                    <th className="px-3 py-2 text-left font-semibold">ผลลัพธ์</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  <tr className="bg-green-50">
                    <td className="px-3 py-2.5">
                      <span className="flex items-center gap-1 text-green-700 font-bold">
                        <CheckCircle className="w-3.5 h-3.5" /> 1
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-700">พบบาร์โค้ดในไฟล์ 100 ช่อง</td>
                    <td className="px-3 py-2.5 text-slate-600">F/G/H/I จาก Sub-Class Structure, N จาก MBC Forecast, J จาก DATA_SPACEMAN</td>
                  </tr>
                  <tr className="bg-blue-50">
                    <td className="px-3 py-2.5">
                      <span className="flex items-center gap-1 text-blue-700 font-bold">
                        <Database className="w-3.5 h-3.5" /> 2
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-700">ไม่พบใน 100 ช่อง แต่พบใน DATA_SPACEMAN</td>
                    <td className="px-3 py-2.5 text-slate-600">F/G/H/I จาก DESC_A/B/C/CATEGORY, <span className="font-semibold text-blue-700">N = ว่าง</span></td>
                  </tr>
                  <tr className="bg-red-50">
                    <td className="px-3 py-2.5">
                      <span className="flex items-center gap-1 text-red-500 font-bold">
                        <XCircle className="w-3.5 h-3.5" /> 3
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-700">ไม่พบในทั้งสองแหล่ง</td>
                    <td className="px-3 py-2.5 text-slate-600">กรอกเองด้วยปุ่มแก้ไข ✏️</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>

          {/* Columns */}
          <section>
            <h3 className="font-bold text-slate-800 mb-3">คอลัมน์ที่ระบบเติมให้</h3>
            <div className="rounded-xl border border-slate-200 overflow-hidden">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-600">
                    <th className="px-3 py-2 text-left font-semibold w-10">คอล</th>
                    <th className="px-3 py-2 text-left font-semibold">ชื่อ</th>
                    <th className="px-3 py-2 text-left font-semibold">แหล่งข้อมูล</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-slate-700">
                  {[
                    ["F", "DIVISION",       "Sub-Class Structure (Priority 1) / DESC_A จาก DATA_SPACEMAN (Priority 2)"],
                    ["G", "DEPT",            "Sub-Class Structure / DESC_B"],
                    ["H", "SUB-DEPT",        "Sub-Class Structure / DESC_C"],
                    ["I", "Class",           "Sub-Class Structure / CATEGORY"],
                    ["J", "PLANOGRAM",       "DATA_SPACEMAN — ค่าที่พบบ่อยที่สุดตาม Sub-Class prefix 6 หลัก"],
                    ["N", "MBC Forecast",    "ไฟล์ 100 ช่อง คอลัมน์ DF (ว่างถ้าใช้ Priority 2)"],
                    ["O", "Piece 100%",      "DATA_SPACEMAN คอลัมน์ TOTAL_UNITS"],
                    ["P", "%",               "Config Rules (ตรง CATEGORY/SUBCATEGORY/DESC_C) — default 100%"],
                    ["Q", "Net",             "P% × O คำนวณอัตโนมัติ ไม่ต้องกรอก"],
                  ].map(([col, name, src]) => (
                    <tr key={col}>
                      <td className="px-3 py-2">
                        <span className="font-mono font-bold text-[#E91E8C]">{col}</span>
                      </td>
                      <td className="px-3 py-2 font-medium">{name}</td>
                      <td className="px-3 py-2 text-slate-500">{src}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {/* Status legend */}
          <section>
            <h3 className="font-bold text-slate-800 mb-3">สถานะแถวในตารางผลลัพธ์</h3>
            <div className="space-y-2">
              {[
                { badge: "bg-green-100 text-green-700",  icon: <CheckCircle className="w-3.5 h-3.5" />, label: "ยืนยันแล้ว",     desc: "พบในไฟล์ 100 ช่อง และมี PLANOGRAM ใน DATA_SPACEMAN" },
                { badge: "bg-amber-100 text-amber-700",  icon: <AlertTriangle className="w-3.5 h-3.5" />, label: "ไม่มี Planogram", desc: "พบในไฟล์ 100 ช่อง แต่ PLANOGRAM ไม่พบใน DATA_SPACEMAN (คอลัมน์ J จะว่าง)" },
                { badge: "bg-blue-100 text-blue-700",    icon: <Database className="w-3.5 h-3.5" />, label: "จาก Spaceman",    desc: "ไม่พบในไฟล์ 100 ช่อง แต่พบใน DATA_SPACEMAN — F/G/H/I ได้มา คอลัมน์ N ว่าง" },
                { badge: "bg-red-100 text-red-500",      icon: <XCircle className="w-3.5 h-3.5" />, label: "ไม่พบ",           desc: "ไม่พบทั้งในไฟล์ 100 ช่อง และ DATA_SPACEMAN — กรอกเองด้วยปุ่ม ✏️" },
              ].map(({ badge, icon, label, desc }) => (
                <div key={label} className="flex items-start gap-3">
                  <span className={`flex items-center gap-1 px-2 py-1 rounded-full text-xs font-semibold flex-shrink-0 ${badge}`}>
                    {icon}{label}
                  </span>
                  <span className="text-slate-500 text-xs pt-1">{desc}</span>
                </div>
              ))}
            </div>
          </section>

          {/* Config Rules */}
          <section>
            <h3 className="font-bold text-slate-800 mb-2">Config Rules (คอลัมน์ P)</h3>
            <p className="text-slate-600 leading-relaxed">
              ตั้งค่าเปอร์เซ็นต์สำหรับสินค้าบางกลุ่ม โดย filter ด้วย CATEGORY / SUBCATEGORY / DESC_C
              (ใช้ &quot;ทั้งหมด&quot; เพื่อ match ทุกค่า) — Rule แรกที่ตรงชนะ หากไม่มี Rule ใดตรง → ใช้ค่า default 100%
            </p>
          </section>

        </div>
      </div>
    </div>
  );
}

function StatCard({
  label,
  value,
  color,
}: {
  label: string;
  value: number;
  color: "green" | "amber" | "red" | "blue";
}) {
  const colors = {
    green: "bg-green-50 border-green-200 text-green-700",
    amber: "bg-amber-50 border-amber-200 text-amber-700",
    red:   "bg-red-50 border-red-200 text-red-700",
    blue:  "bg-blue-50 border-blue-200 text-blue-700",
  };
  return (
    <div className={`${colors[color]} border rounded-xl p-4 text-center`}>
      <div className="text-3xl font-bold">{value}</div>
      <div className="text-sm font-medium mt-1">{label}</div>
    </div>
  );
}
