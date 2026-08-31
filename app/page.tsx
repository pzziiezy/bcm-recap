"use client";

import { useRef, useState, useEffect, useCallback, startTransition } from "react";
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
import NewRenovateTab from "@/components/NewRenovateTab";
import SpacemanMaster, {
  DriveFileInfo,
  formatDateTime,
  type SpacemanValues,
} from "@/components/SpacemanMaster";
import ConfigMenu, { EXCEPTION_CONFIG_KEY, type SyncStatus } from "@/components/ConfigMenu";
import FillEditTable, { type TabColDef, type EditableFillRow } from "@/components/FillEditTable";
import {
  parseXlsbFiles,
  buildStructureLookup,
  parsePlanogramLookup,
  parseCheckSpace,
  parseFileIndex,
} from "@/lib/processor";
import { buildMinorReportSheets } from "@/lib/minorReport";
import type {
  ExceptionConfig,
  IndexLookup,
  MinorReportSheets,
  MinorReportNewItemRow,
  MinorReportNewNotLinkRow,
  MinorReportDeleteItemRow,
} from "@/lib/types";
import { makeEntry, sendLog } from "@/lib/logger";

// ─── Types ─────────────────────────────────────────────────────────────────

const STEPS = [
  { id: 1, label: "Check Space" },
  { id: 2, label: "FILE_INDEX" },
  { id: 3, label: "Minor Report Template" },
  { id: 4, label: "100 ช่อง" },
  { id: 5, label: "ตรวจสอบ" },
  { id: 6, label: "ดาวน์โหลด" },
];

const MAX_CONCURRENT = 2;

type Status = "idle" | "processing" | "done" | "error";
type AppView = "main" | "spaceman" | "newrenovate";
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

// ─── Check Space status dropdown options ───────────────────────────────────

const NEW_STATUS_OPTIONS = [
  "NEW ADD SOME STORE",
  "NEW ADD ALL STORE",
  "NEW DELETE SOME STORE",
  "NEW DELETE ALL STORE",
  "NEW EXPAND",
];

const DEL_STATUS_OPTIONS = [
  "DELETE SOME STORE",
  "DELETE ALL STORE",
  "DELETE ALL STORE (IN/OUT)",
  "DELETE SOME STORE (IN/OUT)",
];

// ─── Minor Report preview table definitions (Step 5) ───────────────────────
// `col` is unused by FillEditTable itself (only convertToEditableRows/FromEditableRows,
// which Minor Report doesn't need — see minorRowsToEditable/editableToMinorRows below),
// so it's just each field's display order here.

const NEW_ITEM_COLDEFS: TabColDef[] = [
  { field: "upc",                        col: 0,  label: "UPC",        editable: false },
  { field: "name",                       col: 1,  label: "NAME",       editable: false },
  { field: "division",                   col: 2,  label: "DIVISION",   editable: true },
  { field: "department",                 col: 3,  label: "DEPARTMENT", editable: true },
  { field: "salepack",                   col: 4,  label: "SALEPACK",   editable: true },
  { field: "recipe",                     col: 5,  label: "RECIPE",     editable: true },
  { field: "packSize",                   col: 6,  label: "PACK SIZE",  editable: true },
  { field: "totalUnits",                 col: 7,  label: "TOTAL_UNITS", editable: true },
  { field: "purShelfStockPiece",         col: 8,  label: "BCM Shelf stock ON POG (Piece)", editable: true },
  { field: "pctOrdering",                col: 9,  label: "% Ordering", editable: true },
  { field: "netCapacity",                col: 10, label: "Net Capacity", editable: false },
  { field: "attClass",                   col: 11, label: "ATT_CLASS",  editable: true },
  { field: "attCode",                    col: 12, label: "ATT_CODE",   editable: true },
  { field: "storeNumber",                col: 13, label: "STORE NUMBER", editable: false },
  { field: "link",                       col: 14, label: "LINK",       editable: false },
  { field: "forecastSalesPerMonthStore", col: 15, label: "Forecast Sales/Month/Store", editable: true },
  { field: "remark",                     col: 16, label: "REMARK",     editable: true },
];

const NEW_NOT_LINK_COLDEFS: TabColDef[] = [
  { field: "upc",         col: 0, label: "UPC",         editable: false },
  { field: "name",        col: 1, label: "NAME",        editable: false },
  { field: "division",    col: 2, label: "DIVISION",    editable: true },
  { field: "department",  col: 3, label: "DEPARTMENT",  editable: true },
  { field: "attClass",    col: 4, label: "ATT_CLASS",   editable: true },
  { field: "attCode",     col: 5, label: "ATT_CODE",    editable: true },
  { field: "storeNumber", col: 6, label: "STORENUMBER", editable: false },
  { field: "link",        col: 7, label: "LINK",        editable: false },
  { field: "remark",      col: 8, label: "REMARK",      editable: true },
];

const DELETE_ITEM_COLDEFS: TabColDef[] = [
  { field: "upc",         col: 0, label: "UPC",         editable: false },
  { field: "name",        col: 1, label: "NAME",        editable: false },
  { field: "division",    col: 2, label: "DIVISION",    editable: true },
  { field: "department",  col: 3, label: "DEPARTMENT",  editable: true },
  { field: "attClass",    col: 4, label: "ATT_CLASS",   editable: true },
  { field: "attCode",     col: 5, label: "ATT_CODE",    editable: true },
  { field: "storeNumber", col: 6, label: "STORENUMBER", editable: false },
  { field: "link",        col: 7, label: "LINK",        editable: false },
  { field: "remark",      col: 8, label: "REMARK",      editable: true },
];

/** Converts one Minor Report sheet's typed rows into FillEditTable's generic
 *  {rowIndex, fields} shape — rowIndex is just the array index since Minor Report
 *  rows have no underlying spreadsheet row to track. All row fields are plain
 *  strings (a couple are narrowed literal types like link: "LINK"), so the cast
 *  here is just satisfying that, not discarding real type safety. */
function minorRowsToEditable<T>(rows: T[]): EditableFillRow[] {
  return rows.map((row, i) => ({ rowIndex: i, fields: { ...row } as unknown as Record<string, string> }));
}

/** Reverses minorRowsToEditable() after Step 5 edits, preserving row order. */
function editableToMinorRows<T>(rows: EditableFillRow[]): T[] {
  return rows.map(r => r.fields as unknown as T);
}

// ─── Home ──────────────────────────────────────────────────────────────────

export default function Home() {
  const [view, setView] = useState<AppView>("main");
  const [step, setStep] = useState(1);
  const [status, setStatus] = useState<Status>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [pct, setPct] = useState(0);

  const [checkSpaceFile, setCheckSpaceFile] = useState<File | null>(null);
  const [fileIndexFile, setFileIndexFile] = useState<File | null>(null);
  const [templateFiles, setTemplateFiles] = useState<File[]>([]); // Minor Report template (was RECAP)
  const [xlsbFiles, setXlsbFiles] = useState<File[]>([]);
  const [driveFileInfo, setDriveFileInfo] = useState<DriveFileInfo | null>(null);
  const [driveLoading, setDriveLoading] = useState(true);

  // Minor Report preview tables (Step 5) — 3 editable tabs, one per output sheet
  interface MinorTabData {
    displayName: string;
    colDefs: TabColDef[];
    rows: EditableFillRow[];
  }
  const [minorTabs, setMinorTabs]   = useState<MinorTabData[] | null>(null);
  const [previewTab, setPreviewTab] = useState(0);

  // Exception config — start with [] so server and client render identically (no hydration mismatch).
  // localStorage is loaded in useEffect (client-only, after hydration).
  const [exceptionConfig, setExceptionConfig] = useState<ExceptionConfig[]>([]);
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
  const templateBufRef = useRef<ArrayBuffer | null>(null); // uploaded Minor Report template (Step 3)
  const minorReportSheetsRef = useRef<MinorReportSheets | null>(null); // last computed (pre-edit) sheets
  const sessionIdRef = useRef<string>("");
  const pageSessionRef = useRef(`page-${Date.now().toString(36)}`);
  const workersRef = useRef<Map<string, Worker>>(new Map());
  const jobDataRef = useRef<Map<string, { templateBuf: ArrayBuffer; sheets: MinorReportSheets; label: string }>>(new Map());
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

  // Hydration-safe localStorage read — runs only on client, after initial render is committed.
  useEffect(() => {
    try {
      const raw = localStorage.getItem(EXCEPTION_CONFIG_KEY);
      if (raw) setExceptionConfig(JSON.parse(raw) as ExceptionConfig[]);
    } catch { /* ignore */ }
  }, []);

  // Load exception config from Google Sheets on mount (overwrites localStorage cache with fresh data)
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

  const startJobFn = useCallback((id: string, templateBuf: ArrayBuffer, sheets: MinorReportSheets) => {
    const buildSid = `build-${id.slice(0, 8)}`;
    const buildStart = Date.now();

    setJobs((prev) =>
      prev.map((j) =>
        j.id === id ? { ...j, status: "processing", startedAt: new Date(), progress: 5 } : j
      )
    );

    const worker = new Worker(new URL("../lib/minorReportDownload.worker.ts", import.meta.url));
    workersRef.current.set(id, worker);

    worker.onmessage = (e: MessageEvent) => {
      const msg = e.data as { type: string; pct?: number; buffer?: ArrayBuffer; message?: string; missingHeaders?: string[] };
      switch (msg.type) {
        case "init_ok":
          worker.postMessage({ type: "build", sheets });
          break;
        case "progress":
          startTransition(() => {
            setJobs((prev) =>
              prev.map((j) => (j.id === id ? { ...j, progress: msg.pct ?? j.progress } : j))
            );
          });
          break;
        case "done": {
          const label = jobDataRef.current.get(id)?.label ?? buildSid;
          workersRef.current.delete(id);
          jobDataRef.current.delete(id);
          const durSec = ((Date.now() - buildStart) / 1000).toFixed(1);
          sendLog([makeEntry(buildSid, "BUILD_COMPLETE", "INFO",
            `Build ${label} เสร็จในเวลา ${durSec} วินาที`,
            { jobId: id, label, durationSec: durSec }
          )]);
          // A field's expected header text wasn't found in the uploaded template — the
          // column was skipped (never guessed into the wrong place), so warn loudly.
          if (msg.missingHeaders && msg.missingHeaders.length > 0) {
            sendLog([makeEntry(buildSid, "ERROR", "WARN",
              `Build ${label}: หา header ไม่เจอ ${msg.missingHeaders.length} คอลัมน์ (ข้ามไป ไม่เติมข้อมูล): ${msg.missingHeaders.join(", ")}`,
              { jobId: id, label, missingHeaders: msg.missingHeaders }
            )]);
          }
          startTransition(() => {
            setJobs((prev) =>
              prev.map((j) =>
                j.id === id
                  ? { ...j, status: "done", progress: 100, completedAt: new Date(), buffer: msg.buffer }
                  : j
              )
            );
          });
          break;
        }
        case "error": {
          const label = jobDataRef.current.get(id)?.label ?? buildSid;
          workersRef.current.delete(id);
          sendLog([makeEntry(buildSid, "BUILD_FAILED", "ERROR",
            `Build ${label} ล้มเหลว: ${msg.message ?? "Worker error"}`,
            { jobId: id, label, error: msg.message }
          )]);
          startTransition(() => {
            setJobs((prev) =>
              prev.map((j) =>
                j.id === id
                  ? { ...j, status: "failed", error: msg.message ?? "Worker error", completedAt: new Date() }
                  : j
              )
            );
          });
          break;
        }
      }
    };

    worker.onerror = (e: ErrorEvent) => {
      const label = jobDataRef.current.get(id)?.label ?? buildSid;
      workersRef.current.delete(id);
      sendLog([makeEntry(buildSid, "BUILD_FAILED", "ERROR",
        `Build ${label} crash: ${e.message ?? "Worker crashed"}`,
        { jobId: id, label, error: e.message }
      )]);
      startTransition(() => {
        setJobs((prev) =>
          prev.map((j) =>
            j.id === id
              ? { ...j, status: "failed", error: e.message ?? "Worker crashed", completedAt: new Date() }
              : j
          )
        );
      });
    };

    // Transfer buffer to avoid a full copy (slice first to preserve original)
    const buf = templateBuf.slice(0);
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
    startJobFn(next.id, data.templateBuf, data.sheets);
  }, [jobs, startJobFn]);

  // ── Queue actions ───────────────────────────────────────────────────────

  /** Reassembles a MinorReportSheets from the (possibly Step-5-edited) preview tables —
   *  build always uses what's currently on screen, never the raw pre-edit computation. */
  const currentMinorReportSheets = (): MinorReportSheets | null => {
    if (!minorTabs) return minorReportSheetsRef.current;
    return {
      newItem: editableToMinorRows<MinorReportNewItemRow>(minorTabs[0].rows),
      newNotLink: editableToMinorRows<MinorReportNewNotLinkRow>(minorTabs[1].rows),
      deleteItem: editableToMinorRows<MinorReportDeleteItemRow>(minorTabs[2].rows),
    };
  };

  const enqueueJob = () => {
    const sheets = currentMinorReportSheets();
    if (!templateBufRef.current || !sheets) return;
    const id = crypto.randomUUID();
    const num = ++jobCounterRef.current;
    const baseName = templateFiles[0]?.name.replace(/\.[^.]+$/, "") ?? "Minor Report";
    const label = `${baseName}_filled_#${num}.xlsx`;

    jobDataRef.current.set(id, {
      templateBuf: templateBufRef.current.slice(0),
      sheets,
      label,
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
    templateFiles.length === 1 &&
    xlsbFiles.length > 0 &&
    driveFileInfo !== null;

  const handleProcess = async () => {
    if (!checkSpaceFile || !fileIndexFile || !templateFiles[0] || xlsbFiles.length === 0 || !driveFileInfo) return;

    const sessionId = crypto.randomUUID();
    sessionIdRef.current = sessionId;
    const t0 = Date.now();

    setStatus("processing");
    setStep(5);
    setPct(0);

    sendLog([makeEntry(sessionId, "PROCESS_START", "INFO",
      `เริ่มประมวลผล: ${templateFiles[0].name} + ${xlsbFiles.length} ไฟล์ 100 ช่อง + ${driveFileInfo.name}`,
      {
        templateFile: templateFiles[0].name,
        xlsbFiles: xlsbFiles.map((f) => f.name),
        spacemanFile: driveFileInfo.name,
        spacemanFileId: driveFileInfo.id,
      }
    )]);

    try {
      setStatusMsg("อ่านไฟล์ Minor Report template...");
      setPct(5);
      // Just stored — the template's sheets/columns are only inspected at build time,
      // by auto-detecting each sheet's header row (see minorReportDownload.worker.ts).
      const templateBuf = await templateFiles[0].arrayBuffer();
      templateBufRef.current = templateBuf.slice(0);

      setStatusMsg("อ่านไฟล์ Check Space และ FILE_INDEX...");
      setPct(15);
      const [checkSpaceItems, indexLookup] = await Promise.all([
        parseCheckSpace(checkSpaceFile),
        parseFileIndex(fileIndexFile),
      ]);

      sendLog([makeEntry(sessionId, "INDEX_PARSED", "INFO",
        `FILE_INDEX: ${indexLookup.storeList.length} stores, ${indexLookup.pogToByCode.size} POG→BY_CODE`,
        { storeCount: indexLookup.storeList.length, pogByCodeCount: indexLookup.pogToByCode.size }
      )]);

      const newCount = checkSpaceItems.filter(i => !i.status.toUpperCase().startsWith("DELETE")).length;
      const delCount = checkSpaceItems.filter(i => i.status.toUpperCase().startsWith("DELETE")).length;
      sendLog([makeEntry(sessionId, "CHECKSPACE_PARSED", "INFO",
        `Check Space: ${newCount} NEW / ${delCount} DELETE items`,
        { checkSpaceFile: checkSpaceFile.name, fileIndexFile: fileIndexFile.name, totalItems: checkSpaceItems.length }
      )]);

      setStatusMsg("กำลังค้นหาข้อมูลในไฟล์ 100 ช่อง...");
      setPct(30);
      const [barcodeMap, structureMap] = await Promise.all([
        parseXlsbFiles(xlsbFiles),
        buildStructureLookup(xlsbFiles),
      ]);

      sendLog([makeEntry(sessionId, "XLSB_PARSED", "INFO",
        `100 ช่อง: พบบาร์โค้ด ${barcodeMap.size} รายการ, โครงสร้างสินค้า ${structureMap.size} รายการ จาก ${xlsbFiles.length} ไฟล์`,
        { files: xlsbFiles.map((f) => f.name), barcodesFound: barcodeMap.size, structureFound: structureMap.size }
      )]);

      setStatusMsg("กำลังตรวจสอบข้อมูลจาก DATA_SPACEMAN...");
      setPct(50);
      const res = await fetch(`/api/spaceman/file?id=${driveFileInfo.id}`);
      if (!res.ok) throw new Error("ไม่สามารถดาวน์โหลดไฟล์ DATA_SPACEMAN จาก Google Drive ได้");
      const buf = await res.arrayBuffer();
      const spacemanFile = new File([buf], driveFileInfo.name, {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      setStatusMsg("อ่าน DATA_SPACEMAN...");
      setPct(60);
      const planogramResult = await parsePlanogramLookup(spacemanFile, (p) =>
        setPct(60 + p * 0.25)
      );

      sendLog([makeEntry(sessionId, "SPACEMAN_PARSED", "INFO",
        `DATA_SPACEMAN: พบ prefix ${planogramResult.byPrefix.size} รายการ, UPC ${planogramResult.byUpc.size} รายการ`,
        { prefixCount: planogramResult.byPrefix.size, upcCount: planogramResult.byUpc.size, filename: driveFileInfo.name }
      )]);

      setStatusMsg("กำลังสร้าง Minor Report...");
      setPct(90);
      const sheets = buildMinorReportSheets({
        checkSpaceItems,
        indexLookup,
        barcodeMap,
        structureMap,
        byUpc: planogramResult.byUpc,
        exceptionConfig,
      });
      minorReportSheetsRef.current = sheets;

      const durSec = ((Date.now() - t0) / 1000).toFixed(1);
      sendLog([makeEntry(sessionId, "PROCESS_COMPLETE", "INFO",
        `เสร็จใน ${durSec}s — Recap_New_item ${sheets.newItem.length} | Recap_New_not_link ${sheets.newNotLink.length} | Recap_Delete_item ${sheets.deleteItem.length}`,
        {
          newItemRows: sheets.newItem.length,
          newNotLinkRows: sheets.newNotLink.length,
          deleteItemRows: sheets.deleteItem.length,
          durationSec: durSec,
        }
      )]);

      setMinorTabs([
        { displayName: "Recap_New_item",     colDefs: NEW_ITEM_COLDEFS,     rows: minorRowsToEditable(sheets.newItem) },
        { displayName: "Recap_New_not_link", colDefs: NEW_NOT_LINK_COLDEFS, rows: minorRowsToEditable(sheets.newNotLink) },
        { displayName: "Recap_Delete_item",  colDefs: DELETE_ITEM_COLDEFS,  rows: minorRowsToEditable(sheets.deleteItem) },
      ]);
      setPreviewTab(0);

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

  const handleFillTabChange = (tabIdx: number, updatedRows: EditableFillRow[]) => {
    setMinorTabs(prev => {
      if (!prev) return prev;
      const next = [...prev];
      next[tabIdx] = { ...next[tabIdx], rows: updatedRows };
      return next;
    });
  };

  const reset = () => {
    // Jobs persist across resets — do NOT clear them
    templateBufRef.current = null;
    minorReportSheetsRef.current = null;
    setStep(1);
    setStatus("idle");
    setStatusMsg("");
    setPct(0);
    setCheckSpaceFile(null);
    setFileIndexFile(null);
    setTemplateFiles([]);
    setXlsbFiles([]);
    setMinorTabs(null);
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

  // ─── Minor Report tab summary stats ─────────────────────────────────────────
  const newItemRows    = minorTabs?.[0]?.rows ?? [];
  const newNotLinkRows = minorTabs?.[1]?.rows ?? [];
  const deleteItemRows = minorTabs?.[2]?.rows ?? [];

  // "Incomplete" = missing DIVISION, the same basic-enrichment signal used across all 3 sheets
  const pendingNewItem    = newItemRows.filter(r => !r.fields.division).length;
  const pendingNewNotLink = newNotLinkRows.filter(r => !r.fields.division).length;
  const pendingDeleteItem = deleteItemRows.filter(r => !r.fields.division).length;
  const pendingTotal = pendingNewItem + pendingNewNotLink + pendingDeleteItem;

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
            <TabBtn active={view === "newrenovate"} onClick={() => setView("newrenovate")}>
              <FileSpreadsheet className="w-4 h-4" />
              New and Renovate
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

            {/* New and Renovate tab — negative margins cancel outer py-8 (32px) + space-y-8 (32px) */}
            {view === "newrenovate" && (
              <div className="-mx-6 -mt-16">
                <NewRenovateTab exceptionConfig={exceptionConfig} />
              </div>
            )}

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
                      onFiles={(files) => {
                        setCheckSpaceFile(files[0] ?? null);
                        if (files[0]) sendLog([makeEntry(pageSessionRef.current, "FILE_UPLOAD", "INFO",
                          `อัปโหลด Check Space: ${files[0].name}`,
                          { fileType: "Check Space", name: files[0].name, sizeMB: (files[0].size / 1048576).toFixed(2) }
                        )]);
                      }}
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
                      onFiles={(files) => {
                        setFileIndexFile(files[0] ?? null);
                        if (files[0]) sendLog([makeEntry(pageSessionRef.current, "FILE_UPLOAD", "INFO",
                          `อัปโหลด FILE_INDEX: ${files[0].name}`,
                          { fileType: "FILE_INDEX", name: files[0].name, sizeMB: (files[0].size / 1048576).toFixed(2) }
                        )]);
                      }}
                    />
                    <div className="flex gap-3">
                      <NavBtn variant="outline" onClick={() => setStep(1)}>← ย้อนกลับ</NavBtn>
                      <NavBtn onClick={() => setStep(3)} disabled={!fileIndexFile}>
                        ถัดไป →
                      </NavBtn>
                    </div>
                  </Card>
                )}

                {/* Step 3 — Upload Minor Report template */}
                {step === 3 && (
                  <Card title="Step 3 - อัปโหลด Minor Report Template">
                    <DropZone
                      label="TO BE_Minor Report.xlsx"
                      hint="ไฟล์ template เปล่า 3 ชีต (Recap_New_item / Recap_New_not_link / Recap_Delete_item) — format/สี/เส้นขอบในไฟล์นี้จะถูกเก็บไว้เป๊ะๆ"
                      accept=".xlsx,.xls"
                      files={templateFiles}
                      onFiles={(files) => {
                        setTemplateFiles(files);
                        if (files.length > 0) sendLog([makeEntry(pageSessionRef.current, "FILE_UPLOAD", "INFO",
                          `อัปโหลด Minor Report template: ${files.map(f => f.name).join(", ")}`,
                          { fileType: "MinorReportTemplate", files: files.map(f => ({ name: f.name, sizeMB: (f.size / 1048576).toFixed(2) })) }
                        )]);
                      }}
                    />
                    <div className="flex gap-3">
                      <NavBtn variant="outline" onClick={() => setStep(2)}>← ย้อนกลับ</NavBtn>
                      <NavBtn onClick={() => setStep(4)} disabled={templateFiles.length !== 1}>
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
                      onFiles={(files) => {
                        setXlsbFiles(files);
                        if (files.length > 0) sendLog([makeEntry(pageSessionRef.current, "FILE_UPLOAD", "INFO",
                          `อัปโหลด 100 ช่อง: ${files.length} ไฟล์ (${files.map(f => f.name).join(", ")})`,
                          { fileType: "100 ช่อง (XLSB)", count: files.length, files: files.map(f => ({ name: f.name, sizeMB: (f.size / 1048576).toFixed(2) })) }
                        )]);
                      }}
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
                        {minorTabs && (
                          <>
                            {/* ── KPI Summary Cards (also serve as tab nav) ─── */}
                            <div className="grid grid-cols-4 gap-3 mb-4">
                              {/* Card 0 — Recap_New_item */}
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
                                    <span className="text-[11px] font-semibold text-slate-500 truncate w-full text-center">Recap_New_item</span>
                                    <span className={`text-3xl font-bold ${active ? "text-[#E91E8C]" : "text-slate-700"}`}>
                                      {newItemRows.length}
                                    </span>
                                    <span className="text-[10px] text-slate-400">แถวทั้งหมด (LINK)</span>
                                  </button>
                                );
                              })()}

                              {/* Card 1 — Recap_New_not_link */}
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
                                    <span className="text-[11px] font-semibold text-slate-500">Recap_New_not_link</span>
                                    <span className={`text-3xl font-bold ${active ? "text-blue-600" : "text-slate-700"}`}>
                                      {newNotLinkRows.length}
                                    </span>
                                    <span className="text-[10px] text-slate-400">แถวทั้งหมด (ยังไม่ link)</span>
                                  </button>
                                );
                              })()}

                              {/* Card 2 — Recap_Delete_item */}
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
                                    <span className="text-[11px] font-semibold text-slate-500">Recap_Delete_item</span>
                                    <span className={`text-3xl font-bold ${active ? "text-orange-500" : "text-slate-700"}`}>
                                      {deleteItemRows.length}
                                    </span>
                                    <span className="text-[10px] text-slate-400">แถวทั้งหมด</span>
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
                                <span className="text-[10px] text-slate-400">แถวที่ยังไม่สมบูรณ์ (ไม่มี DIVISION)</span>
                                {pendingTotal === 0 && (
                                  <span className="text-[10px] text-green-600 font-semibold mt-0.5">พร้อม Download ✓</span>
                                )}
                              </div>
                            </div>

                            {/* ── Table content ─────────────────────────────── */}
                            <div className="border border-slate-200 rounded-xl mb-6">
                              <div className="px-4 py-2 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
                                <span className="text-xs font-semibold text-slate-700">
                                  {minorTabs[previewTab]?.displayName}
                                </span>
                                <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-pink-100 text-[#E91E8C] font-bold">
                                  {minorTabs[previewTab]?.rows.length ?? 0} แถว
                                </span>
                              </div>
                              <div className="p-3 bg-white">
                                {(minorTabs[previewTab]?.rows.length ?? 0) === 0 ? (
                                  <p className="text-center text-amber-600 text-sm py-3 flex items-center justify-center gap-2">
                                    <AlertTriangle className="w-4 h-4" />
                                    ไม่มีข้อมูลในชีทนี้
                                  </p>
                                ) : minorTabs[previewTab] ? (
                                  <FillEditTable
                                    key={previewTab}
                                    colDefs={minorTabs[previewTab].colDefs}
                                    rows={minorTabs[previewTab].rows}
                                    onChange={(updated) => handleFillTabChange(previewTab, updated)}
                                    onEditSaved={(rowIndex, changes) => {
                                      const tabName = minorTabs[previewTab]?.displayName ?? `tab${previewTab}`;
                                      sendLog([makeEntry(sessionIdRef.current, "USER_EDIT", "INFO",
                                        `แก้ไขแถว ${rowIndex} ใน ${tabName}: ${Object.keys(changes).join(", ")}`,
                                        { tab: tabName, rowIndex, changes }
                                      )]);
                                    }}
                                    onReplaceApplied={(col, from, to, count) => {
                                      const tabName = minorTabs[previewTab]?.displayName ?? `tab${previewTab}`;
                                      const colLabel = minorTabs[previewTab]?.colDefs.find(d => d.field === col)?.label ?? col;
                                      sendLog([makeEntry(sessionIdRef.current, "USER_REPLACE", "INFO",
                                        `Replace ใน ${tabName} คอลัมน์ "${colLabel}": "${from}" → "${to}" (${count} แถว)`,
                                        { tab: tabName, col, colLabel, from, to, count }
                                      )]);
                                    }}
                                    isIncompleteRow={(row) => !row.fields.division}
                                    getOptions={(field, draft) => {
                                      const tab = minorTabs[previewTab];
                                      if (!tab) return [];
                                      const allVals = (f: string) =>
                                        [...new Set(tab.rows.map(r => r.fields[f]).filter(Boolean))];
                                      const hm = spacemanValues.hierarchyMap;
                                      switch (field) {
                                        case "division":
                                          return spacemanValues.descAList;
                                        case "department":
                                          return draft.division && hm.divToDept[draft.division]
                                            ? hm.divToDept[draft.division] : spacemanValues.descBList;
                                        case "link":
                                          return previewTab === 0 ? ["LINK"] : previewTab === 1 ? ["New not link"] : ["NOT LINK"];
                                        default:
                                          return allVals(field);
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

          {/* Overview */}
          <section className="rounded-xl bg-pink-50 border border-pink-100 px-4 py-3">
            <p className="text-slate-700 leading-relaxed">
              ระบบนี้ช่วย<strong>เติมข้อมูลลงใน 3 ชีต</strong>ของไฟล์ RECAP โดยอัตโนมัติ
              โดยดึงข้อมูลจากไฟล์ Check Space, ไฟล์ 100 ช่อง และ DATA_SPACEMAN แล้วให้ผู้ใช้ตรวจสอบ/แก้ไขก่อน Download
            </p>
          </section>

          {/* Upload steps */}
          <section>
            <h3 className="font-bold text-slate-800 mb-3">ขั้นตอนการใช้งาน</h3>
            <div className="space-y-2">
              {[
                { step: "1", color: "bg-pink-500",   label: "อัปโหลด Check Space",   desc: "ไฟล์ที่มีรายการสินค้า NEW / DELETE พร้อม POG matrix" },
                { step: "2", color: "bg-orange-400", label: "อัปโหลด FILE_INDEX",    desc: "ไฟล์ที่ map POG NAME → BY_CODE และ store flags" },
                { step: "3", color: "bg-amber-400",  label: "อัปโหลดไฟล์ RECAP",    desc: "ไฟล์ที่ต้องการเติมข้อมูล (NEW SCM, DEL SCM, NEW_DELETE_IM)" },
                { step: "4", color: "bg-emerald-500",label: "อัปโหลดไฟล์ 100 ช่อง", desc: "ไฟล์ XLSB สำหรับค้นหา Sub-Class Code และ MBC Forecast" },
                { step: "5", color: "bg-blue-500",   label: "ตรวจสอบ & แก้ไข",      desc: "ดูข้อมูลที่ระบบเติมให้ แก้ไขได้ทีละแถว หรือใช้ Replace แบบ Bulk" },
                { step: "6", color: "bg-slate-500",  label: "Download ไฟล์",         desc: "ไฟล์ RECAP ที่เติมข้อมูลครบแล้ว พร้อมใช้งาน" },
              ].map(({ step, color, label, desc }) => (
                <div key={step} className="flex items-start gap-3">
                  <span className={`${color} text-white text-xs font-bold rounded-full w-6 h-6 flex items-center justify-center flex-shrink-0 mt-0.5`}>{step}</span>
                  <div>
                    <span className="font-semibold text-slate-800">{label}</span>
                    <span className="text-slate-500 ml-2">{desc}</span>
                  </div>
                </div>
              ))}
            </div>
          </section>

          {/* 3 sheets */}
          <section>
            <h3 className="font-bold text-slate-800 mb-3">ชีตที่ระบบเติมข้อมูล</h3>
            <div className="space-y-3">

              {/* NDIM */}
              <div className="rounded-xl border border-emerald-200 bg-emerald-50/40 px-4 py-3">
                <p className="font-semibold text-emerald-800 mb-1.5">📋 NEW_DELETE_IM</p>
                <p className="text-xs text-slate-600 leading-relaxed mb-2">
                  เพิ่มแถวสินค้าจาก Check Space แบ่งเป็น 2 ฝั่ง — ฝั่ง New (สินค้าใหม่) และ ฝั่ง Del (สินค้าที่ต้อง Delete)
                </p>
                <div className="grid grid-cols-2 gap-2 text-xs">
                  <div className="bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2">
                    <p className="font-semibold text-emerald-700 mb-1">ฝั่ง New (A–G)</p>
                    <p className="text-slate-600">ลำดับ · Barcode · Name · DC · BY_CODE · Status · Remark</p>
                  </div>
                  <div className="bg-rose-50 border border-rose-200 rounded-lg px-3 py-2">
                    <p className="font-semibold text-rose-700 mb-1">ฝั่ง Del (I–O)</p>
                    <p className="text-slate-600">ลำดับ · Barcode · Name · DC · BY_CODE · Status · Extra_Info</p>
                  </div>
                </div>
                <p className="text-xs text-slate-500 mt-2">
                  Extra_Info: ดึงจากไฟล์ 100 ช่อง (col Extra Info) → ถ้าไม่พบ ใช้ Check Space col D
                </p>
              </div>

              {/* NEW SCM */}
              <div className="rounded-xl border border-blue-200 bg-blue-50/40 px-4 py-3">
                <p className="font-semibold text-blue-800 mb-1.5">📝 NEW SCM</p>
                <p className="text-xs text-slate-600 leading-relaxed mb-2">
                  เพิ่มแถวสินค้าใหม่จาก Check Space แล้วเติมข้อมูล hierarchy และ planogram โดยอัตโนมัติ
                </p>
                <div className="flex flex-wrap gap-1.5 text-xs">
                  {[
                    { col: "F", name: "DIVISION" }, { col: "G", name: "DEPT" },
                    { col: "H", name: "SUB-DEPT" }, { col: "I", name: "Class" },
                    { col: "J", name: "PLANOGRAM" }, { col: "N", name: "MBC FCST" },
                    { col: "O", name: "Piece" }, { col: "P", name: "%" },
                    { col: "Q", name: "Net (คำนวณ)" },
                  ].map(({ col, name }) => (
                    <span key={col} className="bg-white border border-blue-200 rounded px-2 py-0.5 text-slate-700">
                      <span className="font-bold text-[#E91E8C]">{col}</span> {name}
                    </span>
                  ))}
                </div>
              </div>

              {/* DEL SCM */}
              <div className="rounded-xl border border-orange-200 bg-orange-50/40 px-4 py-3">
                <p className="font-semibold text-orange-800 mb-1.5">🗑 DEL SCM</p>
                <p className="text-xs text-slate-600 leading-relaxed mb-2">
                  เพิ่มแถวสินค้าที่ต้อง Delete จาก Check Space พร้อมเติม store flags อัตโนมัติ
                </p>
                <p className="text-xs text-slate-600">
                  ลำดับ · Barcode · Name · Division · POG ROUND (Planogram) · Status · Extra_Info · Store flags
                </p>
                <p className="text-xs text-slate-500 mt-1">
                  DELETE ALL STORE → ไม่เติม store flags | DELETE SOME STORE → เติม store ที่เกี่ยวข้อง
                </p>
              </div>
            </div>
          </section>

          {/* Priority for NEW SCM */}
          <section>
            <h3 className="font-bold text-slate-800 mb-1.5">Priority การค้นหาข้อมูล (NEW SCM)</h3>
            <p className="text-xs text-slate-500 mb-3">ระบบค้นหาข้อมูล F/G/H/I ตามลำดับ Priority ดังนี้</p>
            <div className="rounded-xl border border-slate-200 overflow-hidden">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-600">
                    <th className="px-3 py-2 text-left font-semibold w-10">#</th>
                    <th className="px-3 py-2 text-left font-semibold">เงื่อนไข</th>
                    <th className="px-3 py-2 text-left font-semibold">ผลลัพธ์</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  <tr className="bg-green-50">
                    <td className="px-3 py-2.5">
                      <span className="flex items-center gap-1 text-green-700 font-bold"><CheckCircle className="w-3.5 h-3.5" />1</span>
                    </td>
                    <td className="px-3 py-2.5 font-medium">พบใน<strong>ไฟล์ 100 ช่อง</strong></td>
                    <td className="px-3 py-2.5 text-slate-600">F/G/H/I จาก Sub-Class Structure · J จาก DATA_SPACEMAN · N จาก MBC Forecast</td>
                  </tr>
                  <tr className="bg-blue-50">
                    <td className="px-3 py-2.5">
                      <span className="flex items-center gap-1 text-blue-700 font-bold"><Database className="w-3.5 h-3.5" />2</span>
                    </td>
                    <td className="px-3 py-2.5 font-medium">ไม่พบใน 100 ช่อง แต่พบใน<strong>DATA_SPACEMAN</strong></td>
                    <td className="px-3 py-2.5 text-slate-600">F/G/H/I จาก DESC_A/B/C/CATEGORY · <span className="font-semibold text-blue-700">N = ว่าง</span></td>
                  </tr>
                  <tr className="bg-red-50">
                    <td className="px-3 py-2.5">
                      <span className="flex items-center gap-1 text-red-500 font-bold"><XCircle className="w-3.5 h-3.5" />3</span>
                    </td>
                    <td className="px-3 py-2.5 font-medium">ไม่พบในทั้งสองแหล่ง</td>
                    <td className="px-3 py-2.5 text-slate-600">กรอกเองด้วยปุ่มแก้ไข ✏️ ใน Step 5</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>

          {/* Step 5 edit */}
          <section>
            <h3 className="font-bold text-slate-800 mb-2">การตรวจสอบและแก้ไขใน Step 5</h3>
            <div className="space-y-1.5 text-xs text-slate-600">
              <p>• <strong>แก้ไขทีละแถว</strong> — กดปุ่ม ✏️ ที่แถวที่ต้องการ แก้ไข แล้วกด ✓ บันทึก</p>
              <p>• <strong>Replace แบบ Bulk</strong> — กดปุ่ม Replace เพื่อค้นหาและแทนที่ค่าในคอลัมน์ที่เลือก ทีเดียวหลายแถว</p>
              <p>• <strong>Filter แถวที่ยังไม่สมบูรณ์</strong> — กดปุ่ม "ยังไม่สมบูรณ์" เพื่อดูเฉพาะแถวที่ยังขาดข้อมูลสำคัญ</p>
              <p>• ข้อมูลที่แก้ไขใน Step 5 จะถูกบันทึกลงไฟล์ RECAP เมื่อกด Download</p>
            </div>
          </section>

          {/* Config Rules */}
          <section>
            <h3 className="font-bold text-slate-800 mb-2">Config Rules (คอลัมน์ P — %)</h3>
            <p className="text-slate-600 leading-relaxed text-xs">
              กำหนดเปอร์เซ็นต์ (%) สำหรับสินค้าแต่ละกลุ่มผ่านปุ่ม Config Rules
              โดย match จาก CATEGORY / SUBCATEGORY / DESC_C (ใช้ "ทั้งหมด" เพื่อ match ทุกค่า)
              Rule แรกที่ตรงกันจะถูกใช้ — หากไม่มี Rule ใดตรง ระบบใช้ค่า default <strong>100%</strong>
              และ <strong>คอลัมน์ Q (Net)</strong> จะคำนวณอัตโนมัติจาก O × P%
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
