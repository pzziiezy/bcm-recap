"use client";

import { useRef, useState, useEffect, useLayoutEffect, useMemo, DragEvent } from "react";
import {
  CloudUpload, CheckCircle, XCircle, Clock, RefreshCw, Search,
  ChevronLeft, ChevronRight, ChevronDown, ChevronUp, Database,
  ArrowUpDown, ArrowUp, ArrowDown, Filter, LayoutList, Network, X,
  Eye, Save, Pin,
} from "lucide-react";

const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "";

export interface DriveFileInfo { id: string; name: string; createdTime: string; }

export function formatDateTime(isoString: string): string {
  const d = new Date(isoString);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

export interface SpacemanValues {
  categories: string[];
  subcategories: string[];
  descAList: string[];
  descBList: string[];
  descCList: string[];
  hierarchyMap: { divToDept: Record<string, string[]>; deptToSub: Record<string, string[]>; subToCls: Record<string, string[]> };
  catToSub: Record<string, string[]>; // CATEGORY → SUBCATEGORY (for Config Rules cascade)
}

interface Props {
  onFileInfoChange: (info: DriveFileInfo | null) => void;
  /** True when this tab is visible — needed to measure frozen-column offsets */
  isVisible?: boolean;
  /** Called once after DATA_SPACEMAN finishes parsing — provides unique CATEGORY/SUBCATEGORY/DESC_C values for Config dropdowns */
  onSpacemanValues?: (values: SpacemanValues) => void;
}

type DataRow = Record<string, string>;
type TreeNode = { count: number; children: Map<string, TreeNode> };
type TreeSel = { f02?: string; f03?: string; f04?: string; pg?: string };

const PAGE_SIZE = 100;
const TREE_COLS = ["PLANOFOLDER02", "PLANOFOLDER03", "PLANOFOLDER04", "PLANOGRAM"] as const;

const LEVEL_COLORS = [
  "text-[#be185d] font-semibold text-xs",
  "text-[#0369a1] font-medium text-xs",
  "text-[#c2410c] text-xs",
  "text-[#15803d] text-xs",
];
const LEVEL_INDENT = ["pl-2", "pl-5", "pl-8", "pl-11"];
const LEVEL_DOT = ["bg-[#E91E8C]", "bg-[#00A6E2]", "bg-[#F15A22]", "bg-[#72BF44]"];

export default function SpacemanMaster({ onFileInfoChange, isVisible = false, onSpacemanValues }: Props) {
  // File meta & table data
  const [latestFile, setLatestFile] = useState<DriveFileInfo | null>(null);
  const [loadingMeta, setLoadingMeta] = useState(true);
  const [headers, setHeaders] = useState<string[]>([]);
  const [tableData, setTableData] = useState<DataRow[]>([]);
  const [loadingData, setLoadingData] = useState(false);
  const [parseProgress, setParseProgress] = useState(0);
  const [dataError, setDataError] = useState("");
  const [totalRows, setTotalRows] = useState(0);
  const parseWorkerRef = useRef<Worker | null>(null);

  // Table controls
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const [sortCol, setSortCol] = useState<string | null>(null);
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [showColFilters, setShowColFilters] = useState(false);

  // View mode
  const [viewMode, setViewMode] = useState<"table" | "tree">("table");
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [treeSel, setTreeSel] = useState<TreeSel>({});

  // Column visibility
  const [hiddenCols, setHiddenCols] = useState<Set<string>>(new Set());
  const [showColPicker, setShowColPicker] = useState(false);
  const [colPickerSearch, setColPickerSearch] = useState("");
  const [savedMsg, setSavedMsg] = useState(false);
  const colPickerRef = useRef<HTMLDivElement>(null);

  // Freeze columns
  const [frozenCols, setFrozenCols] = useState<Set<string>>(new Set());
  const [showFreezePicker, setShowFreezePicker] = useState(false);
  const [freezePickerSearch, setFreezePickerSearch] = useState("");
  const [freezeSavedMsg, setFreezeSavedMsg] = useState(false);
  const [frozenLeftOffsets, setFrozenLeftOffsets] = useState<Record<string, number>>({});
  const freezePickerRef = useRef<HTMLDivElement>(null);
  const thRefs = useRef<Map<string, HTMLTableCellElement>>(new Map());

  // Upload
  const [showUpload, setShowUpload] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<"idle" | "success" | "error">("idle");
  const [uploadError, setUploadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [gisReady, setGisReady] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const tokenClientRef = useRef<any>(null);
  const pendingFileRef = useRef<File | null>(null);

  // ── GIS init ──────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!GOOGLE_CLIENT_ID) return;
    const initClient = () => {
      tokenClientRef.current = (window as any).google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: "https://www.googleapis.com/auth/drive",
        callback: async (resp: { access_token?: string; error?: string }) => {
          const file = pendingFileRef.current;
          pendingFileRef.current = null;
          if (resp.error || !resp.access_token || !file) {
            setUploadStatus("error");
            setUploadError(resp.error || "ไม่สามารถรับ access token ได้");
            setUploading(false);
            return;
          }
          try {
            const fd = new FormData();
            fd.append("file", file);
            fd.append("accessToken", resp.access_token);
            const res = await fetch("/api/spaceman/upload", { method: "POST", body: fd });
            if (!res.ok) { const d = await res.json(); throw new Error(d.error || "อัปโหลดล้มเหลว"); }
            setUploadStatus("success");
            setSelectedFile(null);
            const newFile = await fetchLatest();
            if (newFile) loadData(newFile);
          } catch (err) {
            setUploadStatus("error");
            setUploadError(String(err));
          } finally {
            setUploading(false);
          }
        },
      });
      setGisReady(true);
    };
    if ((window as any).google?.accounts?.oauth2) initClient();
    else {
      const script = document.createElement("script");
      script.src = "https://accounts.google.com/gsi/client";
      script.async = true;
      script.onload = initClient;
      document.head.appendChild(script);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Data fetching ─────────────────────────────────────────────────────────
  const fetchLatest = async () => {
    setLoadingMeta(true);
    try {
      const res = await fetch("/api/spaceman/latest");
      const data = await res.json();
      const file: DriveFileInfo | null = data.file ?? null;
      setLatestFile(file);
      onFileInfoChange(file);
      return file;
    } catch {
      setLatestFile(null);
      onFileInfoChange(null);
      return null;
    } finally {
      setLoadingMeta(false);
    }
  };

  const loadData = async (file: DriveFileInfo) => {
    // Terminate any in-progress parse
    parseWorkerRef.current?.terminate();
    parseWorkerRef.current = null;

    setLoadingData(true);
    setParseProgress(0);
    setDataError("");
    setTableData([]);
    setHeaders([]);
    setPage(0);
    setSortCol(null);
    setSortDir("asc");
    setColFilters({});
    setSearch("");
    setTreeSel({});
    setExpanded(new Set());

    try {
      const res = await fetch(`/api/spaceman/file?id=${file.id}`);
      if (!res.ok) throw new Error("ดาวน์โหลดไฟล์ไม่สำเร็จ");
      const buffer = await res.arrayBuffer();

      // Parse in a Web Worker so the main thread stays responsive
      const worker = new Worker(
        new URL("../lib/spaceman.worker.ts", import.meta.url)
      );
      parseWorkerRef.current = worker;

      worker.onmessage = (e: MessageEvent) => {
        const msg = e.data as
          | { type: "progress"; pct: number }
          | { type: "done"; headers: string[]; rows: DataRow[]; totalRows: number; uniqueCategories: string[]; uniqueSubcategories: string[]; uniqueDescA: string[]; uniqueDescB: string[]; uniqueDescC: string[]; hierarchyMap: { divToDept: Record<string, string[]>; deptToSub: Record<string, string[]>; subToCls: Record<string, string[]> }; catToSub: Record<string, string[]> }
          | { type: "error"; message: string };

        if (msg.type === "progress") {
          setParseProgress(msg.pct);
        } else if (msg.type === "done") {
          parseWorkerRef.current = null;
          setHeaders(msg.headers);
          setTableData(msg.rows);
          setTotalRows(msg.totalRows ?? msg.rows.length);
          setParseProgress(100);
          setLoadingData(false);
          if (onSpacemanValues) {
            onSpacemanValues({
              categories:    msg.uniqueCategories    ?? [],
              subcategories: msg.uniqueSubcategories ?? [],
              descAList:     msg.uniqueDescA         ?? [],
              descBList:     msg.uniqueDescB         ?? [],
              descCList:     msg.uniqueDescC         ?? [],
              hierarchyMap:  msg.hierarchyMap        ?? { divToDept: {}, deptToSub: {}, subToCls: {} },
              catToSub:      msg.catToSub            ?? {},
            });
          }
        } else if (msg.type === "error") {
          parseWorkerRef.current = null;
          setDataError(msg.message);
          setLoadingData(false);
        }
      };

      worker.onerror = (e: ErrorEvent) => {
        parseWorkerRef.current = null;
        setDataError(e.message ?? "Parse error in worker");
        setLoadingData(false);
      };

      // Transfer ownership of buffer to worker (zero-copy, frees main-thread memory)
      worker.postMessage({ type: "parse", buffer }, [buffer]);
    } catch (err) {
      setDataError(String(err));
      setLoadingData(false);
    }
  };

  useEffect(() => {
    fetchLatest().then((file) => { if (file) loadData(file); });
    return () => { parseWorkerRef.current?.terminate(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Load saved column prefs once headers are known
  useEffect(() => {
    if (headers.length === 0) return;
    try {
      const savedHidden = localStorage.getItem("spaceman_hidden_cols");
      if (savedHidden) {
        const arr: string[] = JSON.parse(savedHidden);
        setHiddenCols(new Set(arr.filter((h) => headers.includes(h))));
      }
      const savedFrozen = localStorage.getItem("spaceman_frozen_cols");
      if (savedFrozen) {
        const arr: string[] = JSON.parse(savedFrozen);
        setFrozenCols(new Set(arr.filter((h) => headers.includes(h))));
      }
    } catch {}
  }, [headers]);

  // Close picker on outside click
  useEffect(() => {
    if (!showColPicker) return;
    const handler = (e: MouseEvent) => {
      if (colPickerRef.current && !colPickerRef.current.contains(e.target as Node))
        setShowColPicker(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [showColPicker]);

  useEffect(() => {
    if (!showFreezePicker) return;
    const handler = (e: MouseEvent) => {
      if (freezePickerRef.current && !freezePickerRef.current.contains(e.target as Node))
        setShowFreezePicker(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [showFreezePicker]);

  // ── Upload handlers ───────────────────────────────────────────────────────
  const handleFiles = (files: FileList | null) => {
    if (!files || files.length === 0) return;
    setSelectedFile(files[0]);
    setUploadStatus("idle");
    setUploadError("");
  };

  const handleUpload = () => {
    if (!selectedFile) return;
    if (!gisReady || !tokenClientRef.current) {
      setUploadStatus("error");
      setUploadError("Google Identity Services ยังไม่พร้อม กรุณารีเฟรชหน้าแล้วลองใหม่");
      return;
    }
    setUploading(true);
    setUploadStatus("idle");
    setUploadError("");
    pendingFileRef.current = selectedFile;
    tokenClientRef.current.requestAccessToken({ prompt: "" });
  };

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragging(false);
    if (!uploading) handleFiles(e.dataTransfer.files);
  };

  // ── Tree data ─────────────────────────────────────────────────────────────
  const treeData = useMemo((): Map<string, TreeNode> => {
    const root = new Map<string, TreeNode>();
    for (const row of tableData) {
      const vals = TREE_COLS.map((c) => row[c] || "(ไม่ระบุ)");
      let cur = root;
      for (const val of vals) {
        if (!cur.has(val)) cur.set(val, { count: 0, children: new Map() });
        const node = cur.get(val)!;
        node.count++;
        cur = node.children;
      }
    }
    return root;
  }, [tableData]);

  // ── Filter + sort ─────────────────────────────────────────────────────────
  const filtered = useMemo(() => {
    let data = tableData;
    const activeFilters = Object.entries(colFilters).filter(([, v]) => v.trim());
    if (activeFilters.length > 0)
      data = data.filter((row) => activeFilters.every(([col, val]) => (row[col] || "").toLowerCase().includes(val.toLowerCase())));
    if (search.trim()) {
      const q = search.toLowerCase();
      data = data.filter((row) => Object.values(row).some((v) => v.toLowerCase().includes(q)));
    }
    if (sortCol) {
      data = [...data].sort((a, b) => {
        const av = a[sortCol] || "", bv = b[sortCol] || "";
        const aNum = Number(av), bNum = Number(bv);
        const isNum = av !== "" && bv !== "" && !isNaN(aNum) && !isNaN(bNum);
        const cmp = isNum ? aNum - bNum : av.localeCompare(bv, "th");
        return sortDir === "asc" ? cmp : -cmp;
      });
    }
    return data;
  }, [tableData, search, colFilters, sortCol, sortDir]);

  // Apply tree selection on top of existing filters
  const displayData = useMemo(() => {
    if (!treeSel.f02 && !treeSel.f03 && !treeSel.f04 && !treeSel.pg) return filtered;
    return filtered.filter((row) => {
      if (treeSel.f02 && row["PLANOFOLDER02"] !== treeSel.f02) return false;
      if (treeSel.f03 && row["PLANOFOLDER03"] !== treeSel.f03) return false;
      if (treeSel.f04 && row["PLANOFOLDER04"] !== treeSel.f04) return false;
      if (treeSel.pg  && row["PLANOGRAM"]     !== treeSel.pg)  return false;
      return true;
    });
  }, [filtered, treeSel]);

  const totalPages = Math.ceil(displayData.length / PAGE_SIZE);
  const pageData = displayData.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  // ── Table handlers ────────────────────────────────────────────────────────
  const handleSearch = (val: string) => { setSearch(val); setPage(0); };
  const handleSort = (col: string) => {
    if (sortCol === col) { if (sortDir === "asc") setSortDir("desc"); else { setSortCol(null); setSortDir("asc"); } }
    else { setSortCol(col); setSortDir("asc"); }
    setPage(0);
  };
  const handleColFilter = (col: string, val: string) => { setColFilters((p) => ({ ...p, [col]: val })); setPage(0); };
  const clearAllFilters = () => { setColFilters({}); setSearch(""); setSortCol(null); setSortDir("asc"); setPage(0); };
  const activeFilterCount = Object.values(colFilters).filter((v) => v.trim()).length + (search.trim() ? 1 : 0);

  const visibleHeaders = useMemo(() => headers.filter((h) => !hiddenCols.has(h)), [headers, hiddenCols]);

  const frozenHeaders = useMemo(
    () => visibleHeaders.filter((h) => frozenCols.has(h)),
    [visibleHeaders, frozenCols]
  );
  const scrollableHeaders = useMemo(
    () => visibleHeaders.filter((h) => !frozenCols.has(h)),
    [visibleHeaders, frozenCols]
  );
  const allDisplayHeaders = useMemo(
    () => [...frozenHeaders, ...scrollableHeaders],
    [frozenHeaders, scrollableHeaders]
  );

  const toggleCol = (col: string) =>
    setHiddenCols((prev) => { const n = new Set(prev); n.has(col) ? n.delete(col) : n.add(col); return n; });

  const saveColPrefs = () => {
    localStorage.setItem("spaceman_hidden_cols", JSON.stringify([...hiddenCols]));
    setSavedMsg(true);
    setTimeout(() => setSavedMsg(false), 2000);
  };

  const toggleFrozenCol = (col: string) =>
    setFrozenCols((prev) => { const n = new Set(prev); n.has(col) ? n.delete(col) : n.add(col); return n; });

  const unfreezeCol = (col: string) =>
    setFrozenCols((prev) => { const n = new Set(prev); n.delete(col); return n; });

  const saveFreezePrefs = () => {
    localStorage.setItem("spaceman_frozen_cols", JSON.stringify([...frozenCols]));
    setFreezeSavedMsg(true);
    setTimeout(() => setFreezeSavedMsg(false), 2000);
  };

  // ── Tree handlers ─────────────────────────────────────────────────────────
  const handleTreeSelect = (level: number, pathParts: string[], label: string) => {
    const parts = [...pathParts, label];
    const newSel: TreeSel = {};
    if (parts[0]) newSel.f02 = parts[0];
    if (level >= 1 && parts[1]) newSel.f03 = parts[1];
    if (level >= 2 && parts[2]) newSel.f04 = parts[2];
    if (level >= 3 && parts[3]) newSel.pg = parts[3];
    setTreeSel(newSel);
    setPage(0);
  };

  const clearTreeSel = () => { setTreeSel({}); setPage(0); };

  // Measure frozen column widths to compute cumulative sticky left offsets.
  // Must depend on isVisible: offsetWidth returns 0 when display:none, so we
  // re-measure the moment the tab becomes visible.
  useLayoutEffect(() => {
    if (!isVisible || frozenHeaders.length === 0) {
      setFrozenLeftOffsets({});
      return;
    }
    const offsets: Record<string, number> = {};
    let cumLeft = 0;
    for (const h of frozenHeaders) {
      offsets[h] = cumLeft;
      cumLeft += thRefs.current.get(h)?.offsetWidth ?? 120;
    }
    setFrozenLeftOffsets(offsets);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [frozenHeaders, page, isVisible]);

  const toggleExpand = (pathKey: string) => {
    setExpanded((prev) => {
      const next = new Set(prev);
      next.has(pathKey) ? next.delete(pathKey) : next.add(pathKey);
      return next;
    });
  };

  const isOnPath = (level: number, pathParts: string[], label: string) => {
    const parts = [...pathParts, label];
    const vals: (string | undefined)[] = [treeSel.f02, treeSel.f03, treeSel.f04, treeSel.pg];
    for (let i = 0; i <= level; i++) if (vals[i] !== parts[i]) return false;
    return true;
  };

  const isLeafSel = (level: number, pathParts: string[], label: string) => {
    if (!isOnPath(level, pathParts, label)) return false;
    const vals = [treeSel.f02, treeSel.f03, treeSel.f04, treeSel.pg];
    for (let i = level + 1; i < 4; i++) if (vals[i]) return false;
    return true;
  };

  const breadcrumb = [treeSel.f02, treeSel.f03, treeSel.f04, treeSel.pg].filter(Boolean) as string[];

  // ── Tree renderer ─────────────────────────────────────────────────────────
  const renderTree = (map: Map<string, TreeNode>, level: number, pathParts: string[]): React.ReactNode =>
    [...map.entries()]
      .sort(([a], [b]) => a.localeCompare(b, "th"))
      .map(([label, node]) => {
        const pathKey = [...pathParts, label].join("\0");
        const isExp = expanded.has(pathKey);
        const hasKids = node.children.size > 0 && level < 3;
        const onPath = isOnPath(level, pathParts, label);
        const isLeaf = isLeafSel(level, pathParts, label);

        return (
          <div key={pathKey}>
            <div
              onClick={() => handleTreeSelect(level, pathParts, label)}
              className={`flex items-center gap-1 py-1 pr-2 cursor-pointer rounded-lg transition-all select-none
                ${LEVEL_INDENT[level]}
                ${isLeaf ? "bg-pink-100" : onPath ? "bg-pink-50" : "hover:bg-slate-100"}`}
            >
              {hasKids ? (
                <button
                  onClick={(e) => { e.stopPropagation(); toggleExpand(pathKey); }}
                  className="flex-shrink-0 w-4 h-4 flex items-center justify-center text-slate-400 hover:text-slate-600"
                >
                  {isExp ? <ChevronDown className="w-3 h-3" /> : <ChevronRight className="w-3 h-3" />}
                </button>
              ) : (
                <span className="flex-shrink-0 w-4 h-4 flex items-center justify-center">
                  <span className={`w-1.5 h-1.5 rounded-full ${LEVEL_DOT[level]}`} />
                </span>
              )}
              <span className={`flex-1 truncate ${isLeaf ? "text-[#E91E8C] font-semibold text-xs" : LEVEL_COLORS[level]}`}>
                {label}
              </span>
              <span className={`text-[10px] px-1.5 py-0.5 rounded-full flex-shrink-0 tabular-nums
                ${isLeaf ? "bg-[#E91E8C] text-white" : "bg-slate-100 text-slate-500"}`}>
                {node.count.toLocaleString()}
              </span>
            </div>
            {hasKids && isExp && renderTree(node.children, level + 1, [...pathParts, label])}
          </div>
        );
      });

  // ── Sticky style helper for frozen columns ────────────────────────────────
  const getStickyStyle = (h: string, zIndex: number): React.CSSProperties | undefined => {
    if (!frozenCols.has(h)) return undefined;
    const isLastFrozen = h === frozenHeaders[frozenHeaders.length - 1];
    return {
      position: "sticky",
      left: frozenLeftOffsets[h] ?? 0,
      zIndex,
      ...(isLastFrozen ? { boxShadow: "2px 0 6px -2px rgba(0,0,0,0.10)" } : {}),
    };
  };

  // ── Shared table content ──────────────────────────────────────────────────
  const tableContent = (
    <>
      <table className="w-full text-sm">
        <thead>
          <tr className="bg-slate-50 border-b border-slate-200">
            {allDisplayHeaders.map((h) => {
              const isFrozen = frozenCols.has(h);
              return (
                <th
                  key={h}
                  ref={(el) => { if (el) thRefs.current.set(h, el); else thRefs.current.delete(h); }}
                  onClick={() => handleSort(h)}
                  style={getStickyStyle(h, 20)}
                  className={`px-4 py-2.5 text-left text-xs font-semibold text-slate-600 whitespace-nowrap cursor-pointer select-none transition-colors group
                    ${isFrozen ? "bg-pink-50 hover:bg-pink-100" : "bg-slate-50 hover:bg-slate-100"}`}
                >
                  <div className="flex items-center gap-1">
                    {isFrozen && (
                      <button
                        onClick={(e) => { e.stopPropagation(); unfreezeCol(h); }}
                        className="text-[#E91E8C] hover:text-red-500 transition-colors flex-shrink-0"
                        title="ยกเลิก Freeze"
                      >
                        <Pin className="w-3 h-3" />
                      </button>
                    )}
                    {h}
                    {sortCol === h
                      ? sortDir === "asc" ? <ArrowUp className="w-3 h-3 text-[#E91E8C]" /> : <ArrowDown className="w-3 h-3 text-[#E91E8C]" />
                      : <ArrowUpDown className="w-3 h-3 text-slate-300 group-hover:text-slate-400" />}
                  </div>
                </th>
              );
            })}
          </tr>
          {showColFilters && (
            <tr className="bg-white border-b border-slate-200">
              {allDisplayHeaders.map((h) => {
                const isFrozen = frozenCols.has(h);
                return (
                  <th key={h}
                    style={isFrozen ? getStickyStyle(h, 10) : undefined}
                    className={`px-2 py-1.5 ${isFrozen ? "bg-white" : ""}`}
                  >
                    <input
                      type="text"
                      value={colFilters[h] || ""}
                      onChange={(e) => handleColFilter(h, e.target.value)}
                      placeholder="filter..."
                      className="w-full px-2 py-1 text-xs border border-slate-200 rounded focus:outline-none focus:ring-1 focus:ring-pink-300 focus:border-[#E91E8C] font-normal min-w-[80px]"
                    />
                  </th>
                );
              })}
            </tr>
          )}
        </thead>
        <tbody className="divide-y divide-slate-100">
          {pageData.map((row, i) => (
            <tr key={i} className="group hover:bg-pink-50/40 transition-colors">
              {allDisplayHeaders.map((h) => {
                const isFrozen = frozenCols.has(h);
                return (
                  <td key={h}
                    style={isFrozen ? getStickyStyle(h, 1) : undefined}
                    className={`px-4 py-2 text-slate-700 whitespace-nowrap max-w-xs truncate ${isFrozen ? "bg-white group-hover:bg-pink-50" : ""}`}
                  >
                    {search.trim() ? highlightMatch(row[h] || "", search) : (row[h] || "")}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>

      {totalPages > 1 && (
        <div className="px-6 py-3 border-t border-slate-100 flex items-center justify-between text-sm text-slate-500 bg-white sticky bottom-0">
          <span>
            แสดง {page * PAGE_SIZE + 1}–{Math.min((page + 1) * PAGE_SIZE, displayData.length)} จาก{" "}
            {displayData.length.toLocaleString()} แถว
            {(search || Object.values(colFilters).some(Boolean)) && ` (กรองจาก ${tableData.length.toLocaleString()})`}
          </span>
          <div className="flex items-center gap-1">
            <button onClick={() => setPage((p) => Math.max(0, p - 1))} disabled={page === 0}
              className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed">
              <ChevronLeft className="w-4 h-4" />
            </button>
            <span className="px-3 font-medium">{page + 1} / {totalPages}</span>
            <button onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))} disabled={page >= totalPages - 1}
              className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed">
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}
    </>
  );

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className="space-y-6">
      {/* ── Upload card ── */}
      <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#00A6E2] via-[#FFD100] via-[#F15A22] to-[#72BF44]" />
        <div className="px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-1 h-5 rounded-full bg-gradient-to-b from-[#E91E8C] to-[#F15A22]" />
            <div>
              <h2 className="font-bold text-slate-800 text-lg">DATA_SPACEMAN</h2>
              {loadingMeta ? (
                <p className="text-xs text-slate-400 mt-0.5 flex items-center gap-1">
                  <span className="inline-block animate-spin rounded-full h-3 w-3 border border-slate-300 border-t-slate-500" />
                  กำลังตรวจสอบ...
                </p>
              ) : latestFile ? (
                <p className="text-xs text-slate-500 mt-0.5 flex items-center gap-1">
                  <Clock className="w-3 h-3" />
                  อัปโหลดล่าสุด:{" "}
                  <strong className="text-slate-700">{formatDateTime(latestFile.createdTime)}</strong>
                  <span className="text-slate-400 ml-1">— {latestFile.name}</span>
                </p>
              ) : (
                <p className="text-xs text-amber-600 mt-0.5">ยังไม่มีไฟล์ใน Google Drive</p>
              )}
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => fetchLatest().then((f) => f && loadData(f))}
              disabled={loadingMeta || loadingData}
              className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-[#E91E8C] transition-colors disabled:opacity-40 px-3 py-1.5 rounded-lg hover:bg-pink-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${loadingMeta || loadingData ? "animate-spin" : ""}`} />
              รีเฟรช
            </button>
            <button
              onClick={() => { setShowUpload((v) => !v); setUploadStatus("idle"); }}
              className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white hover:from-[#d41679] hover:to-[#be185d] transition-all shadow-sm"
            >
              <CloudUpload className="w-3.5 h-3.5" />
              อัปโหลดไฟล์ใหม่
              {showUpload ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
            </button>
          </div>
        </div>

        {showUpload && (
          <div className="px-6 pb-6 space-y-3 border-t border-pink-50 pt-4">
            <div
              onClick={() => !uploading && inputRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              className={`flex flex-col items-center justify-center gap-2 border-2 border-dashed rounded-xl p-8 transition-all duration-200
                ${uploading ? "opacity-60 cursor-not-allowed border-pink-200 bg-pink-50/30"
                  : dragging ? "border-[#E91E8C] bg-pink-50 scale-[1.01] cursor-pointer"
                  : selectedFile ? "border-green-300 bg-green-50/40 cursor-pointer"
                  : "border-pink-200 bg-pink-50/30 hover:border-[#E91E8C] hover:bg-pink-50 cursor-pointer"}`}
            >
              <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden"
                onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }} />
              {uploading ? (
                <div className="animate-spin rounded-full h-8 w-8 border-4 border-pink-200 border-t-[#E91E8C]" />
              ) : selectedFile ? (
                <CheckCircle className="w-8 h-8 text-green-500" />
              ) : (
                <CloudUpload className={`w-8 h-8 ${dragging ? "text-[#E91E8C]" : "text-pink-300"}`} />
              )}
              <div className="text-center">
                {uploading ? <p className="text-sm text-slate-500">กำลังอัปโหลดไปยัง Google Drive...</p>
                  : selectedFile ? (
                    <><p className="font-semibold text-green-700 text-sm">{selectedFile.name}</p>
                    <p className="text-xs text-slate-400 mt-0.5">คลิกเพื่อเลือกไฟล์ใหม่</p></>
                  ) : (
                    <><p className="font-semibold text-slate-700 text-sm">เลือกไฟล์ DATA_SPACEMAN</p>
                    <p className="text-xs text-slate-400 mt-0.5">คลิกหรือลากไฟล์ .xlsx มาวางที่นี่</p></>
                  )}
              </div>
            </div>
            {selectedFile && !uploading && uploadStatus !== "success" && (
              <button onClick={handleUpload}
                className="w-full flex items-center justify-center gap-2 py-3 rounded-xl font-semibold text-sm bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white hover:from-[#d41679] hover:to-[#be185d] transition-all shadow-sm">
                <CloudUpload className="w-4 h-4" />อัปเดตเป็นไฟล์ล่าสุด
              </button>
            )}
            {uploadStatus === "success" && (
              <div className="flex items-center gap-2 px-4 py-2.5 bg-green-50 border border-green-200 rounded-xl text-green-700 text-sm">
                <CheckCircle className="w-4 h-4 flex-shrink-0" />อัปโหลดสำเร็จ! ข้อมูลด้านล่างได้รับการอัปเดตแล้ว
              </div>
            )}
            {uploadStatus === "error" && (
              <div className="flex items-start gap-2 px-4 py-2.5 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
                <XCircle className="w-4 h-4 flex-shrink-0 mt-0.5" /><span>อัปโหลดล้มเหลว: {uploadError}</span>
              </div>
            )}
          </div>
        )}
      </div>

      {/* ── Data card ── */}
      <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#00A6E2] via-[#FFD100] via-[#F15A22] to-[#72BF44]" />

        {/* Card header */}
        <div className="px-6 py-4 border-b border-pink-50 flex items-center justify-between gap-4 flex-wrap">
          <div className="flex items-center gap-3">
            <div className="w-1 h-5 rounded-full bg-gradient-to-b from-[#E91E8C] to-[#F15A22]" />
            <div className="flex items-center gap-2 flex-wrap">
              <Database className="w-4 h-4 text-slate-400" />
              <h2 className="font-bold text-slate-800 text-lg">ข้อมูลใน QRY_Product_by_POG</h2>
              {tableData.length > 0 && (
                <span className="text-xs bg-pink-100 text-[#E91E8C] px-2 py-0.5 rounded-full font-medium">
                  {totalRows.toLocaleString()} แถว
                </span>
              )}
              {totalRows > tableData.length && (
                <span className="text-xs bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full font-medium">
                  แสดง {tableData.length.toLocaleString()} แถวแรก
                </span>
              )}
              {displayData.length !== tableData.length && (
                <span className="text-xs bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full font-medium">
                  กรองแล้ว: {displayData.length.toLocaleString()} แถว
                </span>
              )}
            </div>
          </div>

          {tableData.length > 0 && (
            <div className="flex items-center gap-2 flex-wrap">
              {/* View toggle */}
              <div className="flex bg-slate-100 rounded-lg p-0.5 gap-0.5">
                <button
                  onClick={() => setViewMode("table")}
                  className={`flex items-center gap-1.5 text-xs px-2.5 py-1.5 rounded-md transition-all ${
                    viewMode === "table" ? "bg-white text-slate-700 shadow-sm font-medium" : "text-slate-500 hover:text-slate-700"
                  }`}
                >
                  <LayoutList className="w-3.5 h-3.5" /> ตาราง
                </button>
                <button
                  onClick={() => setViewMode("tree")}
                  className={`flex items-center gap-1.5 text-xs px-2.5 py-1.5 rounded-md transition-all ${
                    viewMode === "tree" ? "bg-white text-slate-700 shadow-sm font-medium" : "text-slate-500 hover:text-slate-700"
                  }`}
                >
                  <Network className="w-3.5 h-3.5" /> Tree
                </button>
              </div>

              {/* Global search */}
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="ค้นหาทุกคอลัมน์..."
                  value={search}
                  onChange={(e) => handleSearch(e.target.value)}
                  className="pl-8 pr-3 py-1.5 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-pink-200 focus:border-[#E91E8C] w-48"
                />
              </div>

              {/* Column filter toggle */}
              <button
                onClick={() => setShowColFilters((v) => !v)}
                className={`flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border transition-all ${
                  showColFilters || activeFilterCount > 0
                    ? "bg-pink-50 border-[#E91E8C] text-[#E91E8C]"
                    : "border-slate-200 text-slate-500 hover:border-slate-300"
                }`}
              >
                <Filter className="w-3.5 h-3.5" />
                Filter
                {activeFilterCount > 0 && (
                  <span className="bg-[#E91E8C] text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold">
                    {activeFilterCount}
                  </span>
                )}
              </button>
              {(activeFilterCount > 0 || treeSel.f02) && (
                <button
                  onClick={() => { clearAllFilters(); clearTreeSel(); }}
                  className="text-xs text-slate-400 hover:text-red-500 transition-colors"
                >
                  ล้างทั้งหมด
                </button>
              )}

              {/* Column visibility picker */}
              <div className="relative" ref={colPickerRef}>
                <button
                  onClick={() => setShowColPicker((v) => !v)}
                  className={`flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border transition-all ${
                    showColPicker || hiddenCols.size > 0
                      ? "bg-pink-50 border-[#E91E8C] text-[#E91E8C]"
                      : "border-slate-200 text-slate-500 hover:border-slate-300"
                  }`}
                >
                  <Eye className="w-3.5 h-3.5" />
                  คอลัมน์
                  {hiddenCols.size > 0 && (
                    <span className="bg-[#E91E8C] text-white rounded-full px-1.5 text-[10px] font-bold">
                      -{hiddenCols.size}
                    </span>
                  )}
                </button>

                {showColPicker && (
                  <div className="absolute right-0 top-full mt-1 z-50 bg-white rounded-xl border border-slate-200 shadow-xl w-72 overflow-hidden">
                    {/* Picker header */}
                    <div className="px-3 py-2.5 border-b border-slate-100 flex items-center justify-between">
                      <span className="text-xs font-semibold text-slate-700">แสดง / ซ่อนคอลัมน์</span>
                      <div className="flex items-center gap-2">
                        <button onClick={() => setHiddenCols(new Set())} className="text-[10px] text-blue-500 hover:text-blue-700">แสดงทั้งหมด</button>
                        <span className="text-slate-300">|</span>
                        <button onClick={() => setHiddenCols(new Set(headers))} className="text-[10px] text-slate-500 hover:text-slate-700">ซ่อนทั้งหมด</button>
                      </div>
                    </div>

                    {/* Picker search */}
                    <div className="px-3 py-2 border-b border-slate-100">
                      <div className="relative">
                        <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3 h-3 text-slate-400" />
                        <input
                          type="text"
                          placeholder="ค้นหาชื่อคอลัมน์..."
                          value={colPickerSearch}
                          onChange={(e) => setColPickerSearch(e.target.value)}
                          className="w-full pl-7 pr-2 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-pink-300 focus:border-[#E91E8C]"
                        />
                      </div>
                    </div>

                    {/* Column list */}
                    <div className="max-h-64 overflow-y-auto p-2 space-y-0.5">
                      {headers
                        .filter((h) => !colPickerSearch || h.toLowerCase().includes(colPickerSearch.toLowerCase()))
                        .map((h) => (
                          <label key={h} className="flex items-center gap-2.5 px-2 py-1.5 rounded-lg hover:bg-slate-50 cursor-pointer">
                            <input
                              type="checkbox"
                              checked={!hiddenCols.has(h)}
                              onChange={() => toggleCol(h)}
                              className="accent-[#E91E8C] w-3.5 h-3.5 flex-shrink-0"
                            />
                            <span className={`text-xs flex-1 truncate ${hiddenCols.has(h) ? "text-slate-400 line-through" : "text-slate-700"}`}>
                              {h}
                            </span>
                          </label>
                        ))}
                    </div>

                    {/* Picker footer */}
                    <div className="px-3 py-2.5 border-t border-slate-100 flex items-center justify-between bg-slate-50/50">
                      <span className="text-[10px] text-slate-400">
                        แสดง {headers.length - hiddenCols.size} / {headers.length} คอลัมน์
                      </span>
                      <button
                        onClick={saveColPrefs}
                        className={`flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg transition-all font-medium ${
                          savedMsg
                            ? "bg-green-100 text-green-700 border border-green-200"
                            : "bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white hover:from-[#d41679] hover:to-[#be185d] shadow-sm"
                        }`}
                      >
                        {savedMsg ? <CheckCircle className="w-3 h-3" /> : <Save className="w-3 h-3" />}
                        {savedMsg ? "บันทึกแล้ว!" : "บันทึกค่าเริ่มต้น"}
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* Freeze columns picker */}
              <div className="relative" ref={freezePickerRef}>
                <button
                  onClick={() => setShowFreezePicker((v) => !v)}
                  className={`flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border transition-all ${
                    showFreezePicker || frozenCols.size > 0
                      ? "bg-pink-50 border-[#E91E8C] text-[#E91E8C]"
                      : "border-slate-200 text-slate-500 hover:border-slate-300"
                  }`}
                >
                  <Pin className="w-3.5 h-3.5" />
                  Freeze
                  {frozenCols.size > 0 && (
                    <span className="bg-[#E91E8C] text-white rounded-full px-1.5 text-[10px] font-bold">
                      {frozenCols.size}
                    </span>
                  )}
                </button>

                {showFreezePicker && (
                  <div className="absolute right-0 top-full mt-1 z-50 bg-white rounded-xl border border-slate-200 shadow-xl w-72 overflow-hidden">
                    {/* Picker header */}
                    <div className="px-3 py-2.5 border-b border-slate-100 flex items-center justify-between">
                      <span className="text-xs font-semibold text-slate-700">Freeze คอลัมน์</span>
                      <button onClick={() => setFrozenCols(new Set())} className="text-[10px] text-slate-500 hover:text-red-500">ล้าง Freeze ทั้งหมด</button>
                    </div>

                    {/* Picker search */}
                    <div className="px-3 py-2 border-b border-slate-100">
                      <div className="relative">
                        <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3 h-3 text-slate-400" />
                        <input
                          type="text"
                          placeholder="ค้นหาชื่อคอลัมน์..."
                          value={freezePickerSearch}
                          onChange={(e) => setFreezePickerSearch(e.target.value)}
                          className="w-full pl-7 pr-2 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-pink-300 focus:border-[#E91E8C]"
                        />
                      </div>
                    </div>

                    {/* Column list */}
                    <div className="max-h-64 overflow-y-auto p-2 space-y-0.5">
                      {visibleHeaders
                        .filter((h) => !freezePickerSearch || h.toLowerCase().includes(freezePickerSearch.toLowerCase()))
                        .map((h) => (
                          <label key={h} className="flex items-center gap-2.5 px-2 py-1.5 rounded-lg hover:bg-slate-50 cursor-pointer">
                            <input
                              type="checkbox"
                              checked={frozenCols.has(h)}
                              onChange={() => toggleFrozenCol(h)}
                              className="accent-[#E91E8C] w-3.5 h-3.5 flex-shrink-0"
                            />
                            <span className={`text-xs flex-1 truncate ${frozenCols.has(h) ? "text-[#E91E8C] font-medium" : "text-slate-700"}`}>
                              {h}
                            </span>
                            {frozenCols.has(h) && <Pin className="w-3 h-3 text-[#E91E8C] flex-shrink-0" />}
                          </label>
                        ))}
                    </div>

                    {/* Picker footer */}
                    <div className="px-3 py-2.5 border-t border-slate-100 flex items-center justify-between bg-slate-50/50">
                      <span className="text-[10px] text-slate-400">
                        {frozenCols.size > 0 ? `${frozenCols.size} คอลัมน์ถูก Freeze` : "ยังไม่มีคอลัมน์ถูก Freeze"}
                      </span>
                      <button
                        onClick={saveFreezePrefs}
                        className={`flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg transition-all font-medium ${
                          freezeSavedMsg
                            ? "bg-green-100 text-green-700 border border-green-200"
                            : "bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white hover:from-[#d41679] hover:to-[#be185d] shadow-sm"
                        }`}
                      >
                        {freezeSavedMsg ? <CheckCircle className="w-3 h-3" /> : <Save className="w-3 h-3" />}
                        {freezeSavedMsg ? "บันทึกแล้ว!" : "บันทึกค่าเริ่มต้น"}
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Loading / error / empty states */}
        {loadingData && (
          <div className="flex flex-col items-center justify-center py-16 gap-4 text-slate-500">
            <div className="animate-spin rounded-full h-8 w-8 border-4 border-pink-200 border-t-[#E91E8C]" />
            <p className="text-sm font-medium">
              {parseProgress < 20
                ? "กำลังดาวน์โหลดไฟล์..."
                : `กำลังประมวลผลข้อมูล... (${parseProgress}%)`}
            </p>
            <div className="w-64 bg-slate-200 rounded-full h-2 overflow-hidden">
              <div
                className="h-2 rounded-full transition-all duration-300 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]"
                style={{ width: `${parseProgress}%` }}
              />
            </div>
          </div>
        )}
        {!loadingData && dataError && (
          <div className="m-6 flex items-start gap-2 px-4 py-3 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
            <XCircle className="w-4 h-4 flex-shrink-0 mt-0.5" /><span>{dataError}</span>
          </div>
        )}
        {!loadingData && !dataError && !latestFile && (
          <div className="flex flex-col items-center justify-center py-16 gap-2 text-slate-400">
            <CloudUpload className="w-10 h-10 text-pink-200" />
            <p className="text-sm">ยังไม่มีไฟล์ใน Google Drive</p>
            <p className="text-xs">คลิก "อัปโหลดไฟล์ใหม่" ด้านบนเพื่อเริ่มต้น</p>
          </div>
        )}
        {!loadingData && !dataError && latestFile && tableData.length === 0 && (
          <div className="flex items-center justify-center py-16 text-slate-400 text-sm">ไม่พบข้อมูล</div>
        )}

        {/* ── Tree view ── */}
        {!loadingData && !dataError && tableData.length > 0 && viewMode === "tree" && (
          <div className="flex" style={{ minHeight: "70vh" }}>
            {/* Left: Tree panel */}
            <div className="w-64 flex-shrink-0 border-r border-slate-100 overflow-y-auto bg-slate-50/40 p-2">
              {/* All rows option */}
              <div
                onClick={clearTreeSel}
                className={`flex items-center gap-2 py-1.5 px-2 rounded-lg cursor-pointer text-xs mb-1 transition-all ${
                  !treeSel.f02 ? "bg-pink-100 text-[#E91E8C] font-semibold" : "text-slate-600 hover:bg-slate-100"
                }`}
              >
                <Database className="w-3 h-3 flex-shrink-0" />
                <span className="flex-1">ทั้งหมด</span>
                <span className={`text-[10px] px-1.5 py-0.5 rounded-full tabular-nums ${
                  !treeSel.f02 ? "bg-[#E91E8C] text-white" : "bg-slate-200 text-slate-600"
                }`}>
                  {filtered.length.toLocaleString()}
                </span>
              </div>
              <div className="h-px bg-slate-200 mx-1 mb-1" />

              {/* Level labels */}
              <div className="flex gap-1 mb-2 px-1">
                {(["F02", "F03", "F04", "POG"] as const).map((lbl, i) => (
                  <span key={lbl} className="text-[9px] px-1.5 py-0.5 rounded font-medium"
                    style={{ background: ["#fce7f3","#e0f2fe","#ffedd5","#dcfce7"][i], color: ["#be185d","#0369a1","#c2410c","#15803d"][i] }}>
                    {lbl}
                  </span>
                ))}
              </div>

              {renderTree(treeData, 0, [])}
            </div>

            {/* Right: breadcrumb + table */}
            <div className="flex-1 flex flex-col overflow-hidden min-w-0">
              {/* Breadcrumb */}
              <div className={`flex items-center gap-1.5 px-4 py-2 border-b border-pink-50 text-xs flex-wrap ${
                breadcrumb.length > 0 ? "bg-pink-50/60" : "bg-slate-50/40"
              }`}>
                {breadcrumb.length === 0 ? (
                  <span className="text-slate-400">คลิก node ใน Tree ทางซ้ายเพื่อกรองข้อมูล</span>
                ) : (
                  <>
                    <span className="text-slate-400">เลือก:</span>
                    {breadcrumb.map((b, i) => (
                      <span key={i} className="flex items-center gap-1">
                        {i > 0 && <ChevronRight className="w-3 h-3 text-slate-300" />}
                        <span className={i === breadcrumb.length - 1 ? "text-[#E91E8C] font-semibold" : "text-slate-600"}>
                          {b}
                        </span>
                      </span>
                    ))}
                    <button onClick={clearTreeSel} className="ml-1 text-slate-400 hover:text-red-400 transition-colors">
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </>
                )}
              </div>

              {/* Table (scrollable) */}
              <div className="flex-1 overflow-auto">
                <div className="overflow-x-auto">
                  {tableContent}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ── Table-only view ── */}
        {!loadingData && !dataError && tableData.length > 0 && viewMode === "table" && (
          <div className="overflow-x-auto">
            {tableContent}
          </div>
        )}
      </div>
    </div>
  );
}

function highlightMatch(text: string, query: string): React.ReactNode {
  if (!query.trim()) return text;
  const idx = text.toLowerCase().indexOf(query.toLowerCase());
  if (idx === -1) return text;
  return (
    <>
      {text.slice(0, idx)}
      <mark className="bg-yellow-200 text-yellow-900 rounded px-0.5">{text.slice(idx, idx + query.length)}</mark>
      {text.slice(idx + query.length)}
    </>
  );
}
