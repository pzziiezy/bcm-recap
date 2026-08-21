"use client";

import { useRef, useState, useEffect, useMemo, DragEvent } from "react";
import * as XLSX from "xlsx";
import {
  CloudUpload, CheckCircle, XCircle, Clock, RefreshCw, Search, Database,
  ChevronLeft, ChevronRight, ChevronDown, ChevronUp, X,
  ArrowUpDown, ArrowUp, ArrowDown, Filter,
} from "lucide-react";

const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "";

interface DriveFileInfo { id: string; name: string; createdTime: string; }

type DataRow = Record<string, string>;

const PAGE_SIZE = 100;

function formatDateTime(isoString: string): string {
  const d = new Date(isoString);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

export default function MasterTab() {
  // File meta
  const [latestFile, setLatestFile] = useState<DriveFileInfo | null>(null);
  const [loadingMeta, setLoadingMeta] = useState(true);

  // Table data
  const [headers, setHeaders] = useState<string[]>([]);
  const [tableData, setTableData] = useState<DataRow[]>([]);
  const [sheetUsed, setSheetUsed] = useState("");
  const [loadingData, setLoadingData] = useState(false);
  const [dataError, setDataError] = useState("");
  const [parseProgress, setParseProgress] = useState(0);

  // Table controls
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const [sortCol, setSortCol] = useState<string | null>(null);
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [showColFilters, setShowColFilters] = useState(false);

  // Upload panel
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
            const res = await fetch("/api/master/upload", { method: "POST", body: fd });
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
      const res = await fetch("/api/master/latest");
      const data = await res.json();
      const file: DriveFileInfo | null = data.file ?? null;
      setLatestFile(file);
      return file;
    } catch {
      setLatestFile(null);
      return null;
    } finally {
      setLoadingMeta(false);
    }
  };

  const loadData = async (file: DriveFileInfo) => {
    setLoadingData(true);
    setParseProgress(10);
    setDataError("");
    setTableData([]);
    setHeaders([]);
    setSheetUsed("");
    setSearch("");
    setPage(0);
    setSortCol(null);
    setSortDir("asc");
    setColFilters({});

    try {
      setParseProgress(20);
      const res = await fetch(`/api/master/file?id=${file.id}`);
      if (!res.ok) throw new Error("ดาวน์โหลดไฟล์ไม่สำเร็จ");
      const buffer = await res.arrayBuffer();
      setParseProgress(60);

      await new Promise<void>((r) => setTimeout(r, 0));
      const wb = XLSX.read(buffer, { type: "array" });

      const targetSheet =
        wb.SheetNames.find((n) => n === "Fixture_2026") ??
        wb.SheetNames.find((n) => n.startsWith("Fixture")) ??
        wb.SheetNames[0];

      if (!targetSheet) throw new Error("ไม่พบ sheet ในไฟล์");

      const ws = wb.Sheets[targetSheet];

      // range: 1 = skip the first remark row, use row 2 as header
      const raw = XLSX.utils.sheet_to_json<DataRow>(ws, { defval: "", range: 1 });
      setParseProgress(90);

      // Filter out _EMPTY, _EMPTY_1 etc. — xlsx auto-generates these for blank cells
      const hdrs = raw.length > 0
        ? Object.keys(raw[0]).filter((h) => !h.startsWith("_"))
        : [];
      setSheetUsed(targetSheet);
      setHeaders(hdrs);
      setTableData(raw);
      // fixture data is now uploaded directly in NewRenovateTab
      setParseProgress(100);
    } catch (err) {
      setDataError(String(err));
    } finally {
      setLoadingData(false);
    }
  };

  useEffect(() => {
    fetchLatest().then((file) => { if (file) loadData(file); });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Upload handlers ───────────────────────────────────────────────────────
  const handleFileSelect = (files: FileList | null) => {
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
    if (!uploading) handleFileSelect(e.dataTransfer.files);
  };

  // ── Sort + Filter ─────────────────────────────────────────────────────────
  const handleSort = (col: string) => {
    if (sortCol === col) {
      if (sortDir === "asc") setSortDir("desc");
      else { setSortCol(null); setSortDir("asc"); }
    } else {
      setSortCol(col);
      setSortDir("asc");
    }
    setPage(0);
  };

  const handleColFilter = (col: string, val: string) => {
    setColFilters((p) => ({ ...p, [col]: val }));
    setPage(0);
  };

  const clearAllFilters = () => {
    setColFilters({});
    setSearch("");
    setSortCol(null);
    setSortDir("asc");
    setPage(0);
  };

  const activeFilterCount =
    Object.values(colFilters).filter((v) => v.trim()).length + (search.trim() ? 1 : 0);

  const displayData = useMemo(() => {
    let data = tableData;

    // Per-column filters
    const activeFilters = Object.entries(colFilters).filter(([, v]) => v.trim());
    if (activeFilters.length > 0)
      data = data.filter((row) =>
        activeFilters.every(([col, val]) =>
          (row[col] || "").toLowerCase().includes(val.toLowerCase())
        )
      );

    // Global search
    if (search.trim()) {
      const q = search.toLowerCase();
      data = data.filter((row) =>
        Object.values(row).some((v) => v.toLowerCase().includes(q))
      );
    }

    // Sort
    if (sortCol) {
      data = [...data].sort((a, b) => {
        const av = a[sortCol] || "", bv = b[sortCol] || "";
        const an = Number(av), bn = Number(bv);
        const isNum = av !== "" && bv !== "" && !isNaN(an) && !isNaN(bn);
        const cmp = isNum ? an - bn : av.localeCompare(bv, "th");
        return sortDir === "asc" ? cmp : -cmp;
      });
    }

    return data;
  }, [tableData, colFilters, search, sortCol, sortDir]);

  const totalPages = Math.ceil(displayData.length / PAGE_SIZE);
  const pageRows = displayData.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

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
              <h2 className="font-bold text-slate-800 text-lg">Fixture Index</h2>
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
              <input ref={inputRef} type="file" accept=".xlsx" className="hidden"
                onChange={(e) => { handleFileSelect(e.target.files); e.target.value = ""; }} />
              {uploading ? (
                <div className="animate-spin rounded-full h-8 w-8 border-4 border-pink-200 border-t-[#E91E8C]" />
              ) : selectedFile ? (
                <CheckCircle className="w-8 h-8 text-green-500" />
              ) : (
                <CloudUpload className={`w-8 h-8 ${dragging ? "text-[#E91E8C]" : "text-pink-300"}`} />
              )}
              <div className="text-center">
                {uploading
                  ? <p className="text-sm text-slate-500">กำลังอัปโหลดไปยัง Google Drive...</p>
                  : selectedFile ? (
                    <><p className="font-semibold text-green-700 text-sm">{selectedFile.name}</p>
                    <p className="text-xs text-slate-400 mt-0.5">คลิกเพื่อเลือกไฟล์ใหม่</p></>
                  ) : (
                    <><p className="font-semibold text-slate-700 text-sm">เลือกไฟล์ Fixture_Index</p>
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
              <h2 className="font-bold text-slate-800 text-lg">
                ข้อมูลใน{sheetUsed ? ` ${sheetUsed}` : " Fixture_2026"}
              </h2>
              {tableData.length > 0 && (
                <span className="text-xs bg-pink-100 text-[#E91E8C] px-2 py-0.5 rounded-full font-medium">
                  {tableData.length.toLocaleString()} แถว
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
              {/* Global search */}
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="ค้นหาทุกคอลัมน์..."
                  value={search}
                  onChange={(e) => { setSearch(e.target.value); setPage(0); }}
                  className="pl-8 pr-3 py-1.5 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-pink-200 focus:border-[#E91E8C] w-52"
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

              {activeFilterCount > 0 && (
                <button
                  onClick={clearAllFilters}
                  className="text-xs text-slate-400 hover:text-red-500 transition-colors flex items-center gap-1"
                >
                  <X className="w-3 h-3" />ล้างทั้งหมด
                </button>
              )}
            </div>
          )}
        </div>

        {/* Loading state */}
        {loadingData && (
          <div className="flex flex-col items-center justify-center py-16 gap-4 text-slate-500">
            <div className="animate-spin rounded-full h-8 w-8 border-4 border-pink-200 border-t-[#E91E8C]" />
            <p className="text-sm font-medium">
              {parseProgress < 50 ? "กำลังดาวน์โหลดไฟล์..." : `กำลังประมวลผลข้อมูล... (${parseProgress}%)`}
            </p>
            <div className="w-64 bg-slate-200 rounded-full h-2 overflow-hidden">
              <div
                className="h-2 rounded-full transition-all duration-300 bg-gradient-to-r from-[#E91E8C] via-[#F15A22] to-[#FFD100]"
                style={{ width: `${parseProgress}%` }}
              />
            </div>
          </div>
        )}

        {/* Error state */}
        {!loadingData && dataError && (
          <div className="m-6 flex items-start gap-2 px-4 py-3 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
            <XCircle className="w-4 h-4 flex-shrink-0 mt-0.5" /><span>{dataError}</span>
          </div>
        )}

        {/* Empty — no file in Drive */}
        {!loadingData && !dataError && !latestFile && (
          <div className="flex flex-col items-center justify-center py-16 gap-2 text-slate-400">
            <CloudUpload className="w-10 h-10 text-pink-200" />
            <p className="text-sm">ยังไม่มีไฟล์ใน Google Drive</p>
            <p className="text-xs">คลิก &quot;อัปโหลดไฟล์ใหม่&quot; ด้านบนเพื่อเริ่มต้น</p>
          </div>
        )}

        {/* Empty — file exists but no rows parsed */}
        {!loadingData && !dataError && latestFile && tableData.length === 0 && (
          <div className="flex items-center justify-center py-16 text-slate-400 text-sm">ไม่พบข้อมูลในไฟล์</div>
        )}

        {/* ── Table ── */}
        {!loadingData && !dataError && tableData.length > 0 && (
          <>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    {headers.map((h) => (
                      <th
                        key={h}
                        onClick={() => handleSort(h)}
                        className="px-4 py-2.5 text-left text-xs font-semibold text-slate-600 whitespace-nowrap cursor-pointer select-none transition-colors group hover:bg-slate-100"
                      >
                        <div className="flex items-center gap-1">
                          {h}
                          {sortCol === h
                            ? sortDir === "asc"
                              ? <ArrowUp className="w-3 h-3 text-[#E91E8C]" />
                              : <ArrowDown className="w-3 h-3 text-[#E91E8C]" />
                            : <ArrowUpDown className="w-3 h-3 text-slate-300 group-hover:text-slate-400" />}
                        </div>
                      </th>
                    ))}
                  </tr>

                  {/* Per-column filter row */}
                  {showColFilters && (
                    <tr className="bg-white border-b border-slate-200">
                      {headers.map((h) => (
                        <th key={h} className="px-2 py-1.5">
                          <input
                            type="text"
                            value={colFilters[h] || ""}
                            onChange={(e) => handleColFilter(h, e.target.value)}
                            placeholder="filter..."
                            className="w-full px-2 py-1 text-xs border border-slate-200 rounded focus:outline-none focus:ring-1 focus:ring-pink-300 focus:border-[#E91E8C] font-normal min-w-[80px]"
                          />
                        </th>
                      ))}
                    </tr>
                  )}
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pageRows.map((row, i) => (
                    <tr key={i} className="hover:bg-pink-50/40 transition-colors">
                      {headers.map((h) => (
                        <td key={h} className="px-4 py-2.5 text-slate-700 text-xs whitespace-nowrap">
                          {row[h] ?? ""}
                        </td>
                      ))}
                    </tr>
                  ))}
                  {pageRows.length === 0 && (
                    <tr>
                      <td colSpan={headers.length} className="px-4 py-10 text-center text-slate-400 text-sm">
                        ไม่พบข้อมูลที่ตรงกับเงื่อนไข
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {totalPages > 1 && (
              <div className="px-6 py-3 border-t border-slate-100 flex items-center justify-between text-sm text-slate-500 bg-white sticky bottom-0">
                <span>
                  แสดง {page * PAGE_SIZE + 1}–{Math.min((page + 1) * PAGE_SIZE, displayData.length).toLocaleString()}{" "}
                  จาก {displayData.length.toLocaleString()} แถว
                  {(search || Object.values(colFilters).some(Boolean)) &&
                    ` (กรองจาก ${tableData.length.toLocaleString()})`}
                </span>
                <div className="flex items-center gap-1">
                  <button
                    onClick={() => setPage((p) => Math.max(0, p - 1))}
                    disabled={page === 0}
                    className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed"
                  >
                    <ChevronLeft className="w-4 h-4" />
                  </button>
                  <span className="px-3 font-medium">{page + 1} / {totalPages}</span>
                  <button
                    onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))}
                    disabled={page >= totalPages - 1}
                    className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed"
                  >
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
