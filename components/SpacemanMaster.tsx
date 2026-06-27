"use client";

import { useRef, useState, useEffect, useMemo, DragEvent } from "react";
import * as XLSX from "xlsx";
import {
  CloudUpload,
  CheckCircle,
  XCircle,
  Clock,
  RefreshCw,
  Search,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  ChevronUp,
  Database,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Filter,
} from "lucide-react";

const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "";

export interface DriveFileInfo {
  id: string;
  name: string;
  createdTime: string;
}

export function formatDateTime(isoString: string): string {
  const d = new Date(isoString);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

interface Props {
  onFileInfoChange: (info: DriveFileInfo | null) => void;
}

type DataRow = Record<string, string>;

const PAGE_SIZE = 100;

export default function SpacemanMaster({ onFileInfoChange }: Props) {
  const [latestFile, setLatestFile] = useState<DriveFileInfo | null>(null);
  const [loadingMeta, setLoadingMeta] = useState(true);

  const [headers, setHeaders] = useState<string[]>([]);
  const [tableData, setTableData] = useState<DataRow[]>([]);
  const [loadingData, setLoadingData] = useState(false);
  const [dataError, setDataError] = useState("");

  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const [sortCol, setSortCol] = useState<string | null>(null);
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [showColFilters, setShowColFilters] = useState(false);

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

  // Load Google Identity Services script and init token client with persistent callback
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
            const formData = new FormData();
            formData.append("file", file);
            formData.append("accessToken", resp.access_token);
            const res = await fetch("/api/spaceman/upload", { method: "POST", body: formData });
            if (!res.ok) {
              const data = await res.json();
              throw new Error(data.error || "อัปโหลดล้มเหลว");
            }
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

    if ((window as any).google?.accounts?.oauth2) {
      initClient();
    } else {
      const script = document.createElement("script");
      script.src = "https://accounts.google.com/gsi/client";
      script.async = true;
      script.onload = initClient;
      document.head.appendChild(script);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

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
    setLoadingData(true);
    setDataError("");
    setTableData([]);
    setHeaders([]);
    setPage(0);
    setSortCol(null);
    setSortDir("asc");
    setColFilters({});
    setSearch("");
    try {
      const res = await fetch(`/api/spaceman/file?id=${file.id}`);
      if (!res.ok) throw new Error("ดาวน์โหลดไฟล์ไม่สำเร็จ");
      const buffer = await res.arrayBuffer();

      const wb = XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets["QRY_Product_by_POG"];
      if (!ws) throw new Error('ไม่พบ Sheet "QRY_Product_by_POG" ในไฟล์');

      const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

      const hdrs: string[] = [];
      for (let c = 0; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
        hdrs.push(cell?.v != null ? String(cell.v).trim() : `คอลัมน์ ${c + 1}`);
      }

      const rows: DataRow[] = [];
      for (let r = 1; r <= range.e.r; r++) {
        const row: DataRow = {};
        let hasValue = false;
        for (let c = 0; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          const val = cell?.v != null ? String(cell.v) : "";
          row[hdrs[c]] = val;
          if (val) hasValue = true;
        }
        if (hasValue) rows.push(row);
      }

      setHeaders(hdrs);
      setTableData(rows);
    } catch (err) {
      setDataError(String(err));
    } finally {
      setLoadingData(false);
    }
  };

  useEffect(() => {
    fetchLatest().then((file) => {
      if (file) loadData(file);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Step 1: user picks a file — just store it, don't trigger OAuth yet
  const handleFiles = (files: FileList | null) => {
    if (!files || files.length === 0) return;
    setSelectedFile(files[0]);
    setUploadStatus("idle");
    setUploadError("");
  };

  // Step 2: user clicks the upload button — trigger OAuth popup (direct user gesture)
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

  const filtered = useMemo(() => {
    let data = tableData;

    // Apply per-column filters
    const activeFilters = Object.entries(colFilters).filter(([, v]) => v.trim());
    if (activeFilters.length > 0) {
      data = data.filter((row) =>
        activeFilters.every(([col, val]) =>
          (row[col] || "").toLowerCase().includes(val.toLowerCase())
        )
      );
    }

    // Apply global search
    if (search.trim()) {
      const q = search.toLowerCase();
      data = data.filter((row) =>
        Object.values(row).some((v) => v.toLowerCase().includes(q))
      );
    }

    // Apply sort
    if (sortCol) {
      data = [...data].sort((a, b) => {
        const av = a[sortCol] || "";
        const bv = b[sortCol] || "";
        const aNum = Number(av);
        const bNum = Number(bv);
        const isNum = av !== "" && bv !== "" && !isNaN(aNum) && !isNaN(bNum);
        const cmp = isNum ? aNum - bNum : av.localeCompare(bv, "th");
        return sortDir === "asc" ? cmp : -cmp;
      });
    }

    return data;
  }, [tableData, search, colFilters, sortCol, sortDir]);

  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const pageData = filtered.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  const handleSearch = (val: string) => {
    setSearch(val);
    setPage(0);
  };

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
    setColFilters((prev) => ({ ...prev, [col]: val }));
    setPage(0);
  };

  const clearAllFilters = () => {
    setColFilters({});
    setSearch("");
    setSortCol(null);
    setSortDir("asc");
    setPage(0);
  };

  const activeFilterCount = Object.values(colFilters).filter((v) => v.trim()).length + (search.trim() ? 1 : 0);

  return (
    <div className="space-y-6">
      {/* Upload section */}
      <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#00A6E2] via-[#FFD100] via-[#F15A22] to-[#72BF44]" />

        {/* Header row */}
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
                  <strong className="text-slate-700">
                    {formatDateTime(latestFile.createdTime)}
                  </strong>
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

        {/* Upload panel (collapsible) */}
        {showUpload && (
          <div className="px-6 pb-6 space-y-3 border-t border-pink-50 pt-4">
            {/* Step 1: File picker drop zone */}
            <div
              onClick={() => !uploading && inputRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              className={`
                flex flex-col items-center justify-center gap-2 border-2 border-dashed rounded-xl p-8 transition-all duration-200
                ${uploading ? "opacity-60 cursor-not-allowed border-pink-200 bg-pink-50/30"
                  : dragging ? "border-[#E91E8C] bg-pink-50 scale-[1.01] cursor-pointer"
                  : selectedFile ? "border-green-300 bg-green-50/40 cursor-pointer"
                  : "border-pink-200 bg-pink-50/30 hover:border-[#E91E8C] hover:bg-pink-50 cursor-pointer"}
              `}
            >
              <input
                ref={inputRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }}
              />
              {uploading ? (
                <div className="animate-spin rounded-full h-8 w-8 border-4 border-pink-200 border-t-[#E91E8C]" />
              ) : selectedFile ? (
                <CheckCircle className="w-8 h-8 text-green-500" />
              ) : (
                <CloudUpload className={`w-8 h-8 ${dragging ? "text-[#E91E8C]" : "text-pink-300"}`} />
              )}
              <div className="text-center">
                {uploading ? (
                  <p className="text-sm text-slate-500">กำลังอัปโหลดไปยัง Google Drive...</p>
                ) : selectedFile ? (
                  <>
                    <p className="font-semibold text-green-700 text-sm">{selectedFile.name}</p>
                    <p className="text-xs text-slate-400 mt-0.5">คลิกเพื่อเลือกไฟล์ใหม่</p>
                  </>
                ) : (
                  <>
                    <p className="font-semibold text-slate-700 text-sm">เลือกไฟล์ DATA_SPACEMAN</p>
                    <p className="text-xs text-slate-400 mt-0.5">คลิกหรือลากไฟล์ .xlsx มาวางที่นี่</p>
                  </>
                )}
              </div>
            </div>

            {/* Step 2: Confirm upload button (only shows after file selected) */}
            {selectedFile && !uploading && uploadStatus !== "success" && (
              <button
                onClick={handleUpload}
                className="w-full flex items-center justify-center gap-2 py-3 rounded-xl font-semibold text-sm bg-gradient-to-r from-[#E91E8C] to-[#d41679] text-white hover:from-[#d41679] hover:to-[#be185d] transition-all shadow-sm"
              >
                <CloudUpload className="w-4 h-4" />
                อัปโหลดขึ้น Google Drive
              </button>
            )}

            {uploadStatus === "success" && (
              <div className="flex items-center gap-2 px-4 py-2.5 bg-green-50 border border-green-200 rounded-xl text-green-700 text-sm">
                <CheckCircle className="w-4 h-4 flex-shrink-0" />
                อัปโหลดสำเร็จ! ข้อมูลด้านล่างได้รับการอัปเดตแล้ว
              </div>
            )}
            {uploadStatus === "error" && (
              <div className="flex items-start gap-2 px-4 py-2.5 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
                <XCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
                <span>อัปโหลดล้มเหลว: {uploadError}</span>
              </div>
            )}
          </div>
        )}
      </div>

      {/* Data table section */}
      <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#00A6E2] via-[#FFD100] via-[#F15A22] to-[#72BF44]" />
        <div className="px-6 py-4 border-b border-pink-50 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-1 h-5 rounded-full bg-gradient-to-b from-[#E91E8C] to-[#F15A22]" />
            <div className="flex items-center gap-2">
              <Database className="w-4 h-4 text-slate-400" />
              <h2 className="font-bold text-slate-800 text-lg">
                ข้อมูลใน QRY_Product_by_POG
              </h2>
              {tableData.length > 0 && (
                <span className="text-xs bg-pink-100 text-[#E91E8C] px-2 py-0.5 rounded-full font-medium">
                  {tableData.length.toLocaleString()} แถว
                </span>
              )}
              {filtered.length !== tableData.length && (
                <span className="text-xs bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full font-medium">
                  กรองแล้ว: {filtered.length.toLocaleString()} แถว
                </span>
              )}
            </div>
          </div>

          {tableData.length > 0 && (
            <div className="flex items-center gap-2">
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
                  className="text-xs text-slate-400 hover:text-red-500 transition-colors"
                >
                  ล้างทั้งหมด
                </button>
              )}
            </div>
          )}
        </div>

        <div className="overflow-x-auto">
          {loadingData ? (
            <div className="flex flex-col items-center justify-center py-16 gap-3 text-slate-500">
              <div className="animate-spin rounded-full h-8 w-8 border-4 border-pink-200 border-t-[#E91E8C]" />
              <p className="text-sm">กำลังโหลดข้อมูลจาก Google Drive...</p>
            </div>
          ) : dataError ? (
            <div className="m-6 flex items-start gap-2 px-4 py-3 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
              <XCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
              <span>{dataError}</span>
            </div>
          ) : !latestFile ? (
            <div className="flex flex-col items-center justify-center py-16 gap-2 text-slate-400">
              <CloudUpload className="w-10 h-10 text-pink-200" />
              <p className="text-sm">ยังไม่มีไฟล์ใน Google Drive</p>
              <p className="text-xs">คลิก "อัปโหลดไฟล์ใหม่" ด้านบนเพื่อเริ่มต้น</p>
            </div>
          ) : tableData.length === 0 ? (
            <div className="flex items-center justify-center py-16 text-slate-400 text-sm">
              ไม่พบข้อมูล
            </div>
          ) : (
            <>
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    {headers.map((h) => (
                      <th
                        key={h}
                        onClick={() => handleSort(h)}
                        className="px-4 py-2.5 text-left text-xs font-semibold text-slate-600 whitespace-nowrap cursor-pointer select-none hover:bg-slate-100 transition-colors group"
                      >
                        <div className="flex items-center gap-1">
                          {h}
                          {sortCol === h ? (
                            sortDir === "asc"
                              ? <ArrowUp className="w-3 h-3 text-[#E91E8C]" />
                              : <ArrowDown className="w-3 h-3 text-[#E91E8C]" />
                          ) : (
                            <ArrowUpDown className="w-3 h-3 text-slate-300 group-hover:text-slate-400" />
                          )}
                        </div>
                      </th>
                    ))}
                  </tr>
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
                  {pageData.map((row, i) => (
                    <tr key={i} className="hover:bg-pink-50/40 transition-colors">
                      {headers.map((h) => (
                        <td
                          key={h}
                          className="px-4 py-2 text-slate-700 whitespace-nowrap max-w-xs truncate"
                        >
                          {search.trim()
                            ? highlightMatch(row[h] || "", search)
                            : (row[h] || "")}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Pagination */}
              {totalPages > 1 && (
                <div className="px-6 py-3 border-t border-slate-100 flex items-center justify-between text-sm text-slate-500">
                  <span>
                    แสดง {page * PAGE_SIZE + 1}–{Math.min((page + 1) * PAGE_SIZE, filtered.length)} จาก{" "}
                    {filtered.length.toLocaleString()} แถว
                    {search && ` (กรองจาก ${tableData.length.toLocaleString()})`}
                  </span>
                  <div className="flex items-center gap-1">
                    <button
                      onClick={() => setPage((p) => Math.max(0, p - 1))}
                      disabled={page === 0}
                      className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <span className="px-3 font-medium">
                      {page + 1} / {totalPages}
                    </span>
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
      <mark className="bg-yellow-200 text-yellow-900 rounded px-0.5">
        {text.slice(idx, idx + query.length)}
      </mark>
      {text.slice(idx + query.length)}
    </>
  );
}
