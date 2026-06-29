"use client";
import { useState, useRef, useEffect } from "react";
import {
  Plus, X, Settings2, CloudUpload, Loader2, CheckCircle,
  AlertTriangle, Search, Pencil, Copy, Save,
  ChevronUp, ChevronDown, ChevronsUpDown,
} from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";

export const EXCEPTION_CONFIG_KEY = "recap_exception_config";
const ALL = "ทั้งหมด";
const PAGE_SIZE = 10;

export type SyncStatus = "idle" | "loading" | "saving" | "saved" | "error";

interface Props {
  config: ExceptionConfig[];
  onChange: (updated: ExceptionConfig[]) => void;
  onClose: () => void;
  categories: string[];
  subcategories: string[];
  descCList: string[];
  spacemanLoaded: boolean;
  syncStatus: SyncStatus;
  lastSaved: string | null;
  syncError: string;
}

type DraftFields = Omit<ExceptionConfig, "id" | "createdAt" | "updatedAt">;
type SortDir = "asc" | "desc";
type ColFilters = { category: string; subcategory: string; descC: string; percentage: string; status: string };

const emptyDraft = (): DraftFields => ({
  category: "", subcategory: "", descC: "", percentage: "", status: "active",
});

const emptyFilters = (): ColFilters => ({
  category: "", subcategory: "", descC: "", percentage: "", status: "",
});

function fmtDate(iso: string): string {
  if (!iso) return "—";
  const d = new Date(iso);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// ── Sort icon ─────────────────────────────────────────────────────────────────
function SortIcon({ col, sortCol, sortDir }: { col: string; sortCol: string | null; sortDir: SortDir }) {
  if (sortCol !== col) return <ChevronsUpDown className="w-3 h-3 text-slate-300 ml-0.5 flex-shrink-0" />;
  return sortDir === "asc"
    ? <ChevronUp className="w-3 h-3 text-[#E91E8C] ml-0.5 flex-shrink-0" />
    : <ChevronDown className="w-3 h-3 text-[#E91E8C] ml-0.5 flex-shrink-0" />;
}

// ── Searchable combobox ───────────────────────────────────────────────────────
function SearchSelect({
  label, value, options, onChange, loading = false,
}: {
  label?: string; value: string; options: string[]; onChange: (v: string) => void; loading?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const wrapRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) {
        setOpen(false); setSearch("");
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const allOptions = [ALL, ...options];
  const filtered = search
    ? allOptions.filter((o) => o.toLowerCase().includes(search.toLowerCase()))
    : allOptions;

  const select = (v: string) => { onChange(v); setOpen(false); setSearch(""); };

  return (
    <div ref={wrapRef} className="flex flex-col gap-1 relative">
      {label && <label className="text-xs font-medium text-slate-500">{label}</label>}
      <button
        type="button"
        disabled={loading}
        onClick={() => { if (!loading) { setOpen((o) => !o); setSearch(""); } }}
        className={`text-xs border rounded-lg px-2 py-1.5 bg-white text-left flex items-center justify-between gap-1 focus:outline-none transition-colors min-w-0 ${
          loading
            ? "border-slate-100 bg-slate-50 cursor-not-allowed opacity-70"
            : "border-slate-200 hover:border-pink-300 focus:ring-2 focus:ring-pink-200 focus:border-[#E91E8C]"
        }`}
      >
        {loading ? (
          <span className="flex items-center gap-1.5 text-slate-400 italic">
            <Loader2 className="w-3 h-3 animate-spin flex-shrink-0" />
            กำลังโหลด...
          </span>
        ) : (
          <span className={`truncate ${!value ? "text-slate-300" : value === ALL ? "text-slate-400 italic" : "text-slate-700"}`}>
            {value || "— เลือก —"}
          </span>
        )}
        {!loading && (
          <svg className="w-3 h-3 text-slate-400 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
          </svg>
        )}
      </button>
      {open && (
        <div className="absolute top-full mt-1 left-0 z-[200] bg-white border border-slate-200 rounded-xl shadow-xl flex flex-col" style={{ minWidth: "220px", maxWidth: "340px" }}>
          <div className="flex items-center gap-1.5 px-2 py-1.5 border-b border-slate-100">
            <Search className="w-3.5 h-3.5 text-slate-400 flex-shrink-0" />
            <input
              autoFocus
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="ค้นหา..."
              className="text-xs flex-1 outline-none bg-transparent placeholder-slate-300"
            />
          </div>
          <div className="overflow-y-auto max-h-52">
            {filtered.length === 0
              ? <p className="text-xs text-slate-400 px-3 py-2">ไม่พบข้อมูล</p>
              : filtered.map((o) => (
                <button
                  key={o}
                  type="button"
                  onClick={() => select(o)}
                  className={`w-full text-left text-xs px-3 py-1.5 hover:bg-pink-50 transition-colors ${
                    o === value ? "bg-pink-50 text-[#E91E8C] font-semibold" : o === ALL ? "text-slate-400 italic" : "text-slate-700"
                  }`}
                >
                  {o}
                </button>
              ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Main component ────────────────────────────────────────────────────────────
export default function ConfigMenu({
  config, onChange, onClose,
  categories, subcategories, descCList,
  spacemanLoaded, syncStatus, lastSaved, syncError,
}: Props) {
  const [draft, setDraft] = useState<DraftFields>(emptyDraft());
  const [editId, setEditId] = useState<string | null>(null);
  const [formError, setFormError] = useState("");

  // Table state
  const [sortCol, setSortCol] = useState<string | null>(null);
  const [sortDir, setSortDir] = useState<SortDir>("asc");
  const [colFilters, setColFilters] = useState<ColFilters>(emptyFilters());
  const [page, setPage] = useState(0);

  const isEditing = editId !== null;
  const set = (k: keyof DraftFields, v: string) => setDraft((d) => ({ ...d, [k]: v }));

  // ── Computed table data ──
  const afterFilter = config.filter((e) =>
    (!colFilters.category || e.category.toLowerCase().includes(colFilters.category.toLowerCase())) &&
    (!colFilters.subcategory || e.subcategory.toLowerCase().includes(colFilters.subcategory.toLowerCase())) &&
    (!colFilters.descC || e.descC.toLowerCase().includes(colFilters.descC.toLowerCase())) &&
    (!colFilters.percentage || e.percentage.includes(colFilters.percentage)) &&
    (!colFilters.status || e.status === colFilters.status)
  );

  const afterSort = sortCol
    ? [...afterFilter].sort((a, b) => {
        const av = String((a as unknown as Record<string, unknown>)[sortCol] ?? "");
        const bv = String((b as unknown as Record<string, unknown>)[sortCol] ?? "");
        return sortDir === "asc" ? av.localeCompare(bv, "th") : bv.localeCompare(av, "th");
      })
    : afterFilter;

  const totalPages = Math.max(1, Math.ceil(afterSort.length / PAGE_SIZE));
  const safePage = Math.min(page, totalPages - 1);
  const pageData = afterSort.slice(safePage * PAGE_SIZE, (safePage + 1) * PAGE_SIZE);

  const toggleSort = (col: string) => {
    if (sortCol === col) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortCol(col); setSortDir("asc"); }
    setPage(0);
  };

  const setFilter = (key: keyof ColFilters, value: string) => {
    setColFilters((f) => ({ ...f, [key]: value }));
    setPage(0);
  };

  const hasFilter = Object.values(colFilters).some(Boolean);

  // ── Form actions ──
  const startEdit = (entry: ExceptionConfig) => {
    setEditId(entry.id);
    setDraft({ category: entry.category, subcategory: entry.subcategory, descC: entry.descC, percentage: entry.percentage, status: entry.status });
    setFormError("");
    document.getElementById("config-form")?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const cancelEdit = () => { setEditId(null); setDraft(emptyDraft()); setFormError(""); };

  const copyEntry = (entry: ExceptionConfig) => {
    setEditId(null); // add mode — duplicate check will fire on submit
    setDraft({ category: entry.category, subcategory: entry.subcategory, descC: entry.descC, percentage: entry.percentage, status: entry.status });
    setFormError("");
    document.getElementById("config-form")?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const isDuplicate = (d: DraftFields, excludeId?: string) =>
    config.some((e) => e.id !== excludeId && e.category === d.category && e.subcategory === d.subcategory && e.descC === d.descC);

  const submitForm = () => {
    if (!draft.category || !draft.subcategory || !draft.descC) {
      setFormError("กรุณาเลือกค่าให้ครบทุกช่อง (เลือก 'ทั้งหมด' ถ้าต้องการจับคู่ทุก value)");
      return;
    }
    const pct = parseFloat(draft.percentage);
    if (isNaN(pct) || pct <= 0 || pct > 100) { setFormError("Percentage ต้องเป็นตัวเลข 1–100"); return; }
    if (isDuplicate(draft, isEditing ? editId! : undefined)) {
      setFormError("มี Rule ที่ใช้ CATEGORY / SUBCATEGORY / DESC_C นี้อยู่แล้ว — กรุณาแก้ไข Rule เดิมแทน");
      return;
    }
    setFormError("");
    const now = new Date().toISOString();
    if (isEditing) {
      onChange(config.map((e) => e.id === editId ? { ...e, ...draft, updatedAt: now } : e));
      setEditId(null);
    } else {
      onChange([...config, { ...draft, id: crypto.randomUUID(), createdAt: now, updatedAt: now }]);
    }
    setDraft(emptyDraft());
  };

  const toggleStatus = (id: string) =>
    onChange(config.map((e) => e.id === id ? { ...e, status: e.status === "active" ? "inactive" : "active" } : e));

  // ── SyncBadge ──
  const SyncBadge = () => {
    if (syncStatus === "loading") return <span className="flex items-center gap-1 text-xs text-slate-400"><Loader2 className="w-3.5 h-3.5 animate-spin" /> กำลังโหลด...</span>;
    if (syncStatus === "saving")  return <span className="flex items-center gap-1 text-xs text-amber-500"><Loader2 className="w-3.5 h-3.5 animate-spin" /> กำลังบันทึก...</span>;
    if (syncStatus === "error")   return <span className="flex items-center gap-1 text-xs text-red-500"><AlertTriangle className="w-3.5 h-3.5" /> {syncError || "เชื่อมต่อ Sheets ไม่ได้"}</span>;
    if (syncStatus === "saved" && lastSaved) return <span className="flex items-center gap-1 text-xs text-green-600"><CheckCircle className="w-3.5 h-3.5" /><CloudUpload className="w-3.5 h-3.5" /> บันทึกแล้ว {fmtDate(lastSaved)}</span>;
    if (lastSaved) return <span className="flex items-center gap-1 text-xs text-slate-400"><CloudUpload className="w-3.5 h-3.5" /> อัปเดตล่าสุด {fmtDate(lastSaved)}</span>;
    return null;
  };

  // ── Render ──
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
      {/* Fixed-height modal */}
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-6xl h-[90vh] flex flex-col">

        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100 flex-shrink-0">
          <div className="flex items-center gap-3">
            <Settings2 className="w-5 h-5 text-[#E91E8C]" />
            <div>
              <h2 className="font-bold text-slate-800 text-base">Config — ข้อยกเว้นคอลัมน์ O (%)</h2>
              <SyncBadge />
            </div>
          </div>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-400 transition-colors">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Scrollable content */}
        <div className="overflow-y-auto flex-1 px-6 py-4 space-y-4">

          {/* Info banner — improved layout */}
          <div className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 space-y-2">
            <p className="text-xs font-semibold text-slate-600">วิธีทำงาน</p>
            <div className="grid grid-cols-2 gap-x-6 gap-y-1.5 text-xs text-slate-600">
              <div className="flex gap-2">
                <span className="text-slate-400 flex-shrink-0">📌</span>
                <span><strong>Default คอลัมน์ O = 100%</strong> — ใช้กับทุกแถวที่ไม่มี Rule ตรง</span>
              </div>
              <div className="flex gap-2">
                <span className="text-slate-400 flex-shrink-0">⚡</span>
                <span>Rule ที่ตรงเงื่อนไข → ใช้ <strong>% ที่กำหนดไว้แทน</strong> Default</span>
              </div>
              <div className="flex gap-2">
                <span className="text-slate-400 flex-shrink-0">🔍</span>
                <span>เลือก <strong className="text-[#E91E8C]">ทั้งหมด</strong> ในช่องใด = จับคู่ทุก value ในช่องนั้น</span>
              </div>
              <div className="flex gap-2">
                <span className="text-slate-400 flex-shrink-0">🏆</span>
                <span>Rule <strong>บนสุด = ความสำคัญสูงสุด</strong> (ถ้าตรงหลาย Rule ใช้อันแรก)</span>
              </div>
            </div>
          </div>

          {/* Add / Edit form */}
          <div
            id="config-form"
            className={`rounded-xl border p-4 space-y-3 transition-colors ${
              isEditing ? "bg-amber-50/60 border-amber-200" : "bg-pink-50/60 border-pink-100"
            }`}
          >
            <p className={`text-xs font-semibold ${isEditing ? "text-amber-600" : "text-[#E91E8C]"}`}>
              {isEditing ? "✏️ แก้ไข Rule" : "เพิ่ม Rule ใหม่"}
            </p>

            {!spacemanLoaded && (
              <div className="flex items-center gap-2 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
                <Loader2 className="w-3.5 h-3.5 text-amber-500 animate-spin flex-shrink-0" />
                <p className="text-xs text-amber-700">กำลังโหลดข้อมูลจาก DATA_SPACEMAN — dropdown จะพร้อมใช้งานหลังโหลดเสร็จ</p>
              </div>
            )}

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <SearchSelect label="CATEGORY" value={draft.category} options={categories} onChange={(v) => set("category", v)} loading={!spacemanLoaded} />
              <SearchSelect label="SUBCATEGORY" value={draft.subcategory} options={subcategories} onChange={(v) => set("subcategory", v)} loading={!spacemanLoaded} />
              <SearchSelect label="DESC_C" value={draft.descC} options={descCList} onChange={(v) => set("descC", v)} loading={!spacemanLoaded} />
              <div className="flex flex-col gap-1">
                <label className="text-xs font-medium text-slate-500">Percentage (%)</label>
                <input
                  type="number" min={1} max={100} step={0.01}
                  value={draft.percentage}
                  onChange={(e) => set("percentage", e.target.value)}
                  placeholder="เช่น 50"
                  className="text-xs border border-slate-200 rounded-lg px-2 py-1.5 bg-white focus:outline-none focus:ring-2 focus:ring-pink-200 focus:border-[#E91E8C]"
                />
              </div>
            </div>

            <div className="flex items-center gap-3">
              <span className="text-xs text-slate-500">Status:</span>
              <button
                type="button"
                onClick={() => set("status", draft.status === "active" ? "inactive" : "active")}
                className={`px-3 py-1 rounded-full text-xs font-bold transition-colors ${
                  draft.status === "active" ? "bg-green-100 text-green-700 hover:bg-green-200" : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                }`}
              >
                {draft.status === "active" ? "Active" : "Inactive"}
              </button>
            </div>

            {formError && <p className="text-xs text-red-500">{formError}</p>}

            <div className="flex items-center gap-2">
              <button
                onClick={submitForm}
                disabled={syncStatus === "saving"}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-white text-xs font-semibold rounded-lg disabled:opacity-50 transition-colors ${
                  isEditing ? "bg-amber-500 hover:bg-amber-600" : "bg-[#E91E8C] hover:bg-[#c4187a]"
                }`}
              >
                {isEditing ? <><Save className="w-3.5 h-3.5" /> บันทึก</> : <><Plus className="w-3.5 h-3.5" /> เพิ่ม Rule</>}
              </button>
              {isEditing && (
                <button onClick={cancelEdit} className="px-3 py-1.5 text-slate-500 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-slate-200 transition-colors">
                  ยกเลิก
                </button>
              )}
            </div>
          </div>

          {/* Table */}
          {syncStatus === "loading" ? (
            <div className="flex justify-center py-8"><Loader2 className="w-6 h-6 animate-spin text-pink-300" /></div>
          ) : config.length === 0 ? (
            <p className="text-xs text-slate-400 text-center py-4">ยังไม่มี Rule — ทุกแถวจะใช้ค่า 100% เป็น Default</p>
          ) : (
            <div className="rounded-xl border border-slate-100 flex flex-col overflow-hidden">
              <div className="overflow-x-auto">
                <table className="min-w-full text-xs">
                  <thead>
                    {/* Sort row */}
                    <tr className="bg-slate-50 text-slate-600 border-b border-slate-100">
                      <th className="px-3 py-2 text-left font-semibold w-10">#</th>
                      {(["category","subcategory","descC","percentage","status","createdAt","updatedAt"] as const).map((col) => {
                        const labels: Record<string, string> = {
                          category: "CATEGORY", subcategory: "SUBCATEGORY", descC: "DESC_C",
                          percentage: "%", status: "Status", createdAt: "สร้างเมื่อ", updatedAt: "อัปเดตล่าสุด",
                        };
                        return (
                          <th
                            key={col}
                            onClick={() => toggleSort(col)}
                            className={`px-3 py-2 text-left font-semibold cursor-pointer select-none hover:bg-slate-100 whitespace-nowrap transition-colors ${
                              col === "percentage" || col === "status" ? "text-center" : ""
                            }`}
                          >
                            <span className="inline-flex items-center gap-0.5">
                              {labels[col]}
                              <SortIcon col={col} sortCol={sortCol} sortDir={sortDir} />
                            </span>
                          </th>
                        );
                      })}
                      <th className="px-3 py-2 text-center font-semibold">Actions</th>
                    </tr>
                    {/* Filter row */}
                    <tr className="bg-white border-b border-slate-100">
                      <td className="px-3 py-1.5" />
                      {/* CATEGORY */}
                      <td className="px-2 py-1.5">
                        <input value={colFilters.category} onChange={(e) => setFilter("category", e.target.value)} placeholder="กรอง…" className="w-full text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" />
                      </td>
                      {/* SUBCATEGORY */}
                      <td className="px-2 py-1.5">
                        <input value={colFilters.subcategory} onChange={(e) => setFilter("subcategory", e.target.value)} placeholder="กรอง…" className="w-full text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" />
                      </td>
                      {/* DESC_C */}
                      <td className="px-2 py-1.5">
                        <input value={colFilters.descC} onChange={(e) => setFilter("descC", e.target.value)} placeholder="กรอง…" className="w-full text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" />
                      </td>
                      {/* % */}
                      <td className="px-2 py-1.5 text-center">
                        <input value={colFilters.percentage} onChange={(e) => setFilter("percentage", e.target.value)} placeholder="กรอง…" className="w-16 text-center text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" />
                      </td>
                      {/* Status */}
                      <td className="px-2 py-1.5 text-center">
                        <select value={colFilters.status} onChange={(e) => setFilter("status", e.target.value)} className="text-[11px] border border-slate-200 rounded px-1 py-0.5 focus:outline-none focus:border-pink-300 bg-white text-slate-600">
                          <option value="">ทั้งหมด</option>
                          <option value="active">Active</option>
                          <option value="inactive">Inactive</option>
                        </select>
                      </td>
                      {/* สร้างเมื่อ / อัปเดตล่าสุด / Actions */}
                      <td className="px-3 py-1.5" />
                      <td className="px-3 py-1.5" />
                      <td className="px-3 py-1.5 text-center">
                        {hasFilter && (
                          <button onClick={() => { setColFilters(emptyFilters()); setPage(0); }} className="text-[11px] text-slate-400 hover:text-red-400 transition-colors whitespace-nowrap">ล้าง</button>
                        )}
                      </td>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {pageData.length === 0 ? (
                      <tr><td colSpan={9} className="px-3 py-6 text-center text-xs text-slate-400">ไม่พบรายการที่ตรงกับ Filter</td></tr>
                    ) : pageData.map((entry, i) => {
                      const isThisEdit = editId === entry.id;
                      const globalIdx = safePage * PAGE_SIZE + i + 1;
                      return (
                        <tr
                          key={entry.id}
                          className={`transition-colors ${
                            isThisEdit ? "bg-amber-50 ring-1 ring-inset ring-amber-200"
                              : entry.status === "inactive" ? "opacity-50 hover:opacity-70"
                              : "hover:bg-slate-50"
                          }`}
                        >
                          <td className="px-3 py-2.5 text-slate-400">{globalIdx}</td>
                          <td className="px-3 py-2.5 max-w-[160px]" title={entry.category}>
                            <span className={`truncate block ${entry.category === ALL ? "text-slate-400 italic" : ""}`}>{entry.category}</span>
                          </td>
                          <td className="px-3 py-2.5 max-w-[200px]" title={entry.subcategory}>
                            <span className={`truncate block ${entry.subcategory === ALL ? "text-slate-400 italic" : ""}`}>{entry.subcategory}</span>
                          </td>
                          <td className="px-3 py-2.5 max-w-[160px]" title={entry.descC}>
                            <span className={`truncate block ${entry.descC === ALL ? "text-slate-400 italic" : ""}`}>{entry.descC}</span>
                          </td>
                          <td className="px-3 py-2.5 text-center font-semibold text-[#E91E8C] whitespace-nowrap">{entry.percentage}%</td>
                          <td className="px-3 py-2.5 text-center">
                            <button
                              onClick={() => toggleStatus(entry.id)}
                              disabled={syncStatus === "saving"}
                              className={`px-2 py-0.5 rounded-full text-[10px] font-bold transition-colors disabled:opacity-40 ${
                                entry.status === "active" ? "bg-green-100 text-green-700 hover:bg-green-200" : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                              }`}
                            >
                              {entry.status === "active" ? "Active" : "Inactive"}
                            </button>
                          </td>
                          <td className="px-3 py-2.5 text-slate-400 whitespace-nowrap">{fmtDate(entry.createdAt)}</td>
                          <td className="px-3 py-2.5 text-slate-400 whitespace-nowrap">{fmtDate(entry.updatedAt)}</td>
                          <td className="px-3 py-2.5 text-center">
                            <div className="flex items-center justify-center gap-1">
                              <button
                                onClick={() => startEdit(entry)}
                                title="แก้ไข"
                                disabled={syncStatus === "saving"}
                                className={`p-1.5 rounded-lg transition-colors disabled:opacity-40 ${
                                  isThisEdit ? "bg-amber-100 text-amber-600" : "hover:bg-amber-50 text-slate-400 hover:text-amber-600"
                                }`}
                              >
                                <Pencil className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => copyEntry(entry)}
                                title="คัดลอก → กรอก form"
                                disabled={syncStatus === "saving"}
                                className="p-1.5 rounded-lg hover:bg-blue-50 text-slate-400 hover:text-blue-500 disabled:opacity-40 transition-colors"
                              >
                                <Copy className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Pagination */}
              <div className="flex items-center justify-between px-4 py-2 border-t border-slate-100 bg-slate-50 text-xs text-slate-500 flex-shrink-0">
                <span>
                  {afterSort.length === 0
                    ? "ไม่พบรายการ"
                    : `แสดง ${safePage * PAGE_SIZE + 1}–${Math.min((safePage + 1) * PAGE_SIZE, afterSort.length)} จาก ${afterSort.length} รายการ${hasFilter ? ` (กรองจาก ${config.length})` : ""}`}
                </span>
                <div className="flex items-center gap-1">
                  <button onClick={() => setPage(0)} disabled={safePage === 0} className="px-1.5 py-0.5 rounded hover:bg-slate-200 disabled:opacity-30 transition-colors">«</button>
                  <button onClick={() => setPage((p) => p - 1)} disabled={safePage === 0} className="px-2 py-0.5 rounded hover:bg-slate-200 disabled:opacity-30 transition-colors">‹</button>
                  <span className="px-2">{safePage + 1} / {totalPages}</span>
                  <button onClick={() => setPage((p) => p + 1)} disabled={safePage >= totalPages - 1} className="px-2 py-0.5 rounded hover:bg-slate-200 disabled:opacity-30 transition-colors">›</button>
                  <button onClick={() => setPage(totalPages - 1)} disabled={safePage >= totalPages - 1} className="px-1.5 py-0.5 rounded hover:bg-slate-200 disabled:opacity-30 transition-colors">»</button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer — just close button */}
        <div className="px-6 py-3 border-t border-slate-100 flex justify-end flex-shrink-0">
          <button onClick={onClose} className="px-4 py-2 bg-pink-50 text-[#E91E8C] text-sm font-semibold rounded-xl hover:bg-pink-100 transition-colors">
            ปิด
          </button>
        </div>
      </div>
    </div>
  );
}
