"use client";
import { useState, useRef, useEffect } from "react";
import { Plus, Trash2, X, Settings2, CloudUpload, Loader2, CheckCircle, AlertTriangle, Search } from "lucide-react";
import type { ExceptionConfig } from "@/lib/types";

export const EXCEPTION_CONFIG_KEY = "recap_exception_config";
const ALL = "ทั้งหมด";

export type SyncStatus = "idle" | "loading" | "saving" | "saved" | "error";

interface Props {
  config: ExceptionConfig[];
  onChange: (updated: ExceptionConfig[]) => void;
  onClose: () => void;
  categories: string[];
  subcategories: string[];
  descCList: string[];
  syncStatus: SyncStatus;
  lastSaved: string | null;
  syncError: string;
}

const emptyDraft = (): Omit<ExceptionConfig, "id" | "createdAt" | "updatedAt"> => ({
  category: ALL,
  subcategory: ALL,
  descC: ALL,
  percentage: "100",
  status: "active",
});

function fmtDate(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// ── Searchable combobox ───────────────────────────────────────────────────────
function SearchSelect({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (v: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const wrapRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) {
        setOpen(false);
        setSearch("");
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const allOptions = [ALL, ...options];
  const filtered = search
    ? allOptions.filter((o) => o.toLowerCase().includes(search.toLowerCase()))
    : allOptions;

  const select = (v: string) => {
    onChange(v);
    setOpen(false);
    setSearch("");
  };

  return (
    <div ref={wrapRef} className="flex flex-col gap-1 relative">
      <label className="text-xs font-medium text-slate-500">{label}</label>
      <button
        type="button"
        onClick={() => { setOpen((o) => !o); setSearch(""); }}
        className="text-xs border border-slate-200 rounded-lg px-2 py-1.5 bg-white text-left flex items-center justify-between gap-1 focus:outline-none focus:ring-2 focus:ring-pink-200 focus:border-[#E91E8C] hover:border-pink-300 transition-colors"
      >
        <span className={value === ALL ? "text-slate-400 italic" : "text-slate-700 truncate"}>
          {value}
        </span>
        <svg className="w-3 h-3 text-slate-400 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
          <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
        </svg>
      </button>

      {open && (
        <div className="absolute top-full mt-1 left-0 right-0 z-[60] bg-white border border-slate-200 rounded-xl shadow-xl flex flex-col"
          style={{ minWidth: "220px" }}>
          {/* Search input */}
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
          {/* Options list */}
          <div className="overflow-y-auto max-h-52">
            {filtered.length === 0 ? (
              <p className="text-xs text-slate-400 px-3 py-2">ไม่พบข้อมูล</p>
            ) : (
              filtered.map((o) => (
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
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Main component ────────────────────────────────────────────────────────────
export default function ConfigMenu({
  config,
  onChange,
  onClose,
  categories,
  subcategories,
  descCList,
  syncStatus,
  lastSaved,
  syncError,
}: Props) {
  const [draft, setDraft] = useState(emptyDraft());
  const [error, setError] = useState("");

  const set = (k: keyof typeof draft, v: string) =>
    setDraft((d) => ({ ...d, [k]: v }));

  const addEntry = () => {
    const pct = parseFloat(draft.percentage);
    if (isNaN(pct) || pct <= 0 || pct > 100) {
      setError("Percentage ต้องเป็นตัวเลข 1–100");
      return;
    }
    setError("");
    const now = new Date().toISOString();
    const entry: ExceptionConfig = { ...draft, id: crypto.randomUUID(), createdAt: now, updatedAt: now };
    onChange([...config, entry]);
    setDraft(emptyDraft());
  };

  const remove = (id: string) => onChange(config.filter((e) => e.id !== id));

  const toggleStatus = (id: string) =>
    onChange(config.map((e) =>
      e.id === id ? { ...e, status: e.status === "active" ? "inactive" : "active" } : e
    ));

  const SyncBadge = () => {
    if (syncStatus === "loading") return (
      <span className="flex items-center gap-1 text-xs text-slate-400">
        <Loader2 className="w-3.5 h-3.5 animate-spin" /> กำลังโหลด...
      </span>
    );
    if (syncStatus === "saving") return (
      <span className="flex items-center gap-1 text-xs text-amber-500">
        <Loader2 className="w-3.5 h-3.5 animate-spin" /> กำลังบันทึก...
      </span>
    );
    if (syncStatus === "error") return (
      <span className="flex items-center gap-1 text-xs text-red-500">
        <AlertTriangle className="w-3.5 h-3.5" /> {syncError || "เชื่อมต่อ Sheets ไม่ได้"}
      </span>
    );
    if (syncStatus === "saved" && lastSaved) return (
      <span className="flex items-center gap-1 text-xs text-green-600">
        <CheckCircle className="w-3.5 h-3.5" />
        <CloudUpload className="w-3.5 h-3.5" />
        บันทึกแล้ว {fmtDate(lastSaved)}
      </span>
    );
    if (lastSaved) return (
      <span className="flex items-center gap-1 text-xs text-slate-400">
        <CloudUpload className="w-3.5 h-3.5" />
        อัปเดตล่าสุด {fmtDate(lastSaved)}
      </span>
    );
    return null;
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100">
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

        {/* Content */}
        <div className="overflow-y-auto flex-1 px-6 py-4 space-y-5">
          <p className="text-xs text-slate-500 bg-slate-50 rounded-lg px-3 py-2">
            กำหนดค่า % พิเศษสำหรับ CATEGORY / SUBCATEGORY / DESC_C ที่ต้องการ
            ค่า Default คือ <strong>100%</strong> — รายการที่ตรงตามเงื่อนไขจะใช้ค่า % ที่กำหนดไว้แทน
            (เลือก <strong>ทั้งหมด</strong> เพื่อจับคู่ทุก value ในช่องนั้น · ลำดับบนสุดมีความสำคัญสูงสุด)
          </p>

          {/* Add form */}
          <div className="bg-pink-50/60 rounded-xl border border-pink-100 p-4 space-y-3">
            <p className="text-xs font-semibold text-[#E91E8C]">เพิ่ม Rule ใหม่</p>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <SearchSelect
                label="CATEGORY"
                value={draft.category}
                options={categories}
                onChange={(v) => set("category", v)}
              />
              <SearchSelect
                label="SUBCATEGORY"
                value={draft.subcategory}
                options={subcategories}
                onChange={(v) => set("subcategory", v)}
              />
              <SearchSelect
                label="DESC_C"
                value={draft.descC}
                options={descCList}
                onChange={(v) => set("descC", v)}
              />
              <div className="flex flex-col gap-1">
                <label className="text-xs font-medium text-slate-500">Percentage (%)</label>
                <input
                  type="number"
                  min={1}
                  max={100}
                  step={0.01}
                  value={draft.percentage}
                  onChange={(e) => set("percentage", e.target.value)}
                  placeholder="เช่น 50"
                  className="text-xs border border-slate-200 rounded-lg px-2 py-1.5 bg-white focus:outline-none focus:ring-2 focus:ring-pink-200 focus:border-[#E91E8C]"
                />
              </div>
            </div>

            {/* Status toggle in add form */}
            <div className="flex items-center gap-3">
              <span className="text-xs text-slate-500">Status:</span>
              <button
                type="button"
                onClick={() => set("status", draft.status === "active" ? "inactive" : "active")}
                className={`px-3 py-1 rounded-full text-xs font-bold transition-colors ${
                  draft.status === "active"
                    ? "bg-green-100 text-green-700 hover:bg-green-200"
                    : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                }`}
              >
                {draft.status === "active" ? "Active" : "Inactive"}
              </button>
            </div>

            {error && <p className="text-xs text-red-500">{error}</p>}
            <button
              onClick={addEntry}
              disabled={syncStatus === "saving"}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-[#E91E8C] text-white text-xs font-semibold rounded-lg hover:bg-[#c4187a] disabled:opacity-50 transition-colors"
            >
              <Plus className="w-3.5 h-3.5" />
              เพิ่ม Rule
            </button>
          </div>

          {/* Existing entries */}
          {syncStatus === "loading" ? (
            <div className="flex justify-center py-8">
              <Loader2 className="w-6 h-6 animate-spin text-pink-300" />
            </div>
          ) : config.length === 0 ? (
            <p className="text-xs text-slate-400 text-center py-4">ยังไม่มี Rule — ทุกแถวจะใช้ค่า 100% เป็น Default</p>
          ) : (
            <div className="overflow-x-auto rounded-xl border border-slate-100">
              <table className="min-w-full text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-600">
                    <th className="px-3 py-2 text-left font-semibold">#</th>
                    <th className="px-3 py-2 text-left font-semibold">CATEGORY</th>
                    <th className="px-3 py-2 text-left font-semibold">SUBCATEGORY</th>
                    <th className="px-3 py-2 text-left font-semibold">DESC_C</th>
                    <th className="px-3 py-2 text-center font-semibold">%</th>
                    <th className="px-3 py-2 text-center font-semibold">Status</th>
                    <th className="px-3 py-2 text-left font-semibold whitespace-nowrap">สร้างเมื่อ</th>
                    <th className="px-3 py-2 text-center font-semibold">ลบ</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {config.map((entry, i) => (
                    <tr key={entry.id} className={`hover:bg-slate-50 transition-colors ${entry.status === "inactive" ? "opacity-50" : ""}`}>
                      <td className="px-3 py-2 text-slate-400">{i + 1}</td>
                      <td className="px-3 py-2 max-w-[140px] truncate" title={entry.category}>
                        {entry.category === ALL ? <span className="text-slate-400 italic">{ALL}</span> : entry.category}
                      </td>
                      <td className="px-3 py-2 max-w-[180px] truncate" title={entry.subcategory}>
                        {entry.subcategory === ALL ? <span className="text-slate-400 italic">{ALL}</span> : entry.subcategory}
                      </td>
                      <td className="px-3 py-2 max-w-[140px] truncate" title={entry.descC}>
                        {entry.descC === ALL ? <span className="text-slate-400 italic">{ALL}</span> : entry.descC}
                      </td>
                      <td className="px-3 py-2 text-center font-semibold text-[#E91E8C]">
                        {entry.percentage}%
                      </td>
                      <td className="px-3 py-2 text-center">
                        <button
                          onClick={() => toggleStatus(entry.id)}
                          disabled={syncStatus === "saving"}
                          className={`px-2 py-0.5 rounded-full text-[10px] font-bold transition-colors disabled:opacity-40 ${
                            entry.status === "active"
                              ? "bg-green-100 text-green-700 hover:bg-green-200"
                              : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                          }`}
                        >
                          {entry.status === "active" ? "Active" : "Inactive"}
                        </button>
                      </td>
                      <td className="px-3 py-2 text-slate-400 whitespace-nowrap">
                        {fmtDate(entry.createdAt)}
                      </td>
                      <td className="px-3 py-2 text-center">
                        <button
                          onClick={() => remove(entry.id)}
                          disabled={syncStatus === "saving"}
                          className="p-1 rounded-lg hover:bg-red-50 text-red-400 hover:text-red-600 disabled:opacity-40 transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-slate-100 flex items-center justify-between">
          <span className="text-[10px] text-slate-400">บันทึกลง Google Sheets โดยอัตโนมัติ</span>
          <button
            onClick={onClose}
            className="px-4 py-2 bg-pink-50 text-[#E91E8C] text-sm font-semibold rounded-xl hover:bg-pink-100 transition-colors"
          >
            ปิด
          </button>
        </div>
      </div>
    </div>
  );
}
