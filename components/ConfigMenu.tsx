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

// ── Conflict detection ────────────────────────────────────────────────────────
// Returns true if ruleA "covers" fieldB — i.e. ruleA's field matches ruleB's specific value
function fieldCovers(ruleAVal: string, ruleBVal: string): boolean {
  return ruleAVal === ALL || ruleAVal === ruleBVal;
}

// Returns true if ruleA covers ALL products that ruleB covers (A is a superset of B)
function ruleCovers(a: DraftFields | ExceptionConfig, b: DraftFields | ExceptionConfig): boolean {
  return fieldCovers(a.category, b.category) &&
         fieldCovers(a.subcategory, b.subcategory) &&
         fieldCovers(a.descC, b.descC);
}

// Returns true if a and b have exactly the same (cat, sub, desc) key
function sameKey(a: DraftFields | ExceptionConfig, b: DraftFields | ExceptionConfig): boolean {
  return a.category === b.category && a.subcategory === b.subcategory && a.descC === b.descC;
}

type ConflictKind =
  | "exact-key-active"      // same key, existing active → block
  | "exact-key-inactive"    // same key, existing inactive → block + hint
  | "subset-same-pct"       // new ⊂ existing active, same % → redundant, block
  | "superset-same-pct"     // new ⊃ existing active, same % → existing dead, block
  | "superset-diff-pct";    // new ⊃ existing active, diff % → warn only, allow

interface ConflictResult {
  kind: ConflictKind;
  entries: ExceptionConfig[];
}

function detectConflict(draft: DraftFields, config: ExceptionConfig[], excludeId?: string): ConflictResult | null {
  const others = config.filter((e) => e.id !== excludeId);

  // Check 1: exact key match (any status)
  const exactMatches = others.filter((e) => sameKey(draft, e));
  if (exactMatches.length > 0) {
    const hasActive = exactMatches.some((e) => e.status === "active");
    return {
      kind: hasActive ? "exact-key-active" : "exact-key-inactive",
      entries: exactMatches,
    };
  }

  // Remaining checks only against Active rules
  const activeOthers = others.filter((e) => e.status === "active");
  const samePct = (e: ExceptionConfig) => e.percentage === draft.percentage;

  // Check 2: new is SUBSET of existing active (existing covers new)
  // Block when same %, allow when different %
  const subsetOfActive = activeOthers.filter((e) => !sameKey(draft, e) && ruleCovers(e, draft));
  const subsetSamePct = subsetOfActive.filter(samePct);
  if (subsetSamePct.length > 0) {
    return { kind: "subset-same-pct", entries: subsetSamePct };
  }
  // subset with different % → no conflict (intentional specific override, allowed)

  // Check 3: new is SUPERSET of existing active (new covers existing)
  const supersetOfActive = activeOthers.filter((e) => !sameKey(draft, e) && ruleCovers(draft, e));
  const supersetSamePct = supersetOfActive.filter(samePct);
  if (supersetSamePct.length > 0) {
    return { kind: "superset-same-pct", entries: supersetSamePct };
  }
  const supersetDiffPct = supersetOfActive.filter((e) => !samePct(e));
  if (supersetDiffPct.length > 0) {
    return { kind: "superset-diff-pct", entries: supersetDiffPct };
  }

  return null;
}

function conflictMessage(kind: ConflictKind): string {
  switch (kind) {
    case "exact-key-active":
      return "มีรายการ (Active) ที่ใช้ CATEGORY / SUBCATEGORY / DESC_C เดียวกันอยู่แล้ว — กรุณาแก้ไขรายการเดิมแทน";
    case "exact-key-inactive":
      return "มีรายการนี้อยู่แล้วแต่ถูกตั้งเป็น Inactive — กรุณา Activate หรือแก้ไขรายการเดิมแทนการเพิ่มซ้ำ";
    case "subset-same-pct":
      return "Rule นี้อยู่ใน scope ของ Rule ที่มีอยู่แล้ว และมี % เดียวกัน — เพิ่มแล้วจะไม่มีผลใดๆ (ซ้ำซ้อน)";
    case "superset-same-pct":
      return "Rule นี้ครอบคลุม Rule ที่มีอยู่แล้ว และมี % เดียวกัน — Rule เดิมจะกลายเป็น Dead Rule (ซ้ำซ้อน)";
    case "superset-diff-pct":
      return "Rule นี้ครอบคลุม Rule ที่มีอยู่แล้วซึ่งมี % ต่างกัน — ลำดับ Rule ในตารางจะส่งผลต่อการคำนวณ";
  }
}

function isBlockingConflict(kind: ConflictKind): boolean {
  return kind !== "superset-diff-pct";
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
            <Loader2 className="w-3 h-3 animate-spin flex-shrink-0" /> กำลังโหลด...
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
            <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="ค้นหา..." className="text-xs flex-1 outline-none bg-transparent placeholder-slate-300" />
          </div>
          <div className="overflow-y-auto max-h-52">
            {filtered.length === 0
              ? <p className="text-xs text-slate-400 px-3 py-2">ไม่พบข้อมูล</p>
              : filtered.map((o) => (
                <button key={o} type="button" onClick={() => select(o)}
                  className={`w-full text-left text-xs px-3 py-1.5 hover:bg-pink-50 transition-colors ${
                    o === value ? "bg-pink-50 text-[#E91E8C] font-semibold" : o === ALL ? "text-slate-400 italic" : "text-slate-700"
                  }`}
                >{o}</button>
              ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Conflict card ─────────────────────────────────────────────────────────────
function ConflictCard({ conflict }: { conflict: ConflictResult }) {
  const blocking = isBlockingConflict(conflict.kind);
  return (
    <div className={`rounded-xl border-l-4 p-4 space-y-1.5 text-xs ${
      blocking
        ? "bg-red-200 border-l-red-600 border border-red-300"
        : "bg-amber-200 border-l-amber-500 border border-amber-300"
    }`}>
      <div className={`flex items-start gap-2 ${blocking ? "text-red-900" : "text-amber-900"}`}>
        <AlertTriangle className="w-4 h-4 flex-shrink-0 mt-0.5" />
        <div>
          <p className="font-bold text-sm leading-snug">{blocking ? "ไม่สามารถเพิ่ม Rule นี้ได้" : "คำเตือน"}</p>
          <p className="mt-0.5 leading-relaxed">{conflictMessage(conflict.kind)}</p>
        </div>
      </div>
      <p className={`text-[11px] pl-6 ${blocking ? "text-red-700" : "text-amber-700"}`}>↓ รายการที่ขัดแย้งถูกไฮไลต์ในตารางด้านล่างแล้ว</p>
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

  // Live conflict detection (computed every render when draft is complete)
  const draftComplete = Boolean(draft.category && draft.subcategory && draft.descC && draft.percentage);
  const conflict = draftComplete
    ? detectConflict(draft, config, isEditing ? editId! : undefined)
    : null;

  // ── Computed table data ──
  // When conflict active → override table to show only conflicting entries (highlighted)
  const conflictIds = conflict ? new Set(conflict.entries.map((e) => e.id)) : null;

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

  // Conflict view overrides normal pageData
  const isConflictView = conflictIds !== null;
  const tableRows = isConflictView
    ? config.filter((e) => conflictIds!.has(e.id))
    : pageData;

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
    setEditId(null);
    setDraft({ category: entry.category, subcategory: entry.subcategory, descC: entry.descC, percentage: entry.percentage, status: entry.status });
    setFormError("");
    document.getElementById("config-form")?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const submitForm = () => {
    if (!draft.category || !draft.subcategory || !draft.descC) {
      setFormError("กรุณาเลือกค่าให้ครบทุกช่อง (เลือก 'ทั้งหมด' ถ้าต้องการจับคู่ทุก value)");
      return;
    }
    const pct = parseFloat(draft.percentage);
    if (isNaN(pct) || pct <= 0 || pct > 100) { setFormError("Percentage ต้องเป็นตัวเลข 1–100"); return; }
    setFormError("");
    // Blocking conflicts are shown live — do not allow submit if blocking
    if (conflict && isBlockingConflict(conflict.kind)) return;

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

  const canSubmit = !conflict || !isBlockingConflict(conflict.kind);

  // ── Render ──
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
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

          {/* Info banner — compact */}
          <div className="flex items-center gap-2 flex-wrap text-xs text-slate-500 bg-slate-50 border border-slate-100 rounded-xl px-4 py-2.5">
            <span className="font-semibold text-slate-700 whitespace-nowrap">📋 Rule ถูกอ่านอย่างไร?</span>
            <span className="text-slate-300">|</span>
            <span>CATEGORY + SUBCATEGORY + DESC_C ต้องตรงกันครบทั้ง 3 ช่อง</span>
            <span className="text-slate-300">·</span>
            <span><strong className="text-[#E91E8C]">ทั้งหมด</strong> = จับคู่ทุก value ในช่องนั้น</span>
            <span className="text-slate-300">·</span>
            <span>ไม่มี Rule ตรง → ใช้ <strong>100%</strong> เป็น Default</span>
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
              <SearchSelect label="DESC_C" value={draft.descC} options={descCList} onChange={(v) => set("descC", v)} loading={!spacemanLoaded} />
              <SearchSelect label="CATEGORY" value={draft.category} options={categories} onChange={(v) => set("category", v)} loading={!spacemanLoaded} />
              <SearchSelect label="SUBCATEGORY" value={draft.subcategory} options={subcategories} onChange={(v) => set("subcategory", v)} loading={!spacemanLoaded} />
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

            {/* Validation error */}
            {formError && <p className="text-xs text-red-500">{formError}</p>}

            {/* Live conflict card */}
            {conflict && <ConflictCard conflict={conflict} />}

            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={submitForm}
                disabled={syncStatus === "saving" || !canSubmit}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-white text-xs font-semibold rounded-lg disabled:opacity-40 disabled:cursor-not-allowed transition-colors ${
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
              <button
                onClick={() => { setEditId(null); setDraft(emptyDraft()); setFormError(""); }}
                className="flex items-center gap-1 px-3 py-1.5 text-slate-400 text-xs font-semibold rounded-lg hover:bg-slate-100 hover:text-slate-600 transition-colors"
                title="ล้างค่าในฟอร์มทั้งหมด"
              >
                <X className="w-3 h-3" /> ล้างค่า
              </button>
              {conflict && !isBlockingConflict(conflict.kind) && (
                <span className="text-[11px] text-amber-700 font-medium">⚠️ เพิ่มได้ แต่ควรตรวจสอบลำดับ Rule ในตาราง</span>
              )}
            </div>
          </div>

          {/* Table */}
          {syncStatus === "loading" ? (
            <div className="flex justify-center py-8"><Loader2 className="w-6 h-6 animate-spin text-pink-300" /></div>
          ) : config.length === 0 ? (
            <p className="text-xs text-slate-400 text-center py-4">ยังไม่มี Rule — ทุกแถวจะใช้ค่า 100% เป็น Default</p>
          ) : (
            <div className={`rounded-xl flex flex-col overflow-hidden border ${isConflictView ? "border-red-200" : "border-slate-100"}`}>
              {/* Conflict view indicator */}
              {isConflictView && (
                <div className={`flex items-center gap-2 px-4 py-2 text-xs font-medium flex-shrink-0 ${
                  isBlockingConflict(conflict!.kind) ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700"
                }`}>
                  <AlertTriangle className="w-3.5 h-3.5 flex-shrink-0" />
                  <span>แสดงเฉพาะรายการที่ขัดแย้ง ({tableRows.length} รายการ) — แก้ไขฟอร์มด้านบนเพื่อกลับสู่มุมมองปกติ</span>
                </div>
              )}
              <div className="overflow-x-auto">
                <table className="min-w-full text-xs">
                  <thead>
                    <tr className="bg-slate-50 text-slate-600 border-b border-slate-100">
                      <th className="px-3 py-2 text-left font-semibold w-10">#</th>
                      {(["descC","category","subcategory","percentage","status","createdAt","updatedAt"] as const).map((col) => {
                        const labels: Record<string, string> = {
                          descC: "DESC_C", category: "CATEGORY", subcategory: "SUBCATEGORY",
                          percentage: "%", status: "Status", createdAt: "สร้างเมื่อ", updatedAt: "อัปเดตล่าสุด",
                        };
                        return (
                          <th key={col} onClick={() => toggleSort(col)}
                            className={`px-3 py-2 text-left font-semibold cursor-pointer select-none hover:bg-slate-100 whitespace-nowrap transition-colors ${col === "percentage" || col === "status" ? "text-center" : ""}`}
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
                      <td className="px-2 py-1.5"><input value={colFilters.descC} onChange={(e) => setFilter("descC", e.target.value)} placeholder="กรอง…" className="w-full text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" /></td>
                      <td className="px-2 py-1.5"><input value={colFilters.category} onChange={(e) => setFilter("category", e.target.value)} placeholder="กรอง…" className="w-full text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" /></td>
                      <td className="px-2 py-1.5"><input value={colFilters.subcategory} onChange={(e) => setFilter("subcategory", e.target.value)} placeholder="กรอง…" className="w-full text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" /></td>
                      <td className="px-2 py-1.5 text-center"><input value={colFilters.percentage} onChange={(e) => setFilter("percentage", e.target.value)} placeholder="กรอง…" className="w-16 text-center text-[11px] border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:border-pink-300 placeholder-slate-300" /></td>
                      <td className="px-2 py-1.5 text-center">
                        <select value={colFilters.status} onChange={(e) => setFilter("status", e.target.value)} className="text-[11px] border border-slate-200 rounded px-1 py-0.5 focus:outline-none focus:border-pink-300 bg-white text-slate-600">
                          <option value="">ทั้งหมด</option>
                          <option value="active">Active</option>
                          <option value="inactive">Inactive</option>
                        </select>
                      </td>
                      <td className="px-3 py-1.5" /><td className="px-3 py-1.5" />
                      <td className="px-3 py-1.5 text-center">
                        {hasFilter && (
                          <button onClick={() => { setColFilters(emptyFilters()); setPage(0); }} className="text-[11px] text-slate-400 hover:text-red-400 transition-colors whitespace-nowrap">ล้าง</button>
                        )}
                      </td>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {tableRows.length === 0 ? (
                      <tr><td colSpan={9} className="px-3 py-6 text-center text-xs text-slate-400">ไม่พบรายการที่ตรงกับ Filter</td></tr>
                    ) : tableRows.map((entry, i) => {
                      const isThisEdit = editId === entry.id;
                      const isConflicting = conflictIds?.has(entry.id) ?? false;
                      const globalIdx = isConflictView ? (i + 1) : (safePage * PAGE_SIZE + i + 1);
                      return (
                        <tr key={entry.id} className={`transition-colors ${
                          isConflicting
                            ? isBlockingConflict(conflict!.kind)
                              ? "bg-red-50 ring-1 ring-inset ring-red-200"
                              : "bg-amber-50 ring-1 ring-inset ring-amber-200"
                          : isThisEdit ? "bg-amber-50 ring-1 ring-inset ring-amber-200"
                            : entry.status === "inactive" ? "opacity-50 hover:opacity-70"
                            : "hover:bg-slate-50"
                        }`}>
                          <td className="px-3 py-2.5 text-slate-400">{globalIdx}</td>
                          <td className="px-3 py-2.5 max-w-[160px]" title={entry.descC}>
                            <span className={`truncate block ${entry.descC === ALL ? "text-slate-400 italic" : ""}`}>{entry.descC}</span>
                          </td>
                          <td className="px-3 py-2.5 max-w-[160px]" title={entry.category}>
                            <span className={`truncate block ${entry.category === ALL ? "text-slate-400 italic" : ""}`}>{entry.category}</span>
                          </td>
                          <td className="px-3 py-2.5 max-w-[200px]" title={entry.subcategory}>
                            <span className={`truncate block ${entry.subcategory === ALL ? "text-slate-400 italic" : ""}`}>{entry.subcategory}</span>
                          </td>
                          <td className="px-3 py-2.5 text-center font-semibold text-[#E91E8C] whitespace-nowrap">{entry.percentage}%</td>
                          <td className="px-3 py-2.5 text-center">
                            <button onClick={() => toggleStatus(entry.id)} disabled={syncStatus === "saving"}
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
                              <button onClick={() => startEdit(entry)} title="แก้ไข" disabled={syncStatus === "saving"}
                                className={`p-1.5 rounded-lg transition-colors disabled:opacity-40 ${
                                  isThisEdit ? "bg-amber-100 text-amber-600" : "hover:bg-amber-50 text-slate-400 hover:text-amber-600"
                                }`}
                              ><Pencil className="w-3.5 h-3.5" /></button>
                              <button onClick={() => copyEntry(entry)} title="คัดลอก → กรอก form" disabled={syncStatus === "saving"}
                                className="p-1.5 rounded-lg hover:bg-blue-50 text-slate-400 hover:text-blue-500 disabled:opacity-40 transition-colors"
                              ><Copy className="w-3.5 h-3.5" /></button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {/* Pagination — hidden in conflict view */}
              {!isConflictView && (
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
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-slate-100 flex justify-end flex-shrink-0">
          <button onClick={onClose} className="px-4 py-2 bg-pink-50 text-[#E91E8C] text-sm font-semibold rounded-xl hover:bg-pink-100 transition-colors">
            ปิด
          </button>
        </div>
      </div>
    </div>
  );
}
