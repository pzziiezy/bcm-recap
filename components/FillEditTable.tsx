"use client";
import { useState } from "react";
import { Pencil, Check, X, ArrowLeftRight } from "lucide-react";
import type { FillRow } from "@/lib/download";

// ─── Types ──────────────────────────────────────────────────────────────────

export interface TabColDef {
  field: string;
  col: number;
  label: string;
  editable: boolean;
  /** Field name this column's options should be filtered by (cascade parent) */
  cascade?: string;
  /** Column group for visual zone coloring, e.g. "new" | "del" */
  zone?: string;
}

export interface EditableFillRow {
  rowIndex: number;
  fields: Record<string, string>;
}

// ─── Conversion utilities ────────────────────────────────────────────────────

export function convertToEditableRows(rows: FillRow[], colDefs: TabColDef[]): EditableFillRow[] {
  return rows.map(r => ({
    rowIndex: r.rowIndex,
    fields: Object.fromEntries(
      colDefs.map(({ field, col }) => [field, r.cells.find(c => c.col === col)?.value ?? ""])
    ),
  }));
}

/**
 * Convert edited rows back to FillRow[], preserving any cells outside the colDefs
 * (e.g., store-flag columns in NEW SCM / DEL SCM that aren't shown in the preview).
 */
export function convertFromEditableRows(
  editableRows: EditableFillRow[],
  originalRows: FillRow[],
  colDefs: TabColDef[]
): FillRow[] {
  const previewCols = new Set(colDefs.map(d => d.col));
  return editableRows.map(er => {
    const origRow = originalRows.find(r => r.rowIndex === er.rowIndex);
    const origExtra = origRow ? origRow.cells.filter(c => !previewCols.has(c.col)) : [];
    const editedCells = colDefs
      .map(({ field, col }) => ({ col, value: er.fields[field] ?? "" }))
      .filter(c => c.value !== "");
    return { rowIndex: er.rowIndex, cells: [...editedCells, ...origExtra] };
  });
}

// ─── Zone style maps ─────────────────────────────────────────────────────────

const ZONE_TH: Record<string, string> = {
  new: "bg-emerald-50 text-emerald-700",
  del: "bg-rose-50 text-rose-700",
};
const ZONE_TD: Record<string, string> = {
  new: "bg-emerald-50/40",
  del: "bg-rose-50/40",
};

// ─── Component ───────────────────────────────────────────────────────────────

interface Props {
  colDefs: TabColDef[];
  rows: EditableFillRow[];
  onChange: (updated: EditableFillRow[]) => void;
  /** Returns dropdown options for a field given the current draft values (for cascading). */
  getOptions: (field: string, draft: Record<string, string>) => string[];
  /**
   * Per-zone key highlight: return true if cells in `zone` should be highlighted for this row.
   * Enables independent highlighting per zone (e.g. seqNew drives "new" zone, seqDel drives "del" zone).
   */
  isKeyZone?: (row: EditableFillRow, zone: string) => boolean;
}

type MatchMode = "exact" | "contains";

export default function FillEditTable({ colDefs, rows, onChange, getOptions, isKeyZone }: Props) {
  // ── row edit state ──
  const [editIdx, setEditIdx] = useState<number | null>(null);
  const [draft, setDraft] = useState<Record<string, string>>({});

  // ── find & replace state ──
  const [showReplace, setShowReplace]   = useState(false);
  const [replaceCol, setReplaceCol]     = useState("");
  const [findVal, setFindVal]           = useState("");
  const [replaceVal, setReplaceVal]     = useState("");
  const [matchMode, setMatchMode]       = useState<MatchMode>("exact");
  const [replaceMsg, setReplaceMsg]     = useState<string | null>(null);

  const startEdit = (i: number) => {
    setDraft({ ...rows[i].fields });
    setEditIdx(i);
  };

  const saveEdit = (i: number) => {
    onChange(rows.map((r, idx) => (idx === i ? { ...r, fields: { ...draft } } : r)));
    setEditIdx(null);
  };

  const cancelEdit = () => setEditIdx(null);

  // ── editable columns (for replace column selector) ──
  const editableCols = colDefs.filter(c => c.editable);

  // ── match logic: empty findVal = match empty cells only ──
  const isMatch = (cur: string) => {
    if (findVal === "") return cur === "";
    return matchMode === "exact" ? cur === findVal : cur.includes(findVal);
  };

  const matchCount = replaceCol
    ? rows.filter(r => isMatch(r.fields[replaceCol] ?? "")).length
    : 0;

  const applyReplace = () => {
    if (!replaceCol || matchCount === 0) return;
    const updated = rows.map(r => {
      if (!isMatch(r.fields[replaceCol] ?? "")) return r;
      return { ...r, fields: { ...r.fields, [replaceCol]: replaceVal } };
    });
    onChange(updated);
    setReplaceMsg(`แทนที่ ${matchCount} แถวเรียบร้อย`);
    setTimeout(() => setReplaceMsg(null), 3000);
  };

  const toggleReplace = () => {
    setShowReplace(v => !v);
    setReplaceMsg(null);
  };

  return (
    <div className="rounded border border-slate-200 overflow-hidden">

      {/* ── Find & Replace toolbar toggle ──────────────────────────────────── */}
      <div className="flex items-center justify-between px-2 py-1.5 bg-slate-50 border-b border-slate-200">
        <span className="text-[10px] text-slate-400 select-none">{rows.length} แถว</span>
        <button
          onClick={toggleReplace}
          title="Find & Replace"
          className={`flex items-center gap-1 text-[11px] font-medium px-2.5 py-1 rounded transition-colors ${
            showReplace
              ? "bg-[#E91E8C] text-white"
              : "text-slate-500 hover:bg-slate-200 hover:text-slate-700"
          }`}
        >
          <ArrowLeftRight className="w-3 h-3" />
          Replace
        </button>
      </div>

      {/* ── Find & Replace panel ────────────────────────────────────────────── */}
      {showReplace && (
        <div className="border-b border-amber-200 bg-amber-50/70 px-4 py-3 space-y-2.5">
          {/* Row 1: column selector */}
          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-[11px] font-semibold text-slate-600 w-16 shrink-0">คอลัมน์</span>
            <select
              value={replaceCol}
              onChange={e => {
                setReplaceCol(e.target.value);
                setFindVal("");
                setReplaceVal("");
                setReplaceMsg(null);
              }}
              className="border border-slate-300 rounded px-2 py-1 text-xs bg-white focus:outline-none focus:border-[#E91E8C]"
            >
              <option value="">— เลือกคอลัมน์ที่ต้องการแทนที่ —</option>
              {editableCols.map(c => (
                <option key={c.field} value={c.field}>{c.label}</option>
              ))}
            </select>
          </div>

          {replaceCol && (
            <>
              {/* Row 2: find + replace inputs */}
              <div className="flex flex-wrap gap-3 items-center">
                <div className="flex items-center gap-2">
                  <span className="text-[11px] font-semibold text-slate-600 w-16 shrink-0">ค้นหา</span>
                  <input
                    list="filler-dl-find"
                    value={findVal}
                    onChange={e => { setFindVal(e.target.value); setReplaceMsg(null); }}
                    placeholder="ว่าง = ค้นหาช่องที่ยังไม่มีค่า"
                    className="border border-slate-300 rounded px-2 py-1 text-xs bg-white focus:outline-none focus:border-[#E91E8C] min-w-[180px]"
                  />
                  <datalist id="filler-dl-find">
                    {[...new Set(rows.map(r => r.fields[replaceCol]).filter(Boolean))].map(v => (
                      <option key={v} value={v} />
                    ))}
                  </datalist>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-[11px] font-semibold text-slate-600 w-16 shrink-0">แทนที่</span>
                  <input
                    list="filler-dl-replace"
                    value={replaceVal}
                    onChange={e => { setReplaceVal(e.target.value); setReplaceMsg(null); }}
                    placeholder="ว่าง = ลบค่าออก"
                    className="border border-slate-300 rounded px-2 py-1 text-xs bg-white focus:outline-none focus:border-[#E91E8C] min-w-[180px]"
                  />
                  <datalist id="filler-dl-replace">
                    {getOptions(replaceCol, {}).map(v => (
                      <option key={v} value={v} />
                    ))}
                  </datalist>
                </div>
              </div>

              {/* Row 3: match mode + action */}
              <div className="flex flex-wrap gap-4 items-center">
                <div className="flex items-center gap-3 text-xs text-slate-600">
                  <span className="font-semibold">จับคู่แบบ</span>
                  <label className="flex items-center gap-1 cursor-pointer select-none">
                    <input
                      type="radio"
                      name="matchMode"
                      value="exact"
                      checked={matchMode === "exact"}
                      onChange={() => setMatchMode("exact")}
                      className="accent-[#E91E8C]"
                    />
                    ตรงทั้งหมด
                  </label>
                  <label className="flex items-center gap-1 cursor-pointer select-none">
                    <input
                      type="radio"
                      name="matchMode"
                      value="contains"
                      checked={matchMode === "contains"}
                      onChange={() => setMatchMode("contains")}
                      className="accent-[#E91E8C]"
                    />
                    มีคำนี้
                  </label>
                </div>

                <div className="ml-auto flex items-center gap-3">
                  {replaceMsg ? (
                    <span className="text-[11px] text-green-600 font-semibold">{replaceMsg}</span>
                  ) : (
                    <span className={`text-[11px] font-semibold ${matchCount > 0 ? "text-amber-600" : "text-slate-400"}`}>
                      {matchCount > 0 ? `พบ ${matchCount} แถว` : "ไม่พบแถวที่ตรง"}
                    </span>
                  )}
                  <button
                    onClick={applyReplace}
                    disabled={matchCount === 0}
                    className="px-3 py-1.5 rounded text-[11px] font-semibold bg-[#E91E8C] text-white
                               hover:bg-[#d01879] disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                  >
                    แทนที่ทั้งหมด ({matchCount})
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* ── Table ────────────────────────────────────────────────────────────── */}
      <div className="overflow-x-auto">
        <table className="text-xs w-full border-collapse">
          <thead>
            <tr>
              <th className="w-14 px-2 py-1.5 bg-slate-50 border-b border-slate-200" />
              {colDefs.map(({ field, label, editable, zone }, i) => {
                const isZoneStart = zone !== undefined && zone !== colDefs[i - 1]?.zone;
                return (
                  <th
                    key={field}
                    className={[
                      "px-2 py-1.5 text-left font-semibold border-b border-slate-200 whitespace-nowrap",
                      zone ? ZONE_TH[zone] : (editable ? "text-slate-700 bg-slate-50" : "text-slate-400 bg-slate-50"),
                      isZoneStart ? "border-l-2 border-l-slate-300" : "",
                    ].filter(Boolean).join(" ")}
                  >
                    {label}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => {
              const isEditing = editIdx === i;
              const anyZoneKey = isKeyZone
                ? colDefs.some(cd => cd.zone && isKeyZone(row, cd.zone))
                : false;

              return (
                <tr key={row.rowIndex} className={i % 2 === 0 ? "bg-white" : "bg-slate-50/40"}>
                  {/* Action cell */}
                  <td className={[
                    "px-1 py-1 border-b border-slate-100 text-center align-middle",
                    isEditing ? "border-l-[3px] border-l-[#E91E8C] bg-pink-50" : (anyZoneKey ? "border-l-[3px] border-l-[#E91E8C]" : ""),
                  ].filter(Boolean).join(" ")}>
                    {isEditing ? (
                      <div className="flex gap-0.5 justify-center">
                        <button
                          onClick={() => saveEdit(i)}
                          title="บันทึก"
                          className="p-0.5 rounded text-green-600 hover:bg-green-50"
                        >
                          <Check className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={cancelEdit}
                          title="ยกเลิก"
                          className="p-0.5 rounded text-slate-400 hover:bg-slate-100"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ) : (
                      <button
                        onClick={() => startEdit(i)}
                        title="แก้ไขแถวนี้"
                        className="p-0.5 rounded text-slate-300 hover:text-[#E91E8C] hover:bg-pink-50 transition-colors"
                      >
                        <Pencil className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </td>

                  {/* Data cells */}
                  {colDefs.map(({ field, editable, zone }, ci) => {
                    const isZoneStart = zone !== undefined && zone !== colDefs[ci - 1]?.zone;
                    const zoneIsKey = isKeyZone && zone ? isKeyZone(row, zone) : false;
                    const val = isEditing ? (draft[field] ?? "") : (row.fields[field] ?? "");
                    const dlId = `dl-fill-${i}-${field}`;
                    return (
                      <td
                        key={field}
                        className={[
                          "px-2 py-1 border-b border-slate-100 align-middle",
                          isEditing ? "bg-pink-50" : (zoneIsKey ? "bg-amber-50" : (zone ? ZONE_TD[zone] : "")),
                          isZoneStart ? "border-l-2 border-l-slate-200" : "",
                        ].filter(Boolean).join(" ")}
                      >
                        {isEditing && editable ? (
                          <>
                            <input
                              list={dlId}
                              value={draft[field] ?? ""}
                              onChange={(e) =>
                                setDraft(d => ({ ...d, [field]: e.target.value }))
                              }
                              className="border border-slate-300 rounded px-1.5 py-0.5 text-xs focus:outline-none focus:border-[#E91E8C] min-w-[80px] max-w-[220px] w-full"
                            />
                            <datalist id={dlId}>
                              {getOptions(field, draft).map(o => (
                                <option key={o} value={o} />
                              ))}
                            </datalist>
                          </>
                        ) : (
                          <span
                            className={[
                              "whitespace-nowrap",
                              val ? "text-slate-700" : "text-slate-300",
                              zoneIsKey && val ? "font-medium" : "",
                            ].filter(Boolean).join(" ")}
                          >
                            {val || "—"}
                          </span>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
