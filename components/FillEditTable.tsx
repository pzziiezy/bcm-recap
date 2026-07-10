"use client";
import { useState } from "react";
import { Pencil, Check, X } from "lucide-react";
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
// Cell tints — only applied on non-key rows (key rows let the tr bg show through)
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
  /** When true, row gets a distinct highlight (e.g. rows with a sequence number). */
  isKeyRow?: (row: EditableFillRow) => boolean;
}

export default function FillEditTable({ colDefs, rows, onChange, getOptions, isKeyRow }: Props) {
  const [editIdx, setEditIdx] = useState<number | null>(null);
  const [draft, setDraft] = useState<Record<string, string>>({});

  const startEdit = (i: number) => {
    setDraft({ ...rows[i].fields });
    setEditIdx(i);
  };

  const saveEdit = (i: number) => {
    onChange(rows.map((r, idx) => (idx === i ? { ...r, fields: { ...draft } } : r)));
    setEditIdx(null);
  };

  const cancelEdit = () => setEditIdx(null);

  return (
    <div className="overflow-x-auto rounded border border-slate-200">
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
            const isKey = isKeyRow ? isKeyRow(row) : false;
            // Key rows: tr bg shows through td (no td bg applied below)
            // Non-key rows: alternating white/slate, td gets subtle zone tint
            const trClass = isKey
              ? "bg-amber-50/70"
              : i % 2 === 0 ? "bg-white" : "bg-slate-50/40";

            return (
              <tr key={row.rowIndex} className={trClass}>
                {/* Action cell — left accent border on key rows */}
                <td className={[
                  "px-1 py-1 border-b border-slate-100 text-center align-middle",
                  isKey ? "border-l-[3px] border-l-[#E91E8C]" : "",
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
                  const val = isEditing ? (draft[field] ?? "") : (row.fields[field] ?? "");
                  const dlId = `dl-fill-${i}-${field}`;
                  return (
                    <td
                      key={field}
                      className={[
                        "px-2 py-1 border-b border-slate-100 align-middle",
                        // Zone tint only on non-key rows (key rows use tr bg instead)
                        !isKey && zone ? ZONE_TD[zone] : "",
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
                            isKey && val ? "font-medium" : "",
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
  );
}
