"use client";
import { useState, useMemo } from "react";
import { CheckCircle, AlertTriangle, XCircle, Pencil, Save, X, Database } from "lucide-react";
import type { ProcessedRow, FilledData, HierarchyMap } from "@/lib/types";
import { getHierarchyOptions, isHierarchyKey, buildHierarchyFromRows, mergeHierarchies } from "@/lib/hierarchy";

interface Props {
  rows: ProcessedRow[];
  onChange: (updated: ProcessedRow[]) => void;
  /** Pre-populated unique values from the full RECAP file for each column */
  externalSuggestions?: Partial<Record<string, string[]>>;
  /** Parent→child relationship map extracted from the RECAP file for cascading filters */
  hierarchyMap?: HierarchyMap;
}

const CONFIDENCE_META = {
  confirmed: {
    icon: <CheckCircle className="w-4 h-4 text-green-500" />,
    label: "ยืนยันแล้ว",
    bg: "bg-green-50",
    badge: "bg-green-100 text-green-700",
  },
  inferred: {
    icon: <AlertTriangle className="w-4 h-4 text-amber-500" />,
    label: "ไม่มี Planogram",
    bg: "bg-amber-50",
    badge: "bg-amber-100 text-amber-700",
  },
  not_found: {
    icon: <XCircle className="w-4 h-4 text-red-400" />,
    label: "ไม่พบ",
    bg: "bg-red-50",
    badge: "bg-red-100 text-red-700",
  },
  from_spaceman: {
    icon: <Database className="w-4 h-4 text-blue-500" />,
    label: "จาก Spaceman",
    bg: "bg-blue-50",
    badge: "bg-blue-100 text-blue-700",
  },
};

const FIELDS: { key: keyof FilledData; col: string; minW?: string }[] = [
  { key: "division",  col: "F — DIVISION"  },
  { key: "dept",      col: "G — DEPT"      },
  { key: "subDept",   col: "H — SUB-DEPT"  },
  { key: "cls",       col: "I — Class"     },
  { key: "planogram", col: "J — PLANOGRAM" },
  { key: "colN",      col: "N — MBC Forecast sale" },
  { key: "colPiece",  col: "O — Piece 100%",  minW: "min-w-[64px]" },
  { key: "colO",      col: "P — %",            minW: "min-w-[56px]" },
];

/** Compute column Q: colO% × colPiece */
function computeColQ(data: Record<string, string>): string {
  const pctNum   = parseFloat(data.colO     ?? "0") || 0;
  const pieceNum = parseFloat(data.colPiece ?? "0") || 0;
  if (!data.colPiece || !data.colO) return "";
  return (Math.round((pctNum / 100) * pieceNum * 100) / 100).toFixed(2);
}

export default function ResultsTable({ rows, onChange, externalSuggestions, hierarchyMap }: Props) {
  const [editIdx, setEditIdx] = useState<number | null>(null);
  const [draft, setDraft] = useState<Partial<FilledData>>({});

  // Base suggestions: merge external (full RECAP scan) + values already in rows
  const suggestions = useMemo(() => {
    const map = Object.fromEntries(
      FIELDS.map((f) => [f.key, new Set<string>(externalSuggestions?.[f.key] ?? [])])
    ) as Record<keyof FilledData, Set<string>>;

    for (const row of rows) {
      const data = { ...(row.filled ?? {}), ...(row.override ?? {}) } as Record<string, string>;
      for (const { key } of FIELDS) {
        if (data[key]) map[key].add(data[key]);
      }
    }
    return Object.fromEntries(
      FIELDS.map((f) => [f.key, [...map[f.key]].sort()])
    ) as Record<keyof FilledData, string[]>;
  }, [rows, externalSuggestions]);

  // Hierarchy derived from the result rows themselves (covers values the RECAP file didn't have)
  const rowsHierarchy = useMemo(() => buildHierarchyFromRows(rows), [rows]);

  // Effective hierarchy = RECAP file hierarchy ∪ result-rows hierarchy
  // This ensures newly-matched dept/subDept/cls values appear in the cascade dropdowns
  const effectiveHierarchy = useMemo(
    () => hierarchyMap ? mergeHierarchies(hierarchyMap, rowsHierarchy) : rowsHierarchy,
    [hierarchyMap, rowsHierarchy]
  );

  // Cascade-filtered options for hierarchy columns (live while editing)
  const hierarchyOptions = useMemo(() => {
    if (Object.keys(effectiveHierarchy.divToDept).length === 0) return null;
    return getHierarchyOptions(draft, effectiveHierarchy, suggestions);
  }, [draft, effectiveHierarchy, suggestions]);

  // Final datalist options: hierarchy-filtered for F/G/H/I, flat for the rest
  const datalistOptions = useMemo(
    () =>
      Object.fromEntries(
        FIELDS.map(({ key }) => [
          key,
          hierarchyOptions && isHierarchyKey(key)
            ? hierarchyOptions[key]
            : suggestions[key],
        ])
      ) as Record<keyof FilledData, string[]>,
    [hierarchyOptions, suggestions]
  );

  const startEdit = (i: number) => {
    const r = rows[i];
    setDraft({ ...(r.filled ?? {}), ...(r.override ?? {}) });
    setEditIdx(i);
  };

  const saveEdit = (i: number) => {
    const updated = rows.map((r, idx) =>
      idx === i ? { ...r, override: { ...draft } } : r
    );
    onChange(updated);
    setEditIdx(null);
  };

  const cancelEdit = () => setEditIdx(null);

  const summary = {
    confirmed:    rows.filter((r) => r.confidence === "confirmed").length,
    inferred:     rows.filter((r) => r.confidence === "inferred").length,
    not_found:    rows.filter((r) => r.confidence === "not_found").length,
    from_spaceman: rows.filter((r) => r.confidence === "from_spaceman").length,
  };

  return (
    <div className="space-y-6">
      {/* Datalists — hierarchy ones update live as the draft changes */}
      {FIELDS.map(({ key }) => (
        <datalist key={key} id={`dl-${key}`}>
          {datalistOptions[key].map((v) => (
            <option key={v} value={v} />
          ))}
        </datalist>
      ))}

      {/* Summary pills */}
      <div className="flex flex-wrap gap-3">
        {(["confirmed", "inferred", "not_found", "from_spaceman"] as const).map((c) => (
          <div
            key={c}
            className={`flex items-center gap-2 px-4 py-2 rounded-full text-sm font-semibold ${CONFIDENCE_META[c].badge}`}
          >
            {CONFIDENCE_META[c].icon}
            {CONFIDENCE_META[c].label}: {summary[c]} รายการ
          </div>
        ))}
      </div>

      {/* Table */}
      <div className="overflow-x-auto rounded-xl border border-pink-100 shadow-sm">
        <table className="min-w-full text-sm">
          <thead>
            <tr className="bg-gradient-to-r from-pink-50 to-orange-50 text-slate-700">
              <th className="px-2 py-3 text-left font-semibold whitespace-nowrap">สถานะ</th>
              <th className="px-2 py-3 text-left font-semibold whitespace-nowrap">Barcode</th>
              <th className="px-2 py-3 text-left font-semibold whitespace-nowrap">ชื่อสินค้า</th>
              {FIELDS.map(({ col }) => (
                <th key={col} className="px-2 py-3 text-left font-semibold whitespace-nowrap">
                  {col}
                </th>
              ))}
              <th className="px-2 py-3 text-left font-semibold whitespace-nowrap">Q — Net</th>
              <th className="px-2 py-3 text-center font-semibold whitespace-nowrap">แก้ไข</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {rows.map((row, i) => {
              const meta = CONFIDENCE_META[row.confidence];
              const data = { ...(row.filled ?? {}), ...(row.override ?? {}) };
              const isEditing = editIdx === i;

              return (
                <tr
                  key={row.barcode}
                  className={`${meta.bg} hover:brightness-[0.98] transition-colors`}
                >
                  <td className="px-2 py-3 whitespace-nowrap">
                    <div className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${meta.badge}`}>
                      {meta.icon}
                      {meta.label}
                    </div>
                  </td>
                  <td className="px-2 py-3 font-mono text-xs text-slate-600 whitespace-nowrap">
                    {row.barcode}
                  </td>
                  <td className="px-2 py-3 max-w-[140px] truncate text-slate-700 text-xs" title={row.name}>
                    {row.name}
                  </td>

                  {FIELDS.map(({ key, minW }) => (
                    <td key={key} className={`px-2 py-3 ${minW ?? "min-w-[90px]"}`}>
                      {isEditing ? (
                        <input
                          list={`dl-${key}`}
                          value={draft[key] ?? ""}
                          onChange={(e) => setDraft({ ...draft, [key]: e.target.value })}
                          placeholder="พิมพ์หรือเลือก..."
                          className="w-full min-w-[80px] px-2 py-1.5 text-xs border border-[#E91E8C] rounded-lg focus:outline-none focus:ring-2 focus:ring-pink-200 bg-white"
                        />
                      ) : (
                        <span className="text-slate-700 text-xs whitespace-nowrap">
                          {(data as Record<string, string>)[key] || (
                            <span className="text-slate-300 italic">—</span>
                          )}
                        </span>
                      )}
                    </td>
                  ))}

                  {/* Column Q — computed from colO% × colPiece, always read-only */}
                  <td className="px-2 py-3 min-w-[56px]">
                    {(() => {
                      const effectiveData = isEditing
                        ? { ...(data as Record<string, string>), ...draft }
                        : (data as Record<string, string>);
                      const q = computeColQ(effectiveData);
                      return q
                        ? <span className="text-slate-700 text-xs font-mono">{q}</span>
                        : <span className="text-slate-300 italic text-xs">—</span>;
                    })()}
                  </td>

                  <td className="px-2 py-3 text-center whitespace-nowrap">
                    {isEditing ? (
                      <div className="flex items-center justify-center gap-1">
                        <button
                          onClick={() => saveEdit(i)}
                          className="p-1.5 rounded-lg bg-green-100 text-green-600 hover:bg-green-200 transition-colors"
                          title="บันทึก"
                        >
                          <Save className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={cancelEdit}
                          className="p-1.5 rounded-lg bg-slate-100 text-slate-500 hover:bg-slate-200 transition-colors"
                          title="ยกเลิก"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ) : (
                      <button
                        onClick={() => startEdit(i)}
                        className="p-1.5 rounded-lg bg-pink-50 text-[#E91E8C] hover:bg-pink-100 transition-colors"
                        title="แก้ไข"
                      >
                        <Pencil className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {summary.from_spaceman > 0 && (
        <p className="text-sm text-blue-700 bg-blue-50 border border-blue-200 rounded-lg px-4 py-3">
          ℹ️ มี {summary.from_spaceman} รายการที่ไม่พบในไฟล์ 100 ช่อง —
          ข้อมูลคอลัมน์ F/G/H/I นำมาจาก DATA_SPACEMAN (คอลัมน์ N จะเว้นว่างไว้)
        </p>
      )}
      {summary.not_found > 0 && (
        <p className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-4 py-3">
          ⚠️ มี {summary.not_found} รายการที่ไม่พบในไฟล์ 100 ช่อง —
          กรุณาคลิกปุ่มแก้ไข (✏️) เพื่อกรอกข้อมูลด้วยตนเอง
        </p>
      )}
    </div>
  );
}
