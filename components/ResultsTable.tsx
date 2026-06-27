"use client";
import { useState, useMemo } from "react";
import { CheckCircle, AlertTriangle, XCircle, Pencil, Save, X } from "lucide-react";
import type { ProcessedRow, FilledData } from "@/lib/types";

interface Props {
  rows: ProcessedRow[];
  onChange: (updated: ProcessedRow[]) => void;
  /** Pre-populated unique values from the full RECAP file for each column */
  externalSuggestions?: Partial<Record<string, string[]>>;
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
    label: "อนุมาน",
    bg: "bg-amber-50",
    badge: "bg-amber-100 text-amber-700",
  },
  not_found: {
    icon: <XCircle className="w-4 h-4 text-red-400" />,
    label: "ไม่พบ",
    bg: "bg-red-50",
    badge: "bg-red-100 text-red-700",
  },
};

const FIELDS: { key: keyof FilledData; col: string }[] = [
  { key: "division",  col: "F — DIVISION"  },
  { key: "dept",      col: "G — DEPT"      },
  { key: "subDept",   col: "H — SUB-DEPT"  },
  { key: "cls",       col: "I — Class"     },
  { key: "planogram", col: "J — PLANOGRAM" },
  { key: "colN",      col: "N — MBC Forecast sale"              },
  { key: "colO",      col: "O — Shelf stock ON POG (Piece) 100%" },
];

export default function ResultsTable({ rows, onChange, externalSuggestions }: Props) {
  const [editIdx, setEditIdx] = useState<number | null>(null);
  const [draft, setDraft] = useState<Partial<FilledData>>({});

  // Merge suggestions: external (full RECAP file) + values from current rows
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
    confirmed: rows.filter((r) => r.confidence === "confirmed").length,
    inferred:  rows.filter((r) => r.confidence === "inferred").length,
    not_found: rows.filter((r) => r.confidence === "not_found").length,
  };

  return (
    <div className="space-y-6">
      {/* Datalists for autocomplete */}
      {FIELDS.map(({ key }) => (
        <datalist key={key} id={`dl-${key}`}>
          {suggestions[key].map((v) => (
            <option key={v} value={v} />
          ))}
        </datalist>
      ))}

      {/* Summary pills */}
      <div className="flex flex-wrap gap-3">
        {(["confirmed", "inferred", "not_found"] as const).map((c) => (
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
              <th className="px-3 py-3 text-left font-semibold whitespace-nowrap">สถานะ</th>
              <th className="px-3 py-3 text-left font-semibold whitespace-nowrap">Barcode</th>
              <th className="px-3 py-3 text-left font-semibold whitespace-nowrap">ชื่อสินค้า</th>
              {FIELDS.map(({ col }) => (
                <th key={col} className="px-3 py-3 text-left font-semibold whitespace-nowrap">
                  {col}
                </th>
              ))}
              <th className="px-3 py-3 text-center font-semibold whitespace-nowrap">แก้ไข</th>
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
                  <td className="px-3 py-3 whitespace-nowrap">
                    <div className={`inline-flex items-center gap-1.5 px-2 py-1 rounded-full text-xs font-medium ${meta.badge}`}>
                      {meta.icon}
                      {meta.label}
                    </div>
                  </td>
                  <td className="px-3 py-3 font-mono text-xs text-slate-600 whitespace-nowrap">
                    {row.barcode}
                  </td>
                  <td className="px-3 py-3 max-w-[200px] truncate text-slate-700" title={row.name}>
                    {row.name}
                  </td>

                  {FIELDS.map(({ key }) => (
                    <td key={key} className="px-3 py-3 min-w-[160px]">
                      {isEditing ? (
                        <input
                          list={`dl-${key}`}
                          value={draft[key] ?? ""}
                          onChange={(e) => setDraft({ ...draft, [key]: e.target.value })}
                          placeholder="พิมพ์หรือเลือก..."
                          className="w-full px-2 py-1.5 text-xs border border-[#E91E8C] rounded-lg focus:outline-none focus:ring-2 focus:ring-pink-200 bg-white"
                        />
                      ) : (
                        <span className="text-slate-700 text-xs">
                          {(data as Record<string, string>)[key] || (
                            <span className="text-slate-300 italic">—</span>
                          )}
                        </span>
                      )}
                    </td>
                  ))}

                  <td className="px-3 py-3 text-center whitespace-nowrap">
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

      {summary.not_found > 0 && (
        <p className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-4 py-3">
          ⚠️ มี {summary.not_found} รายการที่ไม่พบในไฟล์ 100 ช่อง —
          กรุณาคลิกปุ่มแก้ไข (✏️) เพื่อกรอกข้อมูลด้วยตนเอง
        </p>
      )}
    </div>
  );
}
