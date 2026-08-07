"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import DropZone from "./DropZone";
import { Database, Search, X } from "lucide-react";

interface FixtureRow {
  SEG: number | string;
  POG: string;
  "Code Fixture": string;
}

type Status = "idle" | "parsing" | "done" | "error";

const PAGE_SIZE = 100;

export default function MasterTab() {
  const [fixtureFile, setFixtureFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [rows, setRows] = useState<FixtureRow[]>([]);
  const [sheetUsed, setSheetUsed] = useState("");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const [errorMsg, setErrorMsg] = useState("");

  const handleParse = async (file: File) => {
    setStatus("parsing");
    setStatusMsg("กำลังอ่านไฟล์...");
    setErrorMsg("");
    setRows([]);
    setSheetUsed("");
    setSearch("");
    setPage(0);

    try {
      await new Promise<void>((r) => setTimeout(r, 0)); // yield to UI
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      // Prefer Fixture_2026, fallback to any sheet starting with "Fixture", then first sheet
      const targetSheet =
        wb.SheetNames.find((n) => n === "Fixture_2026") ??
        wb.SheetNames.find((n) => n.startsWith("Fixture")) ??
        wb.SheetNames[0];

      if (!targetSheet) throw new Error("ไม่พบ sheet ในไฟล์");

      const ws = wb.Sheets[targetSheet];
      const data = XLSX.utils.sheet_to_json<FixtureRow>(ws, { defval: "" });

      setSheetUsed(targetSheet);
      setRows(data);
      setStatus("done");
      setStatusMsg(`โหลดเสร็จ — ${data.length.toLocaleString()} แถว จาก sheet "${targetSheet}"`);
    } catch (e) {
      setStatus("error");
      setErrorMsg(String(e));
    }
  };

  const handleFileChange = (files: File[]) => {
    const f = files[0] ?? null;
    setFixtureFile(f);
    if (f) handleParse(f);
    else {
      setStatus("idle");
      setRows([]);
      setSheetUsed("");
      setSearch("");
      setPage(0);
    }
  };

  const filtered = rows.filter((r) => {
    if (!search) return true;
    const s = search.toLowerCase();
    return (
      String(r.POG).toLowerCase().includes(s) ||
      String(r["Code Fixture"]).toLowerCase().includes(s) ||
      String(r.SEG).toLowerCase().includes(s)
    );
  });

  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const pageRows = filtered.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  const matchRate =
    rows.length > 0
      ? `${rows.length.toLocaleString()} รายการ`
      : "";

  return (
    <div className="max-w-5xl mx-auto p-6 space-y-6">

      {/* Header card */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div className="flex items-start gap-4">
          <div className="bg-gradient-to-br from-pink-50 to-orange-50 rounded-xl p-3 flex-shrink-0">
            <Database className="w-8 h-8 text-[#E91E8C]" />
          </div>
          <div>
            <h2 className="text-lg font-bold text-slate-800">Master Data</h2>
            <p className="text-sm text-slate-500 mt-1">
              จัดการไฟล์ข้อมูล Master สำหรับใช้งานในระบบ
            </p>
            <div className="mt-3 flex flex-wrap gap-2">
              <span className="px-2 py-0.5 rounded-full bg-pink-50 text-[#E91E8C] text-xs font-medium border border-pink-100">
                Fixture Index
              </span>
              <span className="px-2 py-0.5 rounded-full bg-slate-100 text-slate-500 text-xs font-medium">
                Sheet: Fixture_2026
              </span>
              <span className="px-2 py-0.5 rounded-full bg-slate-100 text-slate-500 text-xs font-medium">
                Columns: SEG · POG · Code Fixture
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Fixture Index section */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-5">
        <div className="flex items-center justify-between">
          <h3 className="font-semibold text-slate-700">Fixture Index</h3>
          {status === "done" && matchRate && (
            <span className="text-sm text-emerald-600 font-medium">{matchRate}</span>
          )}
        </div>

        <div className="max-w-md">
          <DropZone
            label="Fixture_Index (APR 2026).xlsx"
            accept=".xlsx"
            files={fixtureFile ? [fixtureFile] : []}
            onFiles={handleFileChange}
            hint={`Sheet: ${sheetUsed || "Fixture_2026"} — SEG, POG, Code Fixture`}
          />
        </div>

        {/* Parsing indicator */}
        {status === "parsing" && (
          <div className="flex items-center gap-3">
            <div className="animate-spin rounded-full h-5 w-5 border-2 border-pink-200 border-t-[#E91E8C] flex-shrink-0" />
            <p className="text-sm text-slate-600">{statusMsg}</p>
          </div>
        )}

        {/* Error */}
        {status === "error" && (
          <div className="bg-red-50 border border-red-200 rounded-xl p-4 text-red-700 text-sm">
            ❌ {errorMsg}
          </div>
        )}

        {/* Data table */}
        {status === "done" && rows.length > 0 && (
          <div className="space-y-3">
            {/* Search bar */}
            <div className="flex items-center justify-between flex-wrap gap-3">
              <p className="text-sm text-slate-500">{statusMsg}</p>
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="ค้นหา POG / Code Fixture / SEG..."
                  value={search}
                  onChange={(e) => { setSearch(e.target.value); setPage(0); }}
                  className="pl-9 pr-8 py-2 text-sm border border-slate-200 rounded-lg
                    focus:outline-none focus:ring-2 focus:ring-pink-200 focus:border-transparent w-72"
                />
                {search && (
                  <button
                    onClick={() => { setSearch(""); setPage(0); }}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>
            </div>

            {/* Table */}
            <div className="overflow-x-auto rounded-xl border border-slate-200">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    <th className="text-left px-4 py-3 font-semibold text-slate-600 w-20 text-xs uppercase tracking-wide">
                      SEG
                    </th>
                    <th className="text-left px-4 py-3 font-semibold text-slate-600 text-xs uppercase tracking-wide">
                      POG
                    </th>
                    <th className="text-left px-4 py-3 font-semibold text-slate-600 w-40 text-xs uppercase tracking-wide">
                      Code Fixture
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pageRows.map((row, i) => (
                    <tr key={i} className="hover:bg-slate-50/70 transition-colors">
                      <td className="px-4 py-2.5 text-slate-500 font-mono text-xs">{row.SEG}</td>
                      <td className="px-4 py-2.5 text-slate-700 text-xs">{row.POG}</td>
                      <td className="px-4 py-2.5 text-slate-700 font-mono text-xs">{row["Code Fixture"]}</td>
                    </tr>
                  ))}
                  {pageRows.length === 0 && (
                    <tr>
                      <td colSpan={3} className="px-4 py-8 text-center text-slate-400 text-sm">
                        ไม่พบข้อมูลที่ตรงกับ &quot;{search}&quot;
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            {totalPages > 1 && (
              <div className="flex items-center justify-between text-sm text-slate-500 pt-1">
                <span>
                  แสดง {page * PAGE_SIZE + 1}–{Math.min((page + 1) * PAGE_SIZE, filtered.length).toLocaleString()}{" "}
                  จาก {filtered.length.toLocaleString()} แถว
                  {search && ` (กรองจาก ${rows.length.toLocaleString()})`}
                </span>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setPage((p) => Math.max(0, p - 1))}
                    disabled={page === 0}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 text-xs
                      disabled:opacity-40 hover:bg-slate-50 transition-colors disabled:cursor-not-allowed"
                  >
                    ← ก่อนหน้า
                  </button>
                  <span className="px-2 text-xs">
                    {page + 1} / {totalPages}
                  </span>
                  <button
                    onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))}
                    disabled={page >= totalPages - 1}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 text-xs
                      disabled:opacity-40 hover:bg-slate-50 transition-colors disabled:cursor-not-allowed"
                  >
                    ถัดไป →
                  </button>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
