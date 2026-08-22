/**
 * Build worker — patches cell values directly inside the XLSX ZIP.
 *
 * Instead of SheetJS read→write (which regenerates styles.xml, drops named
 * ranges, etc.), we:
 *   1. unzip the original XLSX byte-for-byte
 *   2. patch ONLY xl/worksheets/sheet[N].xml  (cell values)
 *   3. append new strings to xl/sharedStrings.xml  (if needed)
 *   4. rezip — every other file (styles.xml, workbook.xml, rels, …) stays
 *      untouched, so colours, freeze panes, named ranges, groups all survive
 *
 * The generic ZIP/XML primitives (SST, cell patching, findSheetPath, insertFillRows)
 * live in ./xlsxPatch — shared with minorReportDownload.worker.ts. Only the RECAP-
 * specific F-J/N-Q column patch (applyRows) stays local to this file.
 */

import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import type { DownloadRow, CheckSpaceFillPlan } from "./download";
import { computeNetCapacity } from "./netCapacity";
import {
  colLetter,
  parseSST,
  appendSST,
  buildSST,
  findSheetPath,
  patchCell,
  patchNumericCell,
  insertFillRows,
} from "./xlsxPatch";

const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

type InMsg =
  | { type: "init"; buffer: ArrayBuffer }
  | { type: "build"; rows: DownloadRow[]; checkSpacePlan?: CheckSpaceFillPlan };

let template: ArrayBuffer | null = null;

// ── RECAP-specific F-J/N-Q cell patcher ─────────────────────────────────────

// String columns to fill (col index → FilledData key)
const STRING_FILL_COLS: Array<[number, string]> = [
  [5, "division"], [6, "dept"], [7, "subDept"], [8, "cls"],
  [9, "planogram"], [13, "colN"], [14, "colPiece"], [15, "colO"],
];

type CellTarget =
  | { kind: "s"; ssIdx: number }
  | { kind: "n"; value: number };

function applyRows(
  sheetXml: string,
  rows: DownloadRow[],
  sstStrings: string[]
): { sheetXml: string; newStrings: string[] } {
  // Build: Excel-row-number → Map<colIdx, CellTarget>
  const target = new Map<number, Map<number, CellTarget>>();

  // Shared string lookup (find existing or queue new)
  const allStrings = [...sstStrings];
  const ssIdx = (v: string): number => {
    let i = allStrings.indexOf(v);
    if (i < 0) { i = allStrings.length; allStrings.push(v); }
    return i;
  };

  for (const row of rows) {
    const data = row.override ? { ...row.filled, ...row.override } : row.filled;
    if (!data) continue;
    const excelRow = row.rowIndex + 1;
    const cols = new Map<number, CellTarget>();

    // String columns
    for (const [ci, key] of STRING_FILL_COLS) {
      const v = (data as Record<string, string>)[key];
      if (v) cols.set(ci, { kind: "s", ssIdx: ssIdx(v) });
    }

    // Column Q (index 16) — Net = colO% × colPiece
    const qVal = computeNetCapacity(
      (data as Record<string, string>).colO,
      (data as Record<string, string>).colPiece
    );
    if (qVal !== null) {
      cols.set(16, { kind: "n", value: qVal });
    }

    if (cols.size) target.set(excelRow, cols);
  }

  if (!target.size) return { sheetXml, newStrings: [] };

  // Process each <row> element in the sheet XML
  const result = sheetXml.replace(
    /(<row\b[^>]*>)([\s\S]*?)(<\/row>)/g,
    (full, open, inner, close) => {
      const rm = /\br="(\d+)"/.exec(open);
      if (!rm) return full;
      const cols = target.get(+rm[1]);
      if (!cols) return full;

      let cells = inner;
      for (const [ci, cell] of cols) {
        if (cell.kind === "s") {
          cells = patchCell(cells, colLetter(ci), +rm[1], cell.ssIdx, ci);
        } else {
          cells = patchNumericCell(cells, colLetter(ci), +rm[1], cell.value, ci);
        }
      }
      return open + cells + close;
    }
  );

  return { sheetXml: result, newStrings: allStrings.slice(sstStrings.length) };
}

// ── Worker message handler ───────────────────────────────────────────────────

addEventListener("message", (e: MessageEvent<InMsg>) => {
  const msg = e.data;

  if (msg.type === "init") {
    template = msg.buffer;
    ctx.postMessage({ type: "init_ok" });
    return;
  }

  if (msg.type === "build") {
    if (!template) {
      ctx.postMessage({ type: "error", message: "Template not initialized" });
      return;
    }
    try {
      ctx.postMessage({ type: "progress", pct: 10 });

      // 1. Unzip the XLSX (it's just a ZIP)
      const files = unzipSync(new Uint8Array(template));
      ctx.postMessage({ type: "progress", pct: 25 });

      // 2. Locate the "NEW SCM" worksheet file
      const wbXml    = strFromU8(files["xl/workbook.xml"]);
      const relsXml  = strFromU8(files["xl/_rels/workbook.xml.rels"]);
      const sheetPath = findSheetPath(wbXml, relsXml, "NEW SCM");
      if (!sheetPath || !files[sheetPath]) throw new Error('Sheet "NEW SCM" not found');
      ctx.postMessage({ type: "progress", pct: 35 });

      // 3. Parse the shared strings table
      const sstPath    = "xl/sharedStrings.xml";
      const sstStrings = files[sstPath] ? parseSST(strFromU8(files[sstPath])) : [];

      // Working SST accumulator — grows as Check Space fills and F-J fills add strings.
      // All operations share the same index space so cell references stay consistent.
      const workingStrings = [...sstStrings];

      // 4. Apply Check Space fills via ZIP-patch (no XLSX.write — format preserved)
      let newScmXml = strFromU8(files[sheetPath]);

      if (msg.checkSpacePlan) {
        // 4a. Insert new rows into NEW SCM (A,D,E,K,L,M,R+ columns)
        //     These rows will then have F-J patched by applyRows in step 5.
        if (msg.checkSpacePlan.newScmRows.length > 0) {
          const r1 = insertFillRows(newScmXml, msg.checkSpacePlan.newScmRows, workingStrings);
          newScmXml = r1.sheetXml;
          workingStrings.push(...r1.newStrings);
        }
        ctx.postMessage({ type: "progress", pct: 50 });

        // 4b. Write NEW_DELETE_IM and DEL SCM sheets
        for (const sf of msg.checkSpacePlan.extraSheets) {
          if (!sf.rows.length) {
            console.warn(`[worker] SKIP ${sf.sheetName}: 0 rows`);
            continue;
          }
          const p = findSheetPath(wbXml, relsXml, sf.sheetName);
          if (!p) {
            console.warn(`[worker] findSheetPath FAILED for "${sf.sheetName}" — sheet not found in workbook.xml`);
            continue;
          }
          if (!files[p]) {
            console.warn(`[worker] ZIP has no entry for path "${p}" (sheet "${sf.sheetName}")`);
            continue;
          }
          const r2 = insertFillRows(strFromU8(files[p]), sf.rows, workingStrings);
          files[p] = strToU8(r2.sheetXml);
          workingStrings.push(...r2.newStrings);
          console.info(`[worker] OK: inserted ${sf.rows.length} rows into "${sf.sheetName}" at ${p}`);
        }
      }
      ctx.postMessage({ type: "progress", pct: 60 });

      // 5. Patch F-J / N-Q in NEW SCM for all rows (existing + newly inserted CS rows).
      //    Pass workingStrings so new indices continue from where CS fills left off.
      const { sheetXml: patchedXml, newStrings: fjStrings } = applyRows(
        newScmXml,
        msg.rows,
        workingStrings
      );
      files[sheetPath] = strToU8(patchedXml);
      ctx.postMessage({ type: "progress", pct: 75 });

      // 6. Append ALL new strings to the SST in one pass
      //    = strings added by CS fills + strings added by F-J patches
      const allNewStrings = [
        ...workingStrings.slice(sstStrings.length),
        ...fjStrings,
      ];
      if (allNewStrings.length > 0) {
        files[sstPath] = strToU8(
          files[sstPath]
            ? appendSST(strFromU8(files[sstPath]), allNewStrings)
            : buildSST([...sstStrings, ...allNewStrings])
        );
      }
      ctx.postMessage({ type: "progress", pct: 88 });

      // 7. Rezip — styles.xml, workbook.xml, relationships, etc. unchanged
      const out = zipSync(files);
      const outBuf = out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
      ctx.postMessage({ type: "progress", pct: 95 });
      ctx.postMessage({ type: "done", buffer: outBuf }, [outBuf]);

    } catch (err) {
      ctx.postMessage({ type: "error", message: String(err) });
    }
  }
});
