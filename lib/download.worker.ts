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
 */

import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import type { DownloadRow, CheckSpaceFillPlan, FillRow } from "./download";

const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

type InMsg =
  | { type: "init"; buffer: ArrayBuffer }
  | { type: "build"; rows: DownloadRow[]; checkSpacePlan?: CheckSpaceFillPlan };

let template: ArrayBuffer | null = null;

// ── XML helpers ──────────────────────────────────────────────────────────────

function encodeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function decodeXml(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

// ── Column helpers ───────────────────────────────────────────────────────────

/** 0-based column index → Excel letter(s). 0→"A", 25→"Z", 26→"AA" */
function colLetter(idx: number): string {
  let s = "";
  let n = idx + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Excel letter(s) → 0-based index. "A"→0, "Z"→25, "AA"→26 */
function colLetterIdx(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64;
  return n - 1;
}

// ── Shared String Table (SST) ────────────────────────────────────────────────

/**
 * Extract plain-text value for every <si> in the SST.
 * Handles both simple <t>...</t> and rich-text <r><t>...</t></r> entries.
 */
function parseSST(xml: string): string[] {
  const out: string[] = [];
  const siRe = /<si>([\s\S]*?)<\/si>/g;
  let m: RegExpExecArray | null;
  while ((m = siRe.exec(xml)) !== null) {
    let text = "";
    const tRe = /<t(?:\s[^>]*)?>([^<]*)<\/t>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRe.exec(m[1])) !== null) text += decodeXml(tm[1]);
    out.push(text);
  }
  return out;
}

/**
 * Append new plain-text strings to the existing SST XML.
 * Existing entries (including rich text) are kept byte-for-byte.
 */
function appendSST(xml: string, newStrings: string[]): string {
  if (!newStrings.length) return xml;
  const newSis = newStrings.map(s => `<si><t>${encodeXml(s)}</t></si>`).join("");
  const at = xml.lastIndexOf("</sst>");
  let result = xml.slice(0, at) + newSis + xml.slice(at);
  result = result
    .replace(/\bcount="(\d+)"/, (_, n) => `count="${+n + newStrings.length}"`)
    .replace(/\buniqueCount="(\d+)"/, (_, n) => `uniqueCount="${+n + newStrings.length}"`);
  return result;
}

function buildSST(strings: string[]): string {
  const n = strings.length;
  const sis = strings.map(s => `<si><t>${encodeXml(s)}</t></si>`).join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
    ` count="${n}" uniqueCount="${n}">${sis}</sst>`
  );
}

// ── Workbook → sheet file lookup ─────────────────────────────────────────────

function findSheetPath(wbXml: string, relsXml: string, name: string): string | null {
  // Encode the sheet name for XML attribute matching
  const xmlName = encodeXml(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  // Step 1: find the <sheet> element with this name (attribute-order independent)
  const sheetMatch = new RegExp(
    `<sheet\\b[^>]*name="${xmlName}"[^>]*/?>`,
    "i"
  ).exec(wbXml);
  if (!sheetMatch) return null;

  // Step 2: extract r:id from the matched element (regardless of attribute order)
  const ridMatch = /\br:id="([^"]+)"/.exec(sheetMatch[0]);
  if (!ridMatch) return null;

  const rm = new RegExp(
    `<Relationship\\b[^>]+Id="${ridMatch[1]}"[^>]+Target="([^"]+)"`,
    "i"
  ).exec(relsXml);
  if (!rm) return null;

  const t = rm[1];
  // Target can be an absolute path (/xl/...) or relative (worksheets/...)
  if (t.startsWith("/xl/")) return t.slice(1);
  if (t.startsWith("worksheets/")) return `xl/${t}`;
  return `xl/worksheets/${t}`;
}

// ── Cell patcher ─────────────────────────────────────────────────────────────

// String columns to fill (col index → FilledData key)
const STRING_FILL_COLS: Array<[number, string]> = [
  [5, "division"], [6, "dept"], [7, "subDept"], [8, "cls"],
  [9, "planogram"], [13, "colN"], [14, "colPiece"], [15, "colO"],
];

type CellTarget =
  | { kind: "s"; ssIdx: number }
  | { kind: "n"; value: number };

/**
 * Write a numeric cell (no t= attribute in XLSX = number).
 * Preserves existing s= style attribute if the cell already exists.
 */
function patchNumericCell(
  inner: string,
  letter: string,
  row: number,
  value: number,
  ci: number
): string {
  const ref = `${letter}${row}`;
  const pat = new RegExp(`<c r="${ref}"([^>]*?)(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const m = pat.exec(inner);

  if (m) {
    const sM = /\bs="(\d+)"/.exec(m[1]);
    const newCell = `<c r="${ref}"${sM ? ` s="${sM[1]}"` : ""}><v>${value}</v></c>`;
    return inner.slice(0, m.index) + newCell + inner.slice(m.index + m[0].length);
  }

  const newCell = `<c r="${ref}"><v>${value}</v></c>`;
  const scanPat = /<c\s+r="([A-Z]+)\d+"/g;
  let at = -1;
  let im: RegExpExecArray | null;
  while ((im = scanPat.exec(inner)) !== null) {
    if (colLetterIdx(im[1]) > ci) { at = im.index; break; }
  }
  return at >= 0
    ? inner.slice(0, at) + newCell + inner.slice(at)
    : inner + newCell;
}

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
    const pctNum   = parseFloat((data as Record<string, string>).colO     ?? "0") || 0;
    const pieceNum = parseFloat((data as Record<string, string>).colPiece ?? "0") || 0;
    if (pctNum > 0 && pieceNum > 0) {
      const qVal = Math.round((pctNum / 100) * pieceNum * 100) / 100;
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

/**
 * Set a specific cell to a shared-string value.
 * If the cell already exists, the original s= (style index) is preserved —
 * only the type and value change.  If it doesn't exist, it's inserted in
 * column order.
 */
function patchCell(
  inner: string,
  letter: string,
  row: number,
  ssIdx: number,
  ci: number
): string {
  const ref = `${letter}${row}`;

  // Match any existing cell element (self-closing or with content)
  const pat = new RegExp(`<c r="${ref}"([^>]*?)(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const m = pat.exec(inner);

  if (m) {
    // Preserve the style attribute; strip everything else (t, old value, formula)
    const sM = /\bs="(\d+)"/.exec(m[1]);
    const newCell = `<c r="${ref}"${sM ? ` s="${sM[1]}"` : ""} t="s"><v>${ssIdx}</v></c>`;
    return inner.slice(0, m.index) + newCell + inner.slice(m.index + m[0].length);
  }

  // Cell doesn't exist — insert before the next higher-indexed column
  const newCell = `<c r="${ref}" t="s"><v>${ssIdx}</v></c>`;
  const scanPat = /<c\s+r="([A-Z]+)\d+"/g;
  let at = -1;
  let im: RegExpExecArray | null;
  while ((im = scanPat.exec(inner)) !== null) {
    if (colLetterIdx(im[1]) > ci) { at = im.index; break; }
  }
  return at >= 0
    ? inner.slice(0, at) + newCell + inner.slice(at)
    : inner + newCell;
}

// ── Check Space fill — upsert rows via ZIP-patch ─────────────────────────────

/**
 * Write Check Space fill data into sheet XML using upsert logic:
 *   • If a <row r="N"> already exists (self-closing or with content), PATCH it —
 *     preserving existing row/cell styles while adding our values.
 *   • Only INSERT a new <row> for positions that have no XML element at all.
 *
 * This is critical for styled RECAP templates that pre-format empty rows with
 * borders/colours.  The old "always-append" approach created duplicate row
 * numbers; Excel silently uses the first occurrence and ignores the later one,
 * so data never appeared in the downloaded file.
 */
function insertFillRows(
  sheetXml: string,
  rows: FillRow[],
  sstStrings: string[]
): { sheetXml: string; newStrings: string[] } {
  if (!rows || rows.length === 0) return { sheetXml, newStrings: [] };

  const allStrings = [...sstStrings];
  const ssIdx = (v: string): number => {
    let i = allStrings.indexOf(v);
    if (i < 0) { i = allStrings.length; allStrings.push(v); }
    return i;
  };

  const sorted = [...rows]
    .filter(r => r.cells && r.cells.length > 0)
    .sort((a, b) => a.rowIndex - b.rowIndex);

  if (sorted.length === 0) return { sheetXml, newStrings: [] };

  // rowNum (1-based) → FillRow for quick lookup
  const rowMap = new Map<number, FillRow>();
  for (const r of sorted) rowMap.set(r.rowIndex + 1, r);

  const patchedRows = new Set<number>();

  // Build cell XML for a row (used when INSERTING a brand-new row)
  const buildCells = (fillRow: FillRow, rowNum: number) =>
    fillRow.cells
      .filter(c => c.value !== "")
      .sort((a, b) => a.col - b.col)
      .map(({ col, value }) => `<c r="${colLetter(col)}${rowNum}" t="s"><v>${ssIdx(value)}</v></c>`)
      .join("");

  // Pass 1 — patch self-closing rows: <row r="N" ... />
  // These are pre-formatted empty rows common in styled RECAP templates.
  // We expand them to open/close form and inject our cell data.
  let result = sheetXml.replace(
    /<row\b([^>]*?)\/>/g,
    (full, attrs) => {
      const rm = /\br="(\d+)"/.exec(attrs);
      if (!rm) return full;
      const rowNum = +rm[1];
      const fillRow = rowMap.get(rowNum);
      if (!fillRow) return full;
      patchedRows.add(rowNum);
      return `<row${attrs}>${buildCells(fillRow, rowNum)}</row>`;
    }
  );

  // Pass 2 — patch open/close rows: <row r="N" ...>...</row>
  // Use patchCell so any existing styled cells keep their s= attribute.
  result = result.replace(
    /(<row\b[^>]*>)([\s\S]*?)(<\/row>)/g,
    (full, open, inner, close) => {
      const rm = /\br="(\d+)"/.exec(open);
      if (!rm) return full;
      const rowNum = +rm[1];
      const fillRow = rowMap.get(rowNum);
      if (!fillRow) return full;
      patchedRows.add(rowNum);
      let cells = inner;
      for (const { col, value } of fillRow.cells) {
        if (!value) continue;
        cells = patchCell(cells, colLetter(col), rowNum, ssIdx(value), col);
      }
      return open + cells + close;
    }
  );

  // Pass 3 — insert rows that had no XML element at all
  const newRowXml = sorted
    .filter(r => !patchedRows.has(r.rowIndex + 1))
    .map(({ rowIndex, cells: _ }) => {
      const rowNum = rowIndex + 1;
      const fr = rowMap.get(rowNum)!;
      return `<row r="${rowNum}">${buildCells(fr, rowNum)}</row>`;
    })
    .join("");

  if (newRowXml) {
    const sdClose = result.lastIndexOf("</sheetData>");
    const sdSelf  = result.indexOf("<sheetData/>");
    if (sdClose >= 0) {
      result = result.slice(0, sdClose) + newRowXml + result.slice(sdClose);
    } else if (sdSelf >= 0) {
      result = result.slice(0, sdSelf)
        + `<sheetData>${newRowXml}</sheetData>`
        + result.slice(sdSelf + 12);
    } else {
      const wsEnd = result.lastIndexOf("</worksheet>");
      if (wsEnd >= 0)
        result = result.slice(0, wsEnd) + `<sheetData>${newRowXml}</sheetData>` + result.slice(wsEnd);
    }
  }

  // Extend <dimension ref="…"> end-row
  const maxRowNum = sorted[sorted.length - 1].rowIndex + 1;
  result = result.replace(
    /(<dimension\b[^>]*ref=")([^"]+)(")/,
    (m, pre, ref, post) => {
      const parts = ref.split(":");
      const startRef = parts[0];
      const endRef   = parts.length >= 2 ? parts[1] : parts[0];
      const endMatch = /^([A-Z]+)(\d+)$/.exec(endRef);
      if (!endMatch) return m;
      const newEnd = Math.max(parseInt(endMatch[2]), maxRowNum);
      return `${pre}${startRef}:${endMatch[1]}${newEnd}${post}`;
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
