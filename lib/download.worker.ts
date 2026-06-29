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
import type { DownloadRow } from "./download";

const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

type InMsg =
  | { type: "init"; buffer: ArrayBuffer }
  | { type: "build"; rows: DownloadRow[] };

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
  const sm = new RegExp(
    `<sheet\\b[^>]+name="${xmlName}"[^>]+r:id="([^"]+)"`,
    "i"
  ).exec(wbXml);
  if (!sm) return null;

  const rm = new RegExp(
    `<Relationship\\b[^>]+Id="${sm[1]}"[^>]+Target="([^"]+)"`,
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
  [9, "planogram"], [13, "colN"], [14, "colO"],
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

    // Column P (index 15) — always computed from effective O and N
    const oNum = parseFloat((data as Record<string, string>).colO ?? "100") || 100;
    const nNum = parseFloat((data as Record<string, string>).colN ?? "0") || 0;
    if (nNum !== 0 || (data as Record<string, string>).colN) {
      const pVal = Math.round((oNum / 100) * nNum * 100) / 100;
      cols.set(15, { kind: "n", value: pVal });
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
      ctx.postMessage({ type: "progress", pct: 40 });

      // 3. Parse the shared strings table
      const sstPath    = "xl/sharedStrings.xml";
      const sstStrings = files[sstPath] ? parseSST(strFromU8(files[sstPath])) : [];
      ctx.postMessage({ type: "progress", pct: 55 });

      // 4. Patch cell values in the sheet XML
      const { sheetXml, newStrings } = applyRows(
        strFromU8(files[sheetPath]),
        msg.rows,
        sstStrings
      );
      files[sheetPath] = strToU8(sheetXml);
      ctx.postMessage({ type: "progress", pct: 70 });

      // 5. Append any new strings to the SST (existing entries untouched)
      if (newStrings.length > 0) {
        files[sstPath] = strToU8(
          files[sstPath]
            ? appendSST(strFromU8(files[sstPath]), newStrings)
            : buildSST([...sstStrings, ...newStrings])
        );
      }
      ctx.postMessage({ type: "progress", pct: 85 });

      // 6. Rezip — styles.xml, workbook.xml, relationships, etc. unchanged
      const out = zipSync(files);
      // Ensure we hand off a clean ArrayBuffer slice (fflate may over-allocate)
      const outBuf = out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
      ctx.postMessage({ type: "progress", pct: 95 });
      ctx.postMessage({ type: "done", buffer: outBuf }, [outBuf]);

    } catch (err) {
      ctx.postMessage({ type: "error", message: String(err) });
    }
  }
});
