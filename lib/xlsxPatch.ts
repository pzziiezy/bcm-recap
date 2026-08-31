/**
 * Shared ZIP/XML primitives for patching an .xlsx file in place, without a SheetJS
 * read→write round-trip (which regenerates styles.xml, drops named ranges, etc.).
 *
 * Used by download.worker.ts (RECAP) and minorReportDownload.worker.ts (Minor Report) —
 * kept dependency-free (no fflate import; callers do their own unzip/zip) so this stays
 * cheap to pull into either Worker bundle.
 */

import type { FillRow } from "./download";

// ── XML helpers ──────────────────────────────────────────────────────────────

export function encodeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function decodeXml(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

// ── Column helpers ───────────────────────────────────────────────────────────

/** 0-based column index → Excel letter(s). 0→"A", 25→"Z", 26→"AA" */
export function colLetter(idx: number): string {
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
export function colLetterIdx(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64;
  return n - 1;
}

// ── Shared String Table (SST) ────────────────────────────────────────────────

/**
 * Extract plain-text value for every <si> in the SST.
 * Handles both simple <t>...</t> and rich-text <r><t>...</t></r> entries.
 */
export function parseSST(xml: string): string[] {
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
export function appendSST(xml: string, newStrings: string[]): string {
  if (!newStrings.length) return xml;
  const newSis = newStrings.map(s => `<si><t>${encodeXml(s)}</t></si>`).join("");
  const at = xml.lastIndexOf("</sst>");
  let result = xml.slice(0, at) + newSis + xml.slice(at);
  result = result
    .replace(/\bcount="(\d+)"/, (_, n) => `count="${+n + newStrings.length}"`)
    .replace(/\buniqueCount="(\d+)"/, (_, n) => `uniqueCount="${+n + newStrings.length}"`);
  return result;
}

export function buildSST(strings: string[]): string {
  const n = strings.length;
  const sis = strings.map(s => `<si><t>${encodeXml(s)}</t></si>`).join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
    ` count="${n}" uniqueCount="${n}">${sis}</sst>`
  );
}

// ── Workbook → sheet file lookup ─────────────────────────────────────────────

export function findSheetPath(wbXml: string, relsXml: string, name: string): string | null {
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

/**
 * Set a specific cell to a shared-string value.
 * If the cell already exists, the original s= (style index) is preserved —
 * only the type and value change.  If it doesn't exist, it's inserted in
 * column order.
 */
export function patchCell(
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

/**
 * Write a numeric cell (no t= attribute in XLSX = number).
 * Preserves existing s= style attribute if the cell already exists.
 */
export function patchNumericCell(
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

// ── Upsert rows via ZIP-patch ─────────────────────────────────────────────────

/**
 * Write FillRow data into sheet XML using upsert logic:
 *   • If a <row r="N"> already exists (self-closing or with content), PATCH it —
 *     preserving existing row/cell styles while adding our values.
 *   • Only INSERT a new <row> for positions that have no XML element at all.
 *
 * This is critical for styled templates that pre-format empty rows with
 * borders/colours. An "always-append" approach would create duplicate row
 * numbers; Excel silently uses the first occurrence and ignores the later one,
 * so data never appears in the downloaded file.
 */
export function insertFillRows(
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
  // These are pre-formatted empty rows common in styled templates.
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
    .map(({ rowIndex }) => {
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

// ── Header-row auto-detection ─────────────────────────────────────────────────

/** Reads one <c>...</c> cell's plain-text value, resolving shared-string refs. */
function readCellText(attrs: string, inner: string, sstStrings: string[]): string | null {
  const vMatch = /<v>([^<]*)<\/v>/.exec(inner);
  if (!vMatch) return null;
  const raw = decodeXml(vMatch[1]);
  return /\bt="s"/.test(attrs) ? (sstStrings[+raw] ?? "") : raw;
}

/**
 * Find the Excel row number (1-based) that contains a cell matching `markerText`
 * (case-insensitive, trimmed) — used to auto-detect a template's header row instead of
 * hardcoding an assumed row count. Counting legend/filler rows from a screenshot has
 * proven unreliable (real templates can have extra blank/filter-only rows that don't
 * show up clearly), so this lets the actual uploaded file tell us the truth every time,
 * regardless of how many rows precede the header in a given template revision.
 *
 * Deliberately searches every cell in the row rather than one fixed column — the
 * template's columns have been reordered before, so the marker's own position can't be
 * assumed either. Cheap enough: header rows are short and this only scans the first
 * `maxScanRows` rows.
 */
export function findHeaderRowNum(
  sheetXml: string,
  sstStrings: string[],
  markerText: string,
  maxScanRows = 30
): number | null {
  const target = markerText.trim().toLowerCase();
  const rowRe = /<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let m: RegExpExecArray | null;
  while ((m = rowRe.exec(sheetXml)) !== null) {
    const rowNum = +m[1];
    if (rowNum > maxScanRows) break;
    const cellRe = /<c r="[A-Z]+\d+"([^>]*)>([\s\S]*?)<\/c>/g;
    let cm: RegExpExecArray | null;
    while ((cm = cellRe.exec(m[2])) !== null) {
      const value = readCellText(cm[1], cm[2], sstStrings);
      if (value !== null && value.trim().toLowerCase() === target) return rowNum;
    }
  }
  return null;
}

/**
 * Header text (trimmed, case-insensitive) → 0-based column index, read from one
 * already-located header row. Paired with findHeaderRowNum(): once the header row is
 * known, this tells every field WHERE it actually lives in the uploaded template —
 * so reordering the template's columns needs no code change, only the header text
 * itself has to stay the same.
 */
export function mapHeaderColumns(
  sheetXml: string,
  sstStrings: string[],
  headerRowNum: number
): Map<string, number> {
  const map = new Map<string, number>();
  const rowRe = new RegExp(`<row r="${headerRowNum}"[^>]*>([\\s\\S]*?)</row>`);
  const rowMatch = rowRe.exec(sheetXml);
  if (!rowMatch) return map;
  const cellRe = /<c r="([A-Z]+)\d+"([^>]*)>([\s\S]*?)<\/c>/g;
  let cm: RegExpExecArray | null;
  while ((cm = cellRe.exec(rowMatch[1])) !== null) {
    const value = readCellText(cm[2], cm[3], sstStrings);
    if (value === null) continue;
    const norm = value.trim().toLowerCase();
    if (norm && !map.has(norm)) map.set(norm, colLetterIdx(cm[1]));
  }
  return map;
}

// ── Replace a data-row region wholesale ───────────────────────────────────────

/**
 * Remove every <row> element numbered > headerRowCount, then append fresh <row>
 * elements built from `rows` right after the header rows.
 *
 * Unlike insertFillRows() (which upserts additively — RECAP repeatedly enriches a
 * fixed, known set of pre-existing barcode rows), this is for reports that are fully
 * regenerated on every build from a variable-length dataset: any stale sample/data
 * rows left over from a previous run (or from the uploaded template itself) must not
 * survive into the output, so the data region is wiped and rebuilt every time.
 *
 * To still preserve the template's exact visual format (borders, number formats,
 * row height) for the new data, this captures the row-level attributes and per-column
 * cell styles from the template's first data row (headerRowCount + 1) BEFORE wiping it,
 * then stamps that same style onto every generated row — so as far as Excel is
 * concerned, every output row looks like the template author's own sample row.
 */
export function replaceDataRows(
  sheetXml: string,
  headerRowCount: number,
  rows: FillRow[],
  sstStrings: string[]
): { sheetXml: string; newStrings: string[] } {
  const allStrings = [...sstStrings];
  const ssIdx = (v: string): number => {
    let i = allStrings.indexOf(v);
    if (i < 0) { i = allStrings.length; allStrings.push(v); }
    return i;
  };

  // Capture the template's sample data row style (row-level attrs + per-column cell
  // styles) before it gets wiped below. Works whether that row is self-closing
  // (<row .../>, a pre-formatted empty row) or open/close (<row ...>...</row>).
  const sampleRowNum = headerRowCount + 1;
  let rowAttrs = "";
  const styleByCol = new Map<number, string>();
  const sampleOpen =
    new RegExp(`<row r="${sampleRowNum}"([^>]*?)/>`).exec(sheetXml) ??
    new RegExp(`<row r="${sampleRowNum}"([^>]*)>([\\s\\S]*?)</row>`).exec(sheetXml);
  if (sampleOpen) {
    rowAttrs = sampleOpen[1].replace(/\bspans="[^"]*"/, "").trim();
    if (rowAttrs) rowAttrs = " " + rowAttrs;
    const inner = sampleOpen[2] ?? "";
    const cellRe = /<c r="([A-Z]+)\d+"([^>]*)(?:\/>|>)/g;
    let cm: RegExpExecArray | null;
    while ((cm = cellRe.exec(inner)) !== null) {
      const sM = /\bs="(\d+)"/.exec(cm[2]);
      if (sM) styleByCol.set(colLetterIdx(cm[1]), sM[1]);
    }
  }

  const buildCells = (fillRow: FillRow, rowNum: number) =>
    fillRow.cells
      .filter(c => c.value !== "")
      .sort((a, b) => a.col - b.col)
      .map(({ col, value }) => {
        const s = styleByCol.get(col);
        return `<c r="${colLetter(col)}${rowNum}"${s ? ` s="${s}"` : ""} t="s"><v>${ssIdx(value)}</v></c>`;
      })
      .join("");

  // Strip existing data rows (self-closing and open/close forms) beyond the header.
  const stripRow = (re: RegExp, getRowNum: (m: RegExpMatchArray) => number) =>
    (xml: string) => xml.replace(re, (full, ...groups) => {
      const m = [full, ...groups] as unknown as RegExpMatchArray;
      return getRowNum(m) > headerRowCount ? "" : full;
    });

  let result = sheetXml;
  result = stripRow(/<row\b([^>]*?)\/>/g, m => {
    const rm = /\br="(\d+)"/.exec(m[1]);
    return rm ? +rm[1] : -1;
  })(result);
  result = stripRow(/(<row\b[^>]*>)([\s\S]*?)(<\/row>)/g, m => {
    const rm = /\br="(\d+)"/.exec(m[1]);
    return rm ? +rm[1] : -1;
  })(result);

  // Append fresh rows for the whole dataset, in order, right after the header rows.
  // Every row carries the template's captured sample-row attrs/styles (see above).
  const sorted = [...rows].sort((a, b) => a.rowIndex - b.rowIndex);
  const newRowXml = sorted
    .map(r => `<row r="${r.rowIndex + 1}"${rowAttrs}>${buildCells(r, r.rowIndex + 1)}</row>`)
    .join("");

  if (newRowXml) {
    const sdClose = result.lastIndexOf("</sheetData>");
    const sdSelf  = result.indexOf("<sheetData/>");
    if (sdClose >= 0) {
      result = result.slice(0, sdClose) + newRowXml + result.slice(sdClose);
    } else if (sdSelf >= 0) {
      result = result.slice(0, sdSelf) + `<sheetData>${newRowXml}</sheetData>` + result.slice(sdSelf + 12);
    } else {
      const wsEnd = result.lastIndexOf("</worksheet>");
      if (wsEnd >= 0)
        result = result.slice(0, wsEnd) + `<sheetData>${newRowXml}</sheetData>` + result.slice(wsEnd);
    }
  }

  // Extend/shrink <dimension ref="…"> end-row to match the actual data written.
  const maxRowNum = sorted.length > 0 ? sorted[sorted.length - 1].rowIndex + 1 : headerRowCount;
  result = result.replace(
    /(<dimension\b[^>]*ref=")([^"]+)(")/,
    (m, pre, ref, post) => {
      const parts = ref.split(":");
      const startRef = parts[0];
      const endRef   = parts.length >= 2 ? parts[1] : parts[0];
      const endMatch = /^([A-Z]+)(\d+)$/.exec(endRef);
      if (!endMatch) return m;
      return `${pre}${startRef}:${endMatch[1]}${maxRowNum}${post}`;
    }
  );

  return { sheetXml: result, newStrings: allStrings.slice(sstStrings.length) };
}
