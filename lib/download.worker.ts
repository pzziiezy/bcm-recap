import * as XLSX from "xlsx";
import type { DownloadRow } from "./download";

const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

type InMsg =
  | { type: "init"; buffer: ArrayBuffer }
  | { type: "build"; rows: DownloadRow[] };

let template: ArrayBuffer | null = null;

function applyRows(ws: XLSX.WorkSheet, rows: DownloadRow[]): void {
  for (const row of rows) {
    const data = row.override
      ? { ...(row.filled ?? {}), ...row.override }
      : row.filled;
    if (!data) continue;
    const w = (c: number, v: string) => {
      const addr = XLSX.utils.encode_cell({ r: row.rowIndex, c });
      const existing = ws[addr];
      // Preserve original cell style (fill color, borders, etc.) if present
      ws[addr] = { ...(existing?.s != null ? { s: existing.s } : {}), t: "s", v };
    };
    w(5,  data.division  ?? "");
    w(6,  data.dept      ?? "");
    w(7,  data.subDept   ?? "");
    w(8,  data.cls       ?? "");
    w(9,  data.planogram ?? "");
    w(13, data.colN      ?? "");
    w(14, data.colO      ?? "");
  }
}

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

      // cellStyles: parse fill/font/border so they survive round-trip
      // sheetStubs: create stubs for styled-but-empty cells so their styles are preserved when we write values into them
      const wb = XLSX.read(template, { type: "array", cellStyles: true, sheetStubs: true });
      const ws = wb.Sheets["NEW SCM"];
      if (!ws) throw new Error('Sheet "NEW SCM" not found');

      ctx.postMessage({ type: "progress", pct: 35 });
      applyRows(ws, msg.rows);

      ctx.postMessage({ type: "progress", pct: 70 });
      const out = XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;

      ctx.postMessage({ type: "progress", pct: 95 });
      ctx.postMessage({ type: "done", buffer: out }, [out]);
    } catch (err) {
      ctx.postMessage({ type: "error", message: String(err) });
    }
  }
});
