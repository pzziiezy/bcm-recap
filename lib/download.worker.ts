import * as XLSX from "xlsx";

interface RowData {
  rowIndex: number;
  filled: Record<string, string> | null;
  override?: Record<string, string>;
}

addEventListener("message", (e: MessageEvent) => {
  const { buffer, results } = e.data as {
    buffer: ArrayBuffer;
    results: RowData[];
  };

  try {
    const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
    const ws = wb.Sheets["NEW SCM"];

    for (const row of results) {
      const data = row.override
        ? { ...(row.filled ?? {}), ...row.override }
        : row.filled;
      if (!data) continue;

      const write = (c: number, v: string) => {
        const addr = XLSX.utils.encode_cell({ r: row.rowIndex, c });
        ws[addr] = { t: "s", v };
      };

      write(5, data.division ?? "");
      write(6, data.dept ?? "");
      write(7, data.subDept ?? "");
      write(8, data.cls ?? "");
      write(9, data.planogram ?? "");
    }

    const out = XLSX.write(wb, { type: "array", bookType: "xlsx" }) as Uint8Array;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (self as any).postMessage({ ok: true, buffer: out.buffer }, [out.buffer]);
  } catch (err) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (self as any).postMessage({ ok: false, error: String(err) });
  }
});
