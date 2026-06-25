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

    // XLSX.write อาจคืนค่าเป็น Array ธรรมดา ไม่ใช่ Uint8Array เสมอไป
    // ใช้ new Uint8Array() เพื่อให้แน่ใจว่าได้ buffer ที่ถูกต้อง
    const raw = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    const uint8 = new Uint8Array(raw as number[]);
    // ส่งกลับโดย copy (ไม่ใช้ transfer) เพื่อหลีกเลี่ยง DedicatedWorkerGlobalScope error
    postMessage({ ok: true, buffer: uint8.buffer.slice(0) });
  } catch (err) {
    postMessage({ ok: false, error: String(err) });
  }
});
