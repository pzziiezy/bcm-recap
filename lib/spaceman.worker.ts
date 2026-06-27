import * as XLSX from "xlsx";

type InMsg = { type: "parse"; buffer: ArrayBuffer };

addEventListener("message", (e: MessageEvent<InMsg>) => {
  if (e.data.type !== "parse") return;
  const { buffer } = e.data;

  try {
    self.postMessage({ type: "progress", pct: 5 });

    // Memory-optimised read: skip computed cell properties we don't need
    const wb = XLSX.read(buffer, {
      type: "array",
      cellText: false,
      cellHTML: false,
      cellNF: false,
      cellDates: false,
    });

    self.postMessage({ type: "progress", pct: 20 });

    const ws = wb.Sheets["QRY_Product_by_POG"];
    if (!ws) {
      self.postMessage({ type: "error", message: 'ไม่พบ Sheet "QRY_Product_by_POG" ในไฟล์' });
      return;
    }

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

    // Extract headers
    const headers: string[] = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
      headers.push(cell?.v != null ? String(cell.v).trim() : `คอลัมน์ ${c + 1}`);
    }

    self.postMessage({ type: "progress", pct: 25 });

    // Extract rows with periodic progress updates
    const totalRows = range.e.r;
    const rows: Record<string, string>[] = [];
    const colCount = range.e.c + 1;

    for (let r = 1; r <= totalRows; r++) {
      const row: Record<string, string> = {};
      let hasValue = false;
      for (let c = 0; c < colCount; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        const val = cell?.v != null ? String(cell.v) : "";
        row[headers[c]] = val;
        if (val) hasValue = true;
      }
      if (hasValue) rows.push(row);

      // Progress every 15,000 rows (avoids postMessage overhead)
      if (r % 15000 === 0) {
        self.postMessage({
          type: "progress",
          pct: 25 + Math.floor((r / totalRows) * 70),
        });
      }
    }

    self.postMessage({ type: "done", headers, rows });
  } catch (err) {
    self.postMessage({ type: "error", message: String(err) });
  }
});
