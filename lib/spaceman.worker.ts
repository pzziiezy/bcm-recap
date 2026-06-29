import * as XLSX from "xlsx";

type InMsg = { type: "parse"; buffer: ArrayBuffer };

// Cap rows sent to main thread to avoid structured-clone crash on huge files.
// Unique values (CATEGORY/SUBCATEGORY/DESC_C) are computed here and sent separately.
const MAX_DISPLAY_ROWS = 50_000;

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

    // Column indices for unique-value extraction (computed in worker to avoid
    // sending all rows to main thread for re-iteration there)
    const catIdx   = headers.indexOf("CATEGORY");
    const subIdx   = headers.indexOf("SUBCATEGORY");
    const descAIdx = headers.indexOf("DESC_A");
    const descBIdx = headers.indexOf("DESC_B");
    const descCIdx = headers.indexOf("DESC_C");

    self.postMessage({ type: "progress", pct: 25 });

    const totalRowsInSheet = range.e.r;
    const rows: Record<string, string>[] = [];
    const colCount = range.e.c + 1;

    const catSet   = new Set<string>();
    const subSet   = new Set<string>();
    const descASet = new Set<string>();
    const descBSet = new Set<string>();
    const descCSet = new Set<string>();
    let totalRows = 0; // actual non-empty rows across the full file

    for (let r = 1; r <= totalRowsInSheet; r++) {
      const row: Record<string, string> = {};
      let hasValue = false;
      for (let c = 0; c < colCount; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        const val = cell?.v != null ? String(cell.v) : "";
        row[headers[c]] = val;
        if (val) hasValue = true;
      }
      if (!hasValue) continue;

      totalRows++;

      // Collect unique values from EVERY row (not limited by display cap)
      if (catIdx   >= 0 && row[headers[catIdx]])   catSet.add(row[headers[catIdx]]);
      if (subIdx   >= 0 && row[headers[subIdx]])   subSet.add(row[headers[subIdx]]);
      if (descAIdx >= 0 && row[headers[descAIdx]]) descASet.add(row[headers[descAIdx]]);
      if (descBIdx >= 0 && row[headers[descBIdx]]) descBSet.add(row[headers[descBIdx]]);
      if (descCIdx >= 0 && row[headers[descCIdx]]) descCSet.add(row[headers[descCIdx]]);

      // Only keep first MAX_DISPLAY_ROWS for the table display
      if (rows.length < MAX_DISPLAY_ROWS) rows.push(row);

      // Progress every 15,000 rows (avoids postMessage overhead)
      if (r % 15000 === 0) {
        self.postMessage({
          type: "progress",
          pct: 25 + Math.floor((r / totalRowsInSheet) * 70),
        });
      }
    }

    self.postMessage({
      type: "done",
      headers,
      rows,
      totalRows,
      uniqueCategories:    [...catSet].sort(),
      uniqueSubcategories: [...subSet].sort(),
      uniqueDescA:         [...descASet].sort(),
      uniqueDescB:         [...descBSet].sort(),
      uniqueDescC:         [...descCSet].sort(),
    });
  } catch (err) {
    self.postMessage({ type: "error", message: String(err) });
  }
});
