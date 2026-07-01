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

    // Hierarchy maps for cascade filtering: DESC_A→DESC_B→DESC_C→CATEGORY→SUBCATEGORY
    const divToDeptMap = new Map<string, Set<string>>();
    const deptToSubMap = new Map<string, Set<string>>();
    const subToClsMap  = new Map<string, Set<string>>();
    const clsToSubMap  = new Map<string, Set<string>>();

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

      // Collect unique flat values from EVERY row (not limited by display cap)
      const vCat   = catIdx   >= 0 ? row[headers[catIdx]]   : "";
      const vSub   = subIdx   >= 0 ? row[headers[subIdx]]   : "";
      const vDescA = descAIdx >= 0 ? row[headers[descAIdx]] : "";
      const vDescB = descBIdx >= 0 ? row[headers[descBIdx]] : "";
      const vDescC = descCIdx >= 0 ? row[headers[descCIdx]] : "";
      if (vCat)   catSet.add(vCat);
      if (vSub)   subSet.add(vSub);
      if (vDescA) descASet.add(vDescA);
      if (vDescB) descBSet.add(vDescB);
      if (vDescC) descCSet.add(vDescC);

      // Build cascade hierarchy (DESC_A→DESC_B→DESC_C→CATEGORY→SUBCATEGORY)
      if (vDescA && vDescB) {
        if (!divToDeptMap.has(vDescA)) divToDeptMap.set(vDescA, new Set());
        divToDeptMap.get(vDescA)!.add(vDescB);
      }
      if (vDescB && vDescC) {
        if (!deptToSubMap.has(vDescB)) deptToSubMap.set(vDescB, new Set());
        deptToSubMap.get(vDescB)!.add(vDescC);
      }
      if (vDescC && vCat) {
        if (!subToClsMap.has(vDescC)) subToClsMap.set(vDescC, new Set());
        subToClsMap.get(vDescC)!.add(vCat);
      }
      if (vCat && vSub) {
        if (!clsToSubMap.has(vCat)) clsToSubMap.set(vCat, new Set());
        clsToSubMap.get(vCat)!.add(vSub);
      }

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

    const toSortedRecord = (m: Map<string, Set<string>>) =>
      Object.fromEntries([...m].map(([k, v]) => [k, [...v].sort()]));

    const hierarchyMap = {
      divToDept: toSortedRecord(divToDeptMap),
      deptToSub: toSortedRecord(deptToSubMap),
      subToCls:  toSortedRecord(subToClsMap),
    };
    const catToSub = toSortedRecord(clsToSubMap);

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
      hierarchyMap,
      catToSub,
    });
  } catch (err) {
    self.postMessage({ type: "error", message: String(err) });
  }
});
