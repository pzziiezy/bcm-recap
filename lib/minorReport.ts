import type {
  PipelineSnapshot,
  ProcessedRow,
  FilledData,
  MinorReportSheets,
  MinorReportNewItemRow,
  MinorReportNewNotLinkRow,
  MinorReportDeleteItemRow,
} from "./types";
import { computeNetCapacity } from "./netCapacity";
import { buildRecapCodes } from "./processor";

/** Step 5 preview edits win over the originally-computed values, same merge pattern
 *  used by download.worker.ts's applyRows() and page.tsx's fillTabs sync. */
function mergedFilled(row: ProcessedRow): Partial<FilledData> | undefined {
  if (!row.filled && !row.override) return undefined;
  return { ...(row.filled ?? {}), ...(row.override ?? {}) };
}

/**
 * Reshape RECAP Auto-Filler's finished pipeline output into the 3 Minor Report sheets.
 * Pure function — no React, no refs, no file parsing. Everything it needs already lives
 * in the snapshot assembled at the end of handleProcess() (see app/page.tsx).
 *
 * Row-explosion logic — see doc §5/§6.2:
 *   Recap_New_item     = one row per (barcode × store) where NEW SCM has a store-flag
 *   Recap_New_not_link = one row per (barcode × store) where FILE_INDEX says the store
 *                        should carry this barcode's planogram, but NEW SCM has no flag
 *                        for that store — NOT mutually exclusive with Recap_New_item,
 *                        a barcode can appear in both if some stores link and some don't
 *   Recap_Delete_item  = one row per (barcode × store) where DEL SCM has a store-flag
 */
export function buildMinorReportSheets(snapshot: PipelineSnapshot): MinorReportSheets {
  const {
    results, indexLookup, barcodeMap, structureMap, byUpc,
    newScmStoreFlags, delScmStoreFlags, newScmRowInfo, delScmRowInfo, attributeMap,
  } = snapshot;

  const resultsByBarcode = new Map<string, ProcessedRow>();
  for (const row of results) {
    if (!resultsByBarcode.has(row.barcode)) resultsByBarcode.set(row.barcode, row);
  }

  const newItem: MinorReportNewItemRow[] = [];
  const newNotLink: MinorReportNewNotLinkRow[] = [];
  const deleteItem: MinorReportDeleteItemRow[] = [];

  // ── Recap_New_item + Recap_New_not_link ──────────────────────────────────
  for (const [barcode, storesWithQty] of newScmStoreFlags) {
    const pr = resultsByBarcode.get(barcode);
    const filled = pr ? mergedFilled(pr) : undefined;
    // Fallback to a direct sheet read for rows processRows() never touched — pre-existing
    // NEW SCM rows whose col F was already filled before this run (parseMissingRows skips them).
    const rowInfo = newScmRowInfo.get(barcode);

    const name      = pr?.name          ?? rowInfo?.name      ?? "";
    const division  = filled?.division  ?? rowInfo?.division  ?? "";
    const dept      = filled?.dept      ?? rowInfo?.dept      ?? "";
    const planogram = filled?.planogram ?? rowInfo?.planogram ?? "";
    const colN      = filled?.colN      ?? rowInfo?.colN      ?? "";
    const colPiece  = filled?.colPiece  ?? rowInfo?.colPiece  ?? "";
    const colO      = filled?.colO      ?? rowInfo?.colO      ?? "";

    const spaceman = byUpc.get(barcode);
    const packInfo = barcodeMap.get(barcode);
    const attr = attributeMap.get(barcode);
    const netCapacity = computeNetCapacity(colO, colPiece);

    for (const store of storesWithQty) {
      newItem.push({
        division,
        department: dept,
        upc: barcode,
        name,
        salepack: spaceman?.salepack ?? "",
        recipe: spaceman?.purchaseItemForSalepack ?? "",
        packSize: packInfo?.packSize ?? "",
        totalUnits: spaceman?.totalUnits ?? "",
        purShelfStockPiece: colPiece,
        pctOrdering: colO,
        netCapacity: netCapacity !== null ? String(netCapacity) : "",
        attClass: attr?.attClass ?? "",
        attCode: attr?.attCode ?? "",
        storeNumber: store,
        link: "LINK",
        forecastSalesPerMonthStore: colN,
      });
    }

    if (planogram) {
      const storesShouldHave = indexLookup.pogToStores.get(planogram);
      if (storesShouldHave) {
        for (const store of storesShouldHave) {
          if (storesWithQty.has(store)) continue; // already reported as LINK above
          newNotLink.push({
            upc: barcode,
            name,
            division,
            department: dept,
            attClass: attr?.attClass ?? "",
            attCode: attr?.attCode ?? "",
            storeNumber: store,
            link: "New not link",
          });
        }
      }
    }
  }

  // ── Recap_Delete_item ─────────────────────────────────────────────────────
  for (const [barcode, stores] of delScmStoreFlags) {
    const rowInfo = delScmRowInfo.get(barcode);
    const attr = attributeMap.get(barcode);

    // DEL SCM has no DEPARTMENT column — derive it via the same 100-ช่อง → structureMap
    // join processRows() uses for NEW SCM, when the barcode happens to be in that lookup.
    const subclassCode = barcodeMap.get(barcode)?.subclassCode;
    const hierarchy = subclassCode ? structureMap.get(subclassCode) : undefined;
    const department = hierarchy ? buildRecapCodes(hierarchy).dept : "";

    for (const store of stores) {
      deleteItem.push({
        upc: barcode,
        name: rowInfo?.name ?? "",
        division: rowInfo?.division ?? "",
        department,
        attClass: attr?.attClass ?? "",
        attCode: attr?.attCode ?? "",
        storeNumber: store,
        link: "NOT LINK",
        remark: rowInfo?.remark ?? "", // pass-through verbatim — no normalization (doc §6.7)
      });
    }
  }

  return { newItem, newNotLink, deleteItem };
}
