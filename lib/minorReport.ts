import type {
  PipelineSnapshot,
  ProcessedRow,
  FilledData,
  StoreFlagMap,
  MinorReportSheets,
  MinorReportNewItemRow,
  MinorReportNewNotLinkRow,
  MinorReportDeleteItemRow,
} from "./types";
import type { CheckSpaceFillPlan } from "./download";
import { computeNetCapacity } from "./netCapacity";
import { buildRecapCodes } from "./processor";

/** Step 5 preview edits win over the originally-computed values, same merge pattern
 *  used by download.worker.ts's applyRows() and page.tsx's fillTabs sync. */
function mergedFilled(row: ProcessedRow): Partial<FilledData> | undefined {
  if (!row.filled && !row.override) return undefined;
  return { ...(row.filled ?? {}), ...(row.override ?? {}) };
}

const NEW_SCM_BARCODE_COL = 3; // matches fillNewSCM()'s BARCODE_COL

/**
 * Barcode → set of store codes flagged in the NEW SCM rows Check Space inserted THIS
 * session only (checkSpacePlan.newScmRows) — as opposed to newScmStoreFlags, which is
 * cumulative across every row for that barcode (this session + any prior session).
 *
 * This distinction only matters for a "NEW EXPAND" item: an item that already sells in
 * some stores and is expanding to more. Confirmed with the user: expanding an existing
 * item's coverage always inserts a brand-new NEW SCM row via Check Space — the old row
 * (with its already-active stores) is never edited directly — so the new row's own
 * store-flag cells are exactly "what's being added this round", cleanly separate from
 * the historical row. For a brand-new item (no prior row at all), this set is identical
 * to newScmStoreFlags's, so NEW ADD ALL/SOME STORE scenarios are unaffected.
 */
function extractThisRoundStoreFlags(
  checkSpacePlan: CheckSpaceFillPlan | null,
  storeColMap: Map<number, string>
): StoreFlagMap {
  const flags: StoreFlagMap = new Map();
  if (!checkSpacePlan) return flags;

  for (const row of checkSpacePlan.newScmRows) {
    const barcodeCell = row.cells.find(c => c.col === NEW_SCM_BARCODE_COL);
    if (!barcodeCell || !barcodeCell.value) continue;
    const barcode = barcodeCell.value;

    for (const cell of row.cells) {
      const storeCode = storeColMap.get(cell.col);
      if (!storeCode || !cell.value) continue;
      if (!flags.has(barcode)) flags.set(barcode, new Set());
      flags.get(barcode)!.add(storeCode);
    }
  }

  return flags;
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
    results, checkSpacePlan, indexLookup, barcodeMap, structureMap, byUpc,
    newScmStoreFlags, delScmStoreFlags, newScmRowInfo, delScmRowInfo, attributeMap,
    newScmStoreColMap,
  } = snapshot;

  const resultsByBarcode = new Map<string, ProcessedRow>();
  for (const row of results) {
    if (!resultsByBarcode.has(row.barcode)) resultsByBarcode.set(row.barcode, row);
  }

  // Only barcodes Check Space actually touched this session drive Recap_New_item /
  // Recap_New_not_link — Minor Report reports "what's new this round", not the full
  // historical contents of NEW SCM (see extractThisRoundStoreFlags() above).
  const thisRoundStoreFlags = extractThisRoundStoreFlags(checkSpacePlan, newScmStoreColMap);

  const newItem: MinorReportNewItemRow[] = [];
  const newNotLink: MinorReportNewNotLinkRow[] = [];
  const deleteItem: MinorReportDeleteItemRow[] = [];

  // ── Recap_New_item + Recap_New_not_link ──────────────────────────────────
  for (const [barcode, storesAddedThisRound] of thisRoundStoreFlags) {
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

    for (const store of storesAddedThisRound) {
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

    // "Not linked" excludes every CUMULATIVE currently-active store for this barcode
    // (historical + this round combined — newScmStoreFlags, not storesAddedThisRound),
    // measured against the FULL store master in FILE_INDEX (indexLookup.storeList —
    // every Store Code column, regardless of POG). Both confirmed with the user against
    // a real INDEX BIG C mini file screenshot and the NEW EXPAND scenario table.
    const allActiveStores = newScmStoreFlags.get(barcode) ?? storesAddedThisRound;
    if (planogram && indexLookup.storeList.length > 0) {
      for (const store of indexLookup.storeList) {
        if (allActiveStores.has(store)) continue; // already selling there — LINK or pre-existing
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
