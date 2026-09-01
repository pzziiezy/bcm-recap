import type {
  MinorReportInput,
  SubclassInfo,
  HierarchyNames,
  SpacemanRowMeta,
  ExceptionConfig,
  MinorReportSheets,
  MinorReportNewItemRow,
  MinorReportNewNotLinkRow,
  MinorReportDeleteItemRow,
} from "./types";
import { computeNetCapacity } from "./netCapacity";
import { buildRecapCodes, findMatchingConfig } from "./processor";

const ATT_CLASS_CONST = "MBC1"; // always a fixed constant — verified against real RECAP data

interface Enrichment {
  division: string;
  dept: string;
  colN: string;      // Forecast Sales/Month/Store — 100 ช่อง's colDF
  colPiece: string;  // TOTAL_UNITS / BCM Shelf stock ON POG (Piece) — DATA_SPACEMAN's totalUnits
  colO: string;       // % Ordering — exception-config match, default "100%"
}

/**
 * DIVISION/DEPARTMENT/forecast/%-ordering resolution for one barcode — mirrors the
 * previously-verified processRows() formula exactly (100 ช่อง → structureMap primary,
 * DATA_SPACEMAN descA/descB fallback, exception-config % match with the same derived-meta
 * fallback), just restructured as a per-barcode lookup instead of a per-RECAP-row scan.
 */
function resolveEnrichment(
  barcode: string,
  barcodeMap: Map<string, SubclassInfo>,
  structureMap: Map<string, HierarchyNames>,
  byUpc: Map<string, SpacemanRowMeta>,
  exceptionConfig: ExceptionConfig[]
): Enrichment {
  const info = barcodeMap.get(barcode);
  const spacemanMeta = byUpc.get(barcode);

  if (!info) {
    if (spacemanMeta) {
      const matched = exceptionConfig.length > 0 ? findMatchingConfig(exceptionConfig, spacemanMeta) : null;
      return {
        division: spacemanMeta.descA,
        dept: spacemanMeta.descB,
        colN: "",
        colPiece: spacemanMeta.totalUnits,
        colO: matched ? `${matched.percentage}%` : "100%",
      };
    }
    return { division: "", dept: "", colN: "", colPiece: "", colO: "100%" };
  }

  const hierarchy = structureMap.get(info.subclassCode);
  if (!hierarchy) {
    return {
      division: "",
      dept: "",
      colN: info.colDF || "",
      colPiece: spacemanMeta?.totalUnits || "",
      colO: "100%",
    };
  }

  const filled = buildRecapCodes(hierarchy);
  const colN = info.colDF || "";
  const colPiece = spacemanMeta?.totalUnits || "";

  const derivedMeta: SpacemanRowMeta = {
    category: filled.cls,
    subcategory: info.subclassCode
      ? (info.subclassName.trim() ? `${info.subclassCode}: ${info.subclassName.trim()}` : info.subclassCode)
      : "",
    descA: filled.division,
    descB: filled.dept,
    descC: filled.subDept,
    totalUnits: "",
  };
  const meta = spacemanMeta ?? derivedMeta;
  const matched = exceptionConfig.length > 0 ? findMatchingConfig(exceptionConfig, meta) : null;

  return {
    division: filled.division,
    dept: filled.dept,
    colN,
    colPiece,
    colO: matched ? `${matched.percentage}%` : "100%",
  };
}

/**
 * Reshapes Check Space + FILE_INDEX + 100 ช่อง + DATA_SPACEMAN directly into the 3 Minor
 * Report sheets. Pure function — no RECAP file, no React, no file parsing.
 *
 * Per Check Space item (barcode × ticked POGs):
 *   Recap_New_item     = one row per (barcode × store) for every store the TICKED POGs
 *                         cover in FILE_INDEX — "what's newly targeted this round".
 *   Recap_New_not_link = one row per (barcode × store) in FILE_INDEX's full store master
 *                         that ISN'T already covered by either the ticked POGs (new) or
 *                         the item's DATA_SPACEMAN-recorded current planogram (existing) —
 *                         this single formula naturally handles NEW ADD ALL/SOME STORE
 *                         (no existing planogram yet, so it's just new vs total) and
 *                         NEW EXPAND (existing + new both subtracted) without special-casing
 *                         the status text at all.
 *   Recap_Delete_item  = one row per (barcode × store): "DELETE ALL STORE" explodes over
 *                         the full store master (matches fillDelSCM's old behavior of
 *                         leaving store-flags blank to mean "everywhere"); "DELETE SOME
 *                         STORE" explodes over the ticked POGs' store union, same as NEW.
 */
export function buildMinorReportSheets(input: MinorReportInput): MinorReportSheets {
  const { checkSpaceItems, indexLookup, barcodeMap, structureMap, byUpc, exceptionConfig } = input;

  const newItem: MinorReportNewItemRow[] = [];
  const newNotLink: MinorReportNewNotLinkRow[] = [];
  const deleteItem: MinorReportDeleteItemRow[] = [];

  for (const item of checkSpaceItems) {
    const isDelete = item.status.toUpperCase().startsWith("DELETE");
    const enrichment = resolveEnrichment(item.barcode, barcodeMap, structureMap, byUpc, exceptionConfig);
    const spaceman = byUpc.get(item.barcode);
    const packInfo = barcodeMap.get(item.barcode);
    const netCapacity = computeNetCapacity(enrichment.colO, enrichment.colPiece);
    // Reference ATT_CODE for rows that can't be tied to one specific POG (e.g. a
    // not-linked store, or a "DELETE ALL STORE" store outside any ticked POG) —
    // first ticked POG's code, same "first occurrence wins" simplification used
    // throughout this codebase for one-barcode-many-attributes cases.
    const firstPogByCode = item.pogs.length > 0 ? (indexLookup.pogToByCode.get(item.pogs[0]) ?? "") : "";

    if (!isDelete) {
      // ── Stores already selling this item — union across EVERY planogram DATA_SPACEMAN
      //    has recorded for this barcode, not just one (a barcode can legitimately sell
      //    on several planograms at once). ──
      const existingStores = new Set<string>();
      for (const plog of spaceman?.planograms ?? []) {
        const stores = indexLookup.pogToStores.get(plog);
        if (stores) for (const s of stores) existingStores.add(s);
      }

      // ── Stores newly targeted this round: union of the ticked POGs, remembering
      //    which POG each store came from so ATT_CODE is correct per store. ──
      const newStoreToPog = new Map<string, string>();
      for (const pog of item.pogs) {
        const stores = indexLookup.pogToStores.get(pog);
        if (!stores) continue;
        for (const store of stores) {
          if (!newStoreToPog.has(store)) newStoreToPog.set(store, pog);
        }
      }

      // Recap_New_item lists only stores that are actually NEW to this barcode — for
      // NEW EXPAND, a ticked POG's store that's already active under an existing
      // planogram (ก้อน B) isn't "new" and is skipped here (confirmed with the user).
      // It still counts toward "active" for the not-link exclusion below, though.
      for (const [store, pog] of newStoreToPog) {
        if (existingStores.has(store)) continue;
        newItem.push({
          division: enrichment.division,
          department: enrichment.dept,
          upc: item.barcode,
          name: item.name,
          salepack: spaceman?.salepack ?? "",
          recipe: spaceman?.purchaseItemForSalepack ?? "",
          packSize: packInfo?.packSize ?? "",
          totalUnits: spaceman?.totalUnits ?? "",
          purShelfStockPiece: enrichment.colPiece,
          pctOrdering: enrichment.colO,
          netCapacity: netCapacity !== null ? String(netCapacity) : "",
          attClass: ATT_CLASS_CONST,
          attCode: indexLookup.pogToByCode.get(pog) ?? "",
          storeNumber: store,
          link: "LINK",
          forecastSalesPerMonthStore: enrichment.colN,
          remark: item.status,
        });
      }

      // TEMP: ก้อน C (Recap_New_not_link) disabled while checking A/B correctness — re-enable
      // by uncommenting this block once Sheet 1's data is confirmed right.
      /*
      const cumulativeActive = new Set<string>(existingStores);
      for (const store of newStoreToPog.keys()) cumulativeActive.add(store);

      if (cumulativeActive.size > 0) {
        for (const store of indexLookup.storeList) {
          if (cumulativeActive.has(store)) continue;
          newNotLink.push({
            upc: item.barcode,
            name: item.name,
            division: enrichment.division,
            department: enrichment.dept,
            attClass: ATT_CLASS_CONST,
            attCode: firstPogByCode,
            storeNumber: store,
            link: "New not link",
            remark: item.status,
          });
        }
      }
      */
    } else {
      // ── DELETE ──
      const isDeleteAll = /DELETE\s+ALL/i.test(item.status);
      const deleteStores = new Set<string>();
      if (isDeleteAll) {
        for (const store of indexLookup.storeList) deleteStores.add(store);
      } else {
        for (const pog of item.pogs) {
          const stores = indexLookup.pogToStores.get(pog);
          if (stores) for (const s of stores) deleteStores.add(s);
        }
      }

      for (const store of deleteStores) {
        deleteItem.push({
          upc: item.barcode,
          name: item.name,
          division: enrichment.division,
          department: enrichment.dept,
          attClass: ATT_CLASS_CONST,
          attCode: firstPogByCode,
          storeNumber: store,
          link: "NOT LINK",
          // Verified against fillDelSCM(): DEL SCM's REMARK column was always written
          // as item.status verbatim (not Check Space's own remark field) — pass-through,
          // no normalization (doc §6.7).
          remark: item.status,
        });
      }
    }
  }

  return { newItem, newNotLink, deleteItem };
}
