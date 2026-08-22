/**
 * Minor Report build worker — fills the user-uploaded `TO BE_Minor Report.xlsx` template
 * with freshly computed data via the same byte-patch technique as download.worker.ts
 * (unzip → patch only the 3 sheet XMLs + sharedStrings.xml → rezip), so every bit of the
 * template's formatting (colors, borders, column widths, header rows) survives untouched.
 *
 * Unlike RECAP's worker (which upserts additively onto a fixed set of pre-existing rows),
 * this WIPES and rebuilds the data region of each sheet on every build — Minor Report is
 * fully regenerated from the current snapshot each time, so no row count assumption about
 * the template holds run to run. See xlsxPatch.ts's replaceDataRows() for why.
 */

import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import {
  parseSST,
  appendSST,
  buildSST,
  findSheetPath,
  replaceDataRows,
} from "./xlsxPatch";
import type { MinorReportFillPlan } from "./minorReportDownload";
import { MINOR_REPORT_SHEET_NAMES, MINOR_REPORT_HEADER_ROW_COUNT } from "./minorReportDownload";

const ctx = self as unknown as {
  postMessage(msg: unknown, transfer?: Transferable[]): void;
};

type InMsg =
  | { type: "init"; buffer: ArrayBuffer }
  | { type: "build"; plan: MinorReportFillPlan };

let template: ArrayBuffer | null = null;

const SHEET_NAMES = Object.values(MINOR_REPORT_SHEET_NAMES);

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
      ctx.postMessage({ type: "progress", pct: 5 });

      // 1. Unzip the template (it's just a ZIP)
      const files = unzipSync(new Uint8Array(template));
      ctx.postMessage({ type: "progress", pct: 15 });

      const wbXml   = strFromU8(files["xl/workbook.xml"]);
      const relsXml = strFromU8(files["xl/_rels/workbook.xml.rels"]);

      const sstPath = "xl/sharedStrings.xml";
      const sstStrings = files[sstPath] ? parseSST(strFromU8(files[sstPath])) : [];
      const workingStrings = [...sstStrings];

      // 2. Patch each of the 3 sheets in turn
      const pctPerSheet = 60 / SHEET_NAMES.length;
      let pct = 15;
      for (const sheetName of SHEET_NAMES) {
        const sheetPath = findSheetPath(wbXml, relsXml, sheetName);
        if (!sheetPath || !files[sheetPath]) {
          throw new Error(`ไม่พบชีต "${sheetName}" ในไฟล์ template — ตรวจสอบว่าอัปโหลดไฟล์ TO BE_Minor Report.xlsx ที่ถูกต้อง`);
        }
        const rows = msg.plan[sheetName as keyof MinorReportFillPlan];
        const { sheetXml, newStrings } = replaceDataRows(
          strFromU8(files[sheetPath]),
          MINOR_REPORT_HEADER_ROW_COUNT[sheetName],
          rows,
          workingStrings
        );
        files[sheetPath] = strToU8(sheetXml);
        workingStrings.push(...newStrings);

        pct += pctPerSheet;
        ctx.postMessage({ type: "progress", pct: Math.round(pct) });
      }

      // 3. Append all new strings to the SST in one pass
      const allNewStrings = workingStrings.slice(sstStrings.length);
      if (allNewStrings.length > 0) {
        files[sstPath] = strToU8(
          files[sstPath]
            ? appendSST(strFromU8(files[sstPath]), allNewStrings)
            : buildSST([...sstStrings, ...allNewStrings])
        );
      }
      ctx.postMessage({ type: "progress", pct: 90 });

      // 4. Rezip — styles.xml, workbook.xml, relationships, etc. unchanged
      const out = zipSync(files);
      const outBuf = out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
      ctx.postMessage({ type: "progress", pct: 100 });
      ctx.postMessage({ type: "done", buffer: outBuf }, [outBuf]);

    } catch (err) {
      ctx.postMessage({ type: "error", message: String(err) });
    }
  }
});
