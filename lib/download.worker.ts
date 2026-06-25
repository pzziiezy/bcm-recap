import * as XLSX from "xlsx";
import type { DownloadRow } from "./download";

const workerScope = self as typeof self & {
  postMessage: (message: unknown, transfer?: Transferable[]) => void;
};

type InitMessage = {
  type: "init";
  buffer: ArrayBuffer;
};

type BuildMessage = {
  type: "build";
  jobId: number;
  rows: DownloadRow[];
};

type WorkerMessage = InitMessage | BuildMessage;

let templateBuffer: ArrayBuffer | null = null;

function applyRows(
  ws: XLSX.WorkSheet,
  rows: DownloadRow[]
): void {
  for (const row of rows) {
    const data = row.override
      ? { ...(row.filled ?? {}), ...row.override }
      : row.filled;
    if (!data) continue;

    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 5 })] = {
      t: "s",
      v: data.division ?? "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 6 })] = {
      t: "s",
      v: data.dept ?? "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 7 })] = {
      t: "s",
      v: data.subDept ?? "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 8 })] = {
      t: "s",
      v: data.cls ?? "",
    };
    ws[XLSX.utils.encode_cell({ r: row.rowIndex, c: 9 })] = {
      t: "s",
      v: data.planogram ?? "",
    };
  }
}

addEventListener("message", (event: MessageEvent<WorkerMessage>) => {
  const message = event.data;

  if (message.type === "init") {
    templateBuffer = message.buffer;
    workerScope.postMessage({ type: "init", ok: true });
    return;
  }

  if (!templateBuffer) {
    workerScope.postMessage({
      type: "build",
      jobId: message.jobId,
      ok: false,
      error: "Template workbook is not initialized",
    });
    return;
  }

  try {
    const wb = XLSX.read(templateBuffer, { type: "array" });
    const ws = wb.Sheets["NEW SCM"];
    if (!ws) throw new Error('Sheet "NEW SCM" not found');

    applyRows(ws, message.rows);

    const output = XLSX.write(wb, {
      type: "array",
      bookType: "xlsx",
    }) as ArrayBuffer;

    workerScope.postMessage(
      { type: "build", jobId: message.jobId, ok: true, buffer: output },
      [output]
    );
  } catch (error) {
    workerScope.postMessage({
      type: "build",
      jobId: message.jobId,
      ok: false,
      error: String(error),
    });
  }
});
