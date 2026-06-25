import * as XLSX from "xlsx";
import { applyRowsToSheet, type DownloadRow } from "./download";

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
    const workbook = XLSX.read(templateBuffer, { type: "array" });
    const ws = workbook.Sheets["NEW SCM"];
    if (!ws) throw new Error('Sheet "NEW SCM" not found');

    applyRowsToSheet(ws, message.rows);

    const output = XLSX.write(workbook, {
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
