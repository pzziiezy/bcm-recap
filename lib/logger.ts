export type LogEvent =
  | "PROCESS_START"
  | "RECAP_PARSED"
  | "XLSB_PARSED"
  | "SPACEMAN_PARSED"
  | "PROCESS_COMPLETE"
  | "BUILD_QUEUED"
  | "BUILD_COMPLETE"
  | "BUILD_FAILED"
  | "ERROR";

export type LogLevel = "INFO" | "WARN" | "ERROR";

export interface LogEntry {
  timestamp: string;
  sessionId: string;
  event: LogEvent;
  level: LogLevel;
  message: string;
  detail: Record<string, unknown>;
}

export function makeEntry(
  sessionId: string,
  event: LogEvent,
  level: LogLevel,
  message: string,
  detail: Record<string, unknown> = {}
): LogEntry {
  return { timestamp: new Date().toISOString(), sessionId, event, level, message, detail };
}

/** Fire-and-forget — never throws, never blocks the caller. */
export function sendLog(entries: LogEntry[]): void {
  fetch("/api/log", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ entries }),
  }).catch(() => { /* logging must never crash the app */ });
}
