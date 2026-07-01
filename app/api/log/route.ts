import { NextResponse } from "next/server";
import { google } from "googleapis";
import type { LogEntry } from "@/lib/logger";

export const runtime = "nodejs";

const SPREADSHEET_ID = "1cZMKuiOOgBvHCUVtWpd-wUCdAWo2u21XDg6FlE6z5Zg";
const LOG_SHEET = "app_log";
const LOG_HEADERS = ["timestamp", "sessionId", "event", "level", "message", "detail"];

function getSheetsClient() {
  const json = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!json) throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not configured");
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(json),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

async function ensureLogSheet(sheets: ReturnType<typeof getSheetsClient>) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets?.some((s) => s.properties?.title === LOG_SHEET);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: LOG_SHEET } } }] },
    });
    // Write header row on first creation
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${LOG_SHEET}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [LOG_HEADERS] },
    });
  }
}

export async function POST(req: Request) {
  try {
    const body = await req.json() as { entries: LogEntry[] };
    const entries = body.entries ?? [];
    if (entries.length === 0) return NextResponse.json({ ok: true });

    const sheets = getSheetsClient();
    await ensureLogSheet(sheets);

    const rows = entries.map((e) => [
      e.timestamp,
      e.sessionId,
      e.event,
      e.level,
      e.message,
      e.detail && Object.keys(e.detail).length > 0 ? JSON.stringify(e.detail) : "",
    ]);

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${LOG_SHEET}!A:F`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: rows },
    });

    return NextResponse.json({ ok: true });
  } catch (err) {
    // Return error but don't crash — client ignores log failures anyway
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
