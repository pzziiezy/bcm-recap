import { NextResponse } from "next/server";
import { getSheetsClient, ensureConfigSheet, SPREADSHEET_ID, SHEET_NAME, HEADERS } from "../_sheets";
import type { ExceptionConfig } from "@/lib/types";

export const runtime = "nodejs";

export async function POST(req: Request) {
  try {
    const body = await req.json() as { entries: ExceptionConfig[] };
    const entries = body.entries ?? [];
    const now = new Date().toISOString();

    const sheets = getSheetsClient();
    await ensureConfigSheet(sheets);

    // Stamp updatedAt = now for all entries being saved
    const stamped: ExceptionConfig[] = entries.map((e) => ({
      ...e,
      createdAt: e.createdAt || now,
      updatedAt: now,
    }));

    // Build rows: header + data
    const headerRow = [...HEADERS];
    const dataRows = stamped.map((e) => [
      e.id, e.category, e.subcategory, e.descC, e.percentage, e.status ?? "active", e.createdAt, e.updatedAt, e.deletedAt ?? "",
    ]);

    const values = [headerRow, ...dataRows];
    const range = `${SHEET_NAME}!A1:I${values.length}`;

    // Clear existing content first, then write fresh
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:I`,
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range,
      valueInputOption: "RAW",
      requestBody: { values },
    });

    return NextResponse.json({ ok: true, savedAt: now, entries: stamped });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
