import { NextResponse } from "next/server";
import { getSheetsClient, ensureConfigSheet, SPREADSHEET_ID, SHEET_NAME, HEADERS } from "../_sheets";
import type { ExceptionConfig } from "@/lib/types";

export const runtime = "nodejs";

export async function GET() {
  try {
    const sheets = getSheetsClient();
    await ensureConfigSheet(sheets);

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1:I`,
    });

    const rows = res.data.values ?? [];
    if (rows.length < 2) {
      return NextResponse.json({ config: [], lastSaved: null });
    }

    const headerRow = rows[0];
    const idx = (col: string) => headerRow.indexOf(col);
    const iId      = idx("id");
    const iCat     = idx("category");
    const iSub     = idx("subcategory");
    const iDesc    = idx("descC");
    const iPct     = idx("percentage");
    const iStatus  = idx("status");
    const iCreated = idx("createdAt");
    const iUpdated = idx("updatedAt");
    const iDeleted = idx("deletedAt");

    const rawStatus = (s: string): ExceptionConfig["status"] => {
      if (s === "inactive") return "inactive";
      if (s === "deleted")  return "deleted";
      return "active";
    };

    const config: ExceptionConfig[] = rows.slice(1)
      .filter((r) => r[iId])
      .map((r) => ({
        id:          r[iId]      ?? "",
        category:    r[iCat]     ?? "ทั้งหมด",
        subcategory: r[iSub]     ?? "ทั้งหมด",
        descC:       r[iDesc]    ?? "ทั้งหมด",
        percentage:  r[iPct]     ?? "100",
        status:      rawStatus(r[iStatus] ?? ""),
        createdAt:   r[iCreated] ?? "",
        updatedAt:   r[iUpdated] ?? "",
        ...(iDeleted >= 0 && r[iDeleted] ? { deletedAt: r[iDeleted] } : {}),
      }));

    // The most recent updatedAt across all entries = last save time
    const lastSaved = config.reduce<string | null>((latest, e) => {
      if (!e.updatedAt) return latest;
      return !latest || e.updatedAt > latest ? e.updatedAt : latest;
    }, null);

    return NextResponse.json({ config, lastSaved });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
