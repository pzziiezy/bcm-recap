/**
 * Shared Google Sheets helper for the Config routes.
 *
 * The service account must be given **Editor** access to the spreadsheet below.
 * Share it with the service account email found in GOOGLE_SERVICE_ACCOUNT_JSON.
 */

import { google } from "googleapis";

export const SPREADSHEET_ID = "1cZMKuiOOgBvHCUVtWpd-wUCdAWo2u21XDg6FlE6z5Zg";
export const SHEET_NAME = "Config";

export const HEADERS = [
  "id", "category", "subcategory", "descC", "percentage", "status", "createdAt", "updatedAt", "deletedAt",
] as const;

export function getSheetsClient() {
  const json = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!json) throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not configured");
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(json),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

/** Ensure the "Config" sheet exists; creates it if missing. */
export async function ensureConfigSheet(sheets: ReturnType<typeof getSheetsClient>) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets?.some(
    (s) => s.properties?.title === SHEET_NAME
  );
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: SHEET_NAME } } }],
      },
    });
  }
}
