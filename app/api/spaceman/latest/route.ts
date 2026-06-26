import { NextResponse } from "next/server";
import { google } from "googleapis";

export const runtime = "nodejs";

const FOLDER_ID = "1jWxdKanbMCpf7pShHWdw1GdRzPDqID6S";

function getAuth() {
  const json = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!json) throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not configured");
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(json),
    scopes: ["https://www.googleapis.com/auth/drive.readonly"],
  });
}

export async function GET() {
  try {
    const auth = getAuth();
    const drive = google.drive({ version: "v3", auth });

    const response = await drive.files.list({
      q: `'${FOLDER_ID}' in parents and trashed=false`,
      orderBy: "createdTime desc",
      pageSize: 1,
      fields: "files(id,name,createdTime)",
    });

    const files = response.data.files;
    if (!files || files.length === 0) {
      return NextResponse.json({ file: null });
    }

    return NextResponse.json({ file: files[0] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
