import { NextRequest, NextResponse } from "next/server";
import { google } from "googleapis";
import { Readable } from "stream";

export const runtime = "nodejs";
export const maxDuration = 60;

const FOLDER_ID = "1jWxdKanbMCpf7pShHWdw1GdRzPDqID6S";

function getAuth() {
  const json = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!json) throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not configured");
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(json),
    scopes: ["https://www.googleapis.com/auth/drive"],
  });
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    if (!file) return NextResponse.json({ error: "No file provided" }, { status: 400 });

    const auth = getAuth();
    const drive = google.drive({ version: "v3", auth });

    const buffer = await file.arrayBuffer();
    const stream = Readable.from(Buffer.from(buffer));

    const response = await drive.files.create({
      requestBody: {
        name: file.name,
        parents: [FOLDER_ID],
      },
      media: {
        mimeType:
          file.type ||
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        body: stream,
      },
      fields: "id,name,createdTime",
    });

    return NextResponse.json(response.data);
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
