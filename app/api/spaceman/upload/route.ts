import { NextRequest, NextResponse } from "next/server";

export const runtime = "nodejs";
export const maxDuration = 60;

const FOLDER_ID = "1jWxdKanbMCpf7pShHWdw1GdRzPDqID6S";

// Upload is performed on behalf of the user using their OAuth access token.
// The token comes from the browser (Google Identity Services) so the file
// is owned by the user — no service-account storage-quota issue.
export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    const accessToken = formData.get("accessToken") as string | null;

    if (!file) return NextResponse.json({ error: "No file provided" }, { status: 400 });
    if (!accessToken) return NextResponse.json({ error: "No access token" }, { status: 401 });

    const metadata = JSON.stringify({ name: file.name, parents: [FOLDER_ID] });

    const body = new FormData();
    body.append("metadata", new Blob([metadata], { type: "application/json" }));
    body.append("file", file);

    const res = await fetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,createdTime",
      {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}` },
        body,
      }
    );

    if (!res.ok) {
      const err = await res.json();
      throw new Error(JSON.stringify(err));
    }

    return NextResponse.json(await res.json());
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
