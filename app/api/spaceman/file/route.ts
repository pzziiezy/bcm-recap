import { NextRequest, NextResponse } from "next/server";

export const runtime = "nodejs";

// Blob URL is public — proxy it through here so callers don't need to know the URL format
export async function GET(request: NextRequest) {
  const url = request.nextUrl.searchParams.get("id");
  if (!url) return NextResponse.json({ error: "No URL provided" }, { status: 400 });

  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`Fetch failed: ${res.status}`);

    const buffer = await res.arrayBuffer();
    return new NextResponse(buffer, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
