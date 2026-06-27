import { NextResponse } from "next/server";
import { list } from "@vercel/blob";

export const runtime = "nodejs";

export async function GET() {
  try {
    const { blobs } = await list({ prefix: "spaceman/" });

    if (blobs.length === 0) return NextResponse.json({ file: null });

    const latest = blobs.sort(
      (a, b) => new Date(b.uploadedAt).getTime() - new Date(a.uploadedAt).getTime()
    )[0];

    return NextResponse.json({
      file: {
        id: latest.url,
        name: latest.pathname.replace("spaceman/", "").replace(/-[a-z0-9]+(\.\w+)$/, "$1"),
        createdTime: latest.uploadedAt,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
