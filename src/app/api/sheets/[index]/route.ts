import { NextResponse } from "next/server";
import { removeLinkedSheet } from "@/lib/config-store";
import { getSession } from "@/lib/session";

export async function DELETE(
  _request: Request,
  { params }: { params: Promise<{ index: string }> }
) {
  try {
    const session = await getSession();
    const refreshToken = session.refreshToken ?? process.env.GOOGLE_REFRESH_TOKEN;
    if (!refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }

    const { index: indexStr } = await params;
    const index = parseInt(indexStr, 10);
    if (isNaN(index) || index < 0) {
      return NextResponse.json(
        { error: "Invalid index" },
        { status: 400 }
      );
    }

    await removeLinkedSheet(refreshToken, index);
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to remove sheet" },
      { status: 500 }
    );
  }
}
