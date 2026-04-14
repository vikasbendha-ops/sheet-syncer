import { NextRequest, NextResponse } from "next/server";
import { getSession } from "@/lib/session";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { getTabNames } from "@/lib/sheets-reader";

export async function GET(request: NextRequest) {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }

    const url = request.nextUrl.searchParams.get("url");
    if (!url) {
      return NextResponse.json(
        { error: "Missing url parameter" },
        { status: 400 }
      );
    }

    const spreadsheetId = extractSpreadsheetId(url);
    const tabs = await getTabNames(session.refreshToken, spreadsheetId);

    return NextResponse.json({ tabs });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to fetch tabs" },
      { status: 500 }
    );
  }
}
