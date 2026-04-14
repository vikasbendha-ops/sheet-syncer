import { NextResponse } from "next/server";
import { runFullSync } from "@/lib/sync-engine";
import { getSession } from "@/lib/session";

export async function POST() {
  try {
    const session = await getSession();
    const refreshToken = session.refreshToken;
    if (!refreshToken) {
      return NextResponse.json(
        { success: false, error: "Not authenticated. Please sign in with Google first." },
        { status: 401 }
      );
    }

    const result = await runFullSync(refreshToken);
    return NextResponse.json(result);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Sync failed";
    return NextResponse.json(
      {
        success: false,
        sheetsProcessed: 0,
        totalEmails: 0,
        errors: [message],
        timestamp: new Date().toISOString(),
      },
      { status: 500 }
    );
  }
}
