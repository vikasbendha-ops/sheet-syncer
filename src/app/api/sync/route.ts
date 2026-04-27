import { NextResponse } from "next/server";
import { runFullSync } from "@/lib/sync-engine";
import { getSession } from "@/lib/session";

export async function POST() {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        {
          success: false,
          sheetsProcessed: 0,
          totalEmails: 0,
          errors: ["Not authenticated. Please sign in with Google first."],
          timestamp: new Date().toISOString(),
        },
        { status: 401 }
      );
    }
    if (!session.masterSheetId) {
      return NextResponse.json(
        {
          success: false,
          sheetsProcessed: 0,
          totalEmails: 0,
          errors: ["No master sheet configured. Set one from the Sheets page."],
          timestamp: new Date().toISOString(),
        },
        { status: 400 }
      );
    }

    const result = await runFullSync(session.refreshToken, session.masterSheetId);
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
