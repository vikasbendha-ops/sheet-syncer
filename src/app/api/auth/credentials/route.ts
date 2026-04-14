import { NextResponse } from "next/server";
import { getSession } from "@/lib/session";

/**
 * Returns the user's refresh token + master sheet ID so they can paste them
 * into env vars for the cron job.
 */
export async function GET() {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }
    return NextResponse.json({
      refreshToken: session.refreshToken,
      masterSheetId: session.masterSheetId ?? null,
    });
  } catch {
    return NextResponse.json(
      { error: "Failed to get credentials" },
      { status: 500 }
    );
  }
}
