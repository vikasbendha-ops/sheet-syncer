import { NextResponse } from "next/server";
import { runFullSync } from "@/lib/sync-engine";

/**
 * Cron-friendly sync endpoint.
 *
 * Authentication: requires `Authorization: Bearer <CRON_SECRET>` header
 * OR `?key=<CRON_SECRET>` query param (some cron services don't support custom headers).
 *
 * Credentials: reads `GOOGLE_REFRESH_TOKEN` and `MASTER_SHEET_ID` from env vars.
 * The user must populate these after their first browser sign-in (see /sheets page).
 */
async function handle(request: Request): Promise<NextResponse> {
  const cronSecret = process.env.CRON_SECRET;
  if (!cronSecret) {
    return NextResponse.json(
      { error: "Server is missing CRON_SECRET env var" },
      { status: 500 }
    );
  }

  const authHeader = request.headers.get("authorization");
  const url = new URL(request.url);
  const queryKey = url.searchParams.get("key");

  if (
    authHeader !== `Bearer ${cronSecret}` &&
    queryKey !== cronSecret
  ) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  const masterSheetId = process.env.MASTER_SHEET_ID;

  if (!refreshToken) {
    return NextResponse.json(
      { error: "GOOGLE_REFRESH_TOKEN env var not set" },
      { status: 500 }
    );
  }
  if (!masterSheetId) {
    return NextResponse.json(
      { error: "MASTER_SHEET_ID env var not set" },
      { status: 500 }
    );
  }

  try {
    const result = await runFullSync(refreshToken, masterSheetId);
    return NextResponse.json(result);
  } catch (err) {
    return NextResponse.json(
      {
        success: false,
        error: err instanceof Error ? err.message : "Cron sync failed",
        timestamp: new Date().toISOString(),
      },
      { status: 500 }
    );
  }
}

export async function GET(request: Request) {
  return handle(request);
}

export async function POST(request: Request) {
  return handle(request);
}
