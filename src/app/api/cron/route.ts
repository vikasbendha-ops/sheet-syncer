import { NextResponse } from "next/server";
import { runFullSync } from "@/lib/sync-engine";

export async function GET(request: Request) {
  // Verify cron secret
  const authHeader = request.headers.get("authorization");
  const cronSecret = process.env.CRON_SECRET;

  if (cronSecret && authHeader !== `Bearer ${cronSecret}`) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  // Cron uses the env var refresh token since there's no user session
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (!refreshToken) {
    return NextResponse.json(
      {
        success: false,
        error:
          "GOOGLE_REFRESH_TOKEN env var is not set. Sign in via the UI first, then copy your refresh token to the env var.",
      },
      { status: 500 }
    );
  }

  try {
    const result = await runFullSync(refreshToken);
    return NextResponse.json(result);
  } catch (err) {
    return NextResponse.json(
      {
        success: false,
        error: err instanceof Error ? err.message : "Cron sync failed",
      },
      { status: 500 }
    );
  }
}
