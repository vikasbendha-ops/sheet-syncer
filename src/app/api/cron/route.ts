import { NextResponse } from "next/server";
import { runFullSync } from "@/lib/sync-engine";

export async function GET(request: Request) {
  // Verify cron secret (Vercel sends this automatically)
  const authHeader = request.headers.get("authorization");
  const cronSecret = process.env.CRON_SECRET;

  if (cronSecret && authHeader !== `Bearer ${cronSecret}`) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  try {
    const result = await runFullSync();
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
