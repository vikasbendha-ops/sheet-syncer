import { NextResponse } from "next/server";
import { runFullSync } from "@/lib/sync-engine";

export async function POST() {
  try {
    const result = await runFullSync();
    return NextResponse.json(result);
  } catch (err) {
    return NextResponse.json(
      {
        success: false,
        error: err instanceof Error ? err.message : "Sync failed",
      },
      { status: 500 }
    );
  }
}
