// src/app/api/multi-sync/claim-slot/route.ts
import { NextResponse } from "next/server";
import { getSession } from "@/lib/session";
import { claimSlot } from "@/lib/multi-sync-config-store";

/**
 * POST returns the next monotonic slot number from the master sheet's
 * `_multi_sync_meta` tab. Called by the page when the user clicks
 * "+ Add section" so each new section gets a permanent, never-reused
 * slot id (which drives the default Present In column name + Master
 * tab name).
 */
export async function POST() {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }
    if (!session.masterSheetId) {
      return NextResponse.json(
        {
          error:
            "No master sheet configured. Set one on the Sheets page first.",
          code: "no_master_sheet",
        },
        { status: 400 }
      );
    }
    const slot = await claimSlot(
      session.refreshToken,
      session.masterSheetId
    );
    return NextResponse.json({ slot });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to claim slot" },
      { status: 500 }
    );
  }
}
