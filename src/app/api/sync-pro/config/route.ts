// src/app/api/sync-pro/config/route.ts
import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import {
  getSyncProConfig,
  saveSyncProConfig,
  clearSyncProConfig,
} from "@/lib/sync-pro-config-store";

const LinkedSheetSchema = z.object({
  url: z.string(),
  nickname: z.string(),
  tabName: z.string(),
  emailColumn: z.string(),
  columnMapping: z.record(z.string(), z.union([z.string(), z.null()])),
});

const SectionSchema = z.object({
  id: z.string(),
  name: z.string(),
  masterTabName: z.string(),
  writePresentIn: z.boolean(),
  propagateColumns: z.array(z.object({ name: z.string() })),
  linkedSheets: z.array(LinkedSheetSchema),
});

const Schema = z.object({
  sections: z.array(SectionSchema),
});

async function requireSession() {
  const session = await getSession();
  if (!session.refreshToken) {
    return {
      error: NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      ),
    };
  }
  if (!session.masterSheetId) {
    return {
      error: NextResponse.json(
        {
          error:
            "No master sheet configured. Set one on the Sheets page first.",
          code: "no_master_sheet",
        },
        { status: 400 }
      ),
    };
  }
  return {
    refreshToken: session.refreshToken,
    masterSheetId: session.masterSheetId,
  };
}

export async function GET() {
  const session = await requireSession();
  if ("error" in session) return session.error;
  try {
    const config = await getSyncProConfig(
      session.refreshToken,
      session.masterSheetId
    );
    return NextResponse.json(config);
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to load config" },
      { status: 500 }
    );
  }
}

export async function PUT(request: Request) {
  const session = await requireSession();
  if ("error" in session) return session.error;
  try {
    const body = await request.json();
    const parsed = Schema.parse(body);
    await saveSyncProConfig(
      session.refreshToken,
      session.masterSheetId,
      parsed
    );
    return NextResponse.json({ ok: true });
  } catch (err) {
    if (err instanceof z.ZodError) {
      return NextResponse.json(
        {
          error: err.issues
            .map((e: { message: string }) => e.message)
            .join(", "),
        },
        { status: 400 }
      );
    }
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to save config" },
      { status: 500 }
    );
  }
}

export async function DELETE() {
  const session = await requireSession();
  if ("error" in session) return session.error;
  try {
    await clearSyncProConfig(session.refreshToken, session.masterSheetId);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to clear config" },
      { status: 500 }
    );
  }
}
