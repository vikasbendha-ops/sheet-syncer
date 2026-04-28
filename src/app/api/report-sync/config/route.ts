import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import {
  getReportConfig,
  saveReportConfig,
  clearReportConfig,
} from "@/lib/report-config-store";

const Schema = z.object({
  sourceUrl: z.string(),
  sourceTabs: z.array(z.string()),
  destinations: z.array(
    z.object({
      url: z.string(),
      tabs: z.array(z.string()),
    })
  ),
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
    const config = await getReportConfig(
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
    await saveReportConfig(
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
    await clearReportConfig(session.refreshToken, session.masterSheetId);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to clear config" },
      { status: 500 }
    );
  }
}
