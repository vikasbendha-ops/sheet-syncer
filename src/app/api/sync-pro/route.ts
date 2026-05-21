// src/app/api/sync-pro/route.ts
import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { runSyncProBatch } from "@/lib/sync-pro-engine";

const LinkedSheetSchema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Linked sheet URL must be a Google Sheets URL",
    }),
  nickname: z.string().min(1, "Linked sheet nickname is required"),
  tabName: z.string().min(1, "Linked sheet tab name is required"),
  emailColumn: z.string(),
  columnMapping: z.record(z.string(), z.union([z.string(), z.null()])),
});

const SectionSchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1, "Section name is required"),
  masterTabName: z.string().min(1, "Master tab name is required"),
  writePresentIn: z.boolean(),
  propagateColumns: z.array(z.object({ name: z.string().min(1) })),
  linkedSheets: z
    .array(LinkedSheetSchema)
    .min(1, "Section needs at least one linked sheet"),
});

const Schema = z.object({
  sections: z.array(SectionSchema).min(1, "Provide at least one section"),
});

export async function POST(request: Request) {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated. Please sign in with Google first." },
        { status: 401 }
      );
    }
    if (!session.masterSheetId) {
      return NextResponse.json(
        {
          error:
            "No master sheet configured. Set one on the Sheets page first.",
        },
        { status: 400 }
      );
    }

    const body = await request.json();
    const parsed = Schema.parse(body);

    const result = await runSyncProBatch(
      session.refreshToken,
      session.masterSheetId,
      parsed.sections
    );
    return NextResponse.json(result);
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
      {
        error: err instanceof Error ? err.message : "Failed to run Sync Pro",
      },
      { status: 500 }
    );
  }
}
