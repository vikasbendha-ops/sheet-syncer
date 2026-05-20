import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { runConsolidatorBatch } from "@/lib/consolidator";

const SourceSchema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Source URL must be a Google Sheets URL",
    }),
  tabs: z.array(z.string().min(1)).min(1, "Select at least one tab"),
});

const SectionSchema = z.object({
  id: z.string().min(1),
  name: z.string(),
  sources: z.array(SourceSchema).min(1, "Section needs at least one source spreadsheet"),
  outputUrl: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Output URL must be a Google Sheets URL",
    }),
  outputTabName: z.string().min(1, "Output tab name is required"),
});

const Schema = z.object({
  sections: z.array(SectionSchema).min(1, "Configure at least one section"),
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

    const body = await request.json();
    const parsed = Schema.parse(body);

    const results = await runConsolidatorBatch(
      session.refreshToken,
      parsed.sections
    );

    return NextResponse.json({ sections: results });
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
        error:
          err instanceof Error ? err.message : "Failed to run consolidator",
      },
      { status: 500 }
    );
  }
}
