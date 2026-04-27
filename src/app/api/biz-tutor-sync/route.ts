import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { runBizTutorSync } from "@/lib/biz-tutor-sync";

const Schema = z.object({
  sourceUrl: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Source URL must be a Google Sheets URL",
    }),
  sourceTabs: z.array(z.string().min(1)).min(1, "Select at least one source tab"),
  lookupUrl: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Lookup URL must be a Google Sheets URL",
    }),
  lookupTabs: z.array(z.string().min(1)).min(1, "Select at least one lookup tab"),
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

    const sourceId = extractSpreadsheetId(parsed.sourceUrl);
    const lookupId = extractSpreadsheetId(parsed.lookupUrl);

    const result = await runBizTutorSync(
      session.refreshToken,
      sourceId,
      parsed.sourceTabs,
      lookupId,
      parsed.lookupTabs
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
        error: err instanceof Error ? err.message : "Failed to run BIZ TUTOR sync",
      },
      { status: 500 }
    );
  }
}
