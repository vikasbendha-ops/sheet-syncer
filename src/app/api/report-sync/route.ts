import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { runReportSync } from "@/lib/report-sync";

const SheetsUrl = z
  .string()
  .url()
  .refine((url) => url.includes("docs.google.com/spreadsheets"), {
    message: "URL must be a Google Sheets URL",
  });

const Schema = z.object({
  sourceUrl: SheetsUrl,
  sourceTabs: z.array(z.string().min(1)).min(1, "Select at least one source tab"),
  destinations: z
    .array(
      z.object({
        url: SheetsUrl,
        tabs: z
          .array(z.string().min(1))
          .min(1, "Select at least one destination tab"),
      })
    )
    .min(1, "Add at least one destination spreadsheet"),
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
    const destinations = parsed.destinations.map((d) => ({
      spreadsheetId: extractSpreadsheetId(d.url),
      tabs: d.tabs,
    }));

    const result = await runReportSync(
      session.refreshToken,
      sourceId,
      parsed.sourceTabs,
      destinations
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
        error: err instanceof Error ? err.message : "Failed to run Report sync",
      },
      { status: 500 }
    );
  }
}
