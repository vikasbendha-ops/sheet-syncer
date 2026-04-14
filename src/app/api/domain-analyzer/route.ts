import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { extractSpreadsheetId } from "@/lib/url-parser";
import {
  analyzeDomainsForTabs,
  writeDomainAnalysesToSheets,
} from "@/lib/domain-analyzer";

const AnalyzeSchema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Must be a Google Sheets URL",
    }),
  tabs: z.array(z.string().min(1)).min(1, "Select at least one tab"),
  emailColumn: z.string().default("auto"),
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
    const parsed = AnalyzeSchema.parse(body);
    const spreadsheetId = extractSpreadsheetId(parsed.url);

    const result = await analyzeDomainsForTabs(
      session.refreshToken,
      spreadsheetId,
      parsed.tabs,
      parsed.emailColumn
    );

    const writeResult = await writeDomainAnalysesToSheets(
      session.refreshToken,
      spreadsheetId,
      result.tabs
    );

    const tabErrors = result.tabs
      .filter((t) => t.error)
      .map((t) => `Failed to analyze "${t.tabName}": ${t.error}`);

    return NextResponse.json({
      success: writeResult.errors.length === 0 && tabErrors.length === 0,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
      writtenTabs: writeResult.writtenTabs,
      errors: [...tabErrors, ...writeResult.errors],
    });
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
        error: err instanceof Error ? err.message : "Failed to analyze domains",
      },
      { status: 500 }
    );
  }
}
