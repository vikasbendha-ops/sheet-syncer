import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { runDuplicateFinder } from "@/lib/duplicate-finder";

const Schema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "URL must be a Google Sheets URL",
    }),
  tabs: z.array(z.string().min(1)).min(1, "Select at least one tab"),
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
    const spreadsheetId = extractSpreadsheetId(parsed.url);

    const result = await runDuplicateFinder(
      session.refreshToken,
      spreadsheetId,
      parsed.tabs
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
        error:
          err instanceof Error ? err.message : "Failed to scan for duplicates",
      },
      { status: 500 }
    );
  }
}
