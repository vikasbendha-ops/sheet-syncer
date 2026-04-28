import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { findEmailsForNames } from "@/lib/email-finder";

const Schema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Must be a Google Sheets URL",
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
    if (!session.masterSheetId) {
      return NextResponse.json(
        {
          error:
            "No master sheet configured. Set one on the Sheets page first (and run a Sync at least once).",
        },
        { status: 400 }
      );
    }

    const body = await request.json();
    const parsed = Schema.parse(body);
    const targetSpreadsheetId = extractSpreadsheetId(parsed.url);

    const result = await findEmailsForNames(
      session.refreshToken,
      session.masterSheetId,
      targetSpreadsheetId,
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
        error: err instanceof Error ? err.message : "Failed to find emails",
      },
      { status: 500 }
    );
  }
}
