import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession, saveSession } from "@/lib/session";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { getSheetsClient } from "@/lib/google-auth";

const SetMasterSchema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Must be a Google Sheets URL",
    }),
});

export async function GET() {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }
    return NextResponse.json({
      masterSheetId: session.masterSheetId ?? null,
    });
  } catch {
    return NextResponse.json({ masterSheetId: null });
  }
}

export async function POST(request: Request) {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }

    const body = await request.json();
    const parsed = SetMasterSchema.parse(body);
    const masterSheetId = extractSpreadsheetId(parsed.url);

    // Verify the user has access to this spreadsheet
    const sheets = getSheetsClient(session.refreshToken);
    try {
      await sheets.spreadsheets.get({
        spreadsheetId: masterSheetId,
        fields: "spreadsheetId",
      });
    } catch (err) {
      return NextResponse.json(
        {
          error: `Cannot access this sheet. Make sure you own it or have edit access. Details: ${
            err instanceof Error ? err.message : String(err)
          }`,
        },
        { status: 400 }
      );
    }

    await saveSession({ ...session, masterSheetId });

    return NextResponse.json({ success: true, masterSheetId });
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
          err instanceof Error ? err.message : "Failed to set master sheet",
      },
      { status: 500 }
    );
  }
}

export async function DELETE() {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }
    const cleaned = { ...session };
    delete cleaned.masterSheetId;
    await saveSession(cleaned);
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json(
      {
        error: err instanceof Error ? err.message : "Failed to clear",
      },
      { status: 500 }
    );
  }
}
