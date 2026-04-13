import { NextResponse } from "next/server";
import { z } from "zod";
import { getLinkedSheets, addLinkedSheet } from "@/lib/config-store";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { readEmailsFromSheet } from "@/lib/sheets-reader";

const AddSheetSchema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Must be a Google Sheets URL",
    }),
  nickname: z.string().min(1).max(100),
  emailColumn: z.string().default("auto"),
});

export async function GET() {
  try {
    const sheets = await getLinkedSheets();
    return NextResponse.json({ sheets });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to fetch sheets" },
      { status: 500 }
    );
  }
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const parsed = AddSheetSchema.parse(body);

    // Verify the sheet is accessible
    const spreadsheetId = extractSpreadsheetId(parsed.url);
    try {
      await readEmailsFromSheet(spreadsheetId, parsed.emailColumn);
    } catch (err) {
      return NextResponse.json(
        {
          error: `Cannot access the sheet. Make sure it's shared with the service account. Details: ${
            err instanceof Error ? err.message : String(err)
          }`,
        },
        { status: 400 }
      );
    }

    // Check for duplicates
    const existing = await getLinkedSheets();
    if (existing.some((s) => s.url === parsed.url)) {
      return NextResponse.json(
        { error: "This sheet is already linked" },
        { status: 409 }
      );
    }

    await addLinkedSheet(parsed.url, parsed.nickname, parsed.emailColumn);
    return NextResponse.json({ success: true }, { status: 201 });
  } catch (err) {
    if (err instanceof z.ZodError) {
      return NextResponse.json(
        { error: err.issues.map((e: { message: string }) => e.message).join(", ") },
        { status: 400 }
      );
    }
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to add sheet" },
      { status: 500 }
    );
  }
}
