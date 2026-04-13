import { NextResponse } from "next/server";
import { z } from "zod";
import { getLinkedSheets, addLinkedSheet } from "@/lib/config-store";
import { extractSpreadsheetId } from "@/lib/url-parser";
import { readEmailsFromSheet } from "@/lib/sheets-reader";
import { getSession } from "@/lib/session";

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

async function getRefreshToken(): Promise<string> {
  const session = await getSession();
  const token = session.refreshToken ?? process.env.GOOGLE_REFRESH_TOKEN;
  if (!token) {
    throw new Error("Not authenticated. Please sign in with Google first.");
  }
  return token;
}

export async function GET() {
  try {
    const refreshToken = await getRefreshToken();
    const sheets = await getLinkedSheets(refreshToken);
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
    const refreshToken = await getRefreshToken();
    const body = await request.json();
    const parsed = AddSheetSchema.parse(body);

    const spreadsheetId = extractSpreadsheetId(parsed.url);
    try {
      await readEmailsFromSheet(refreshToken, spreadsheetId, parsed.emailColumn);
    } catch (err) {
      return NextResponse.json(
        {
          error: `Cannot access the sheet. Make sure you have access to it. Details: ${
            err instanceof Error ? err.message : String(err)
          }`,
        },
        { status: 400 }
      );
    }

    const existing = await getLinkedSheets(refreshToken);
    if (existing.some((s) => s.url === parsed.url)) {
      return NextResponse.json(
        { error: "This sheet is already linked" },
        { status: 409 }
      );
    }

    await addLinkedSheet(refreshToken, parsed.url, parsed.nickname, parsed.emailColumn);
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
