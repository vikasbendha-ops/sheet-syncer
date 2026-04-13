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
  tabs: z.array(z.string().min(1)).min(1, "Select at least one tab"),
  emailColumn: z.string().default("auto"),
});

async function getRefreshToken(): Promise<string> {
  const session = await getSession();
  if (!session.refreshToken) {
    throw new Error("Not authenticated. Please sign in with Google first.");
  }
  return session.refreshToken;
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
    const existing = await getLinkedSheets(refreshToken);

    const added: string[] = [];
    const errors: string[] = [];

    for (const tab of parsed.tabs) {
      // Check for duplicates (same URL + same tab)
      if (existing.some((s) => s.url === parsed.url && s.tabName === tab)) {
        errors.push(`"${tab}" is already linked`);
        continue;
      }

      try {
        await readEmailsFromSheet(refreshToken, spreadsheetId, tab, parsed.emailColumn);
      } catch (err) {
        errors.push(`Cannot read "${tab}": ${err instanceof Error ? err.message : String(err)}`);
        continue;
      }

      const nickname = parsed.tabs.length === 1
        ? parsed.nickname
        : `${parsed.nickname} - ${tab}`;
      await addLinkedSheet(refreshToken, parsed.url, nickname, tab, parsed.emailColumn);
      added.push(tab);
    }

    if (added.length === 0) {
      return NextResponse.json(
        { error: errors.join("; ") || "No tabs could be added" },
        { status: 400 }
      );
    }

    return NextResponse.json(
      { success: true, added, errors },
      { status: 201 }
    );
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
