import { getSheetsClient } from "./google-auth";
import {
  normalizeEmail,
  findEmailColumnIndex,
  columnIndexToLetter,
} from "./email-utils";
import { EmailReadResult } from "@/types";

export async function readEmailsFromSheet(
  refreshToken: string,
  spreadsheetId: string,
  emailColumnHint: string = "auto"
): Promise<EmailReadResult> {
  const sheets = getSheetsClient(refreshToken);

  let columnLetter: string;

  if (emailColumnHint === "auto") {
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "1:1",
    });

    const headers = (headerResponse.data.values?.[0] ?? []) as string[];
    const emailColIndex = findEmailColumnIndex(headers);

    if (emailColIndex === -1) {
      throw new Error(
        `No email column found in sheet ${spreadsheetId}. Headers: ${headers.join(", ")}`
      );
    }

    columnLetter = columnIndexToLetter(emailColIndex);
  } else {
    columnLetter = emailColumnHint.toUpperCase();
  }

  const dataResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${columnLetter}2:${columnLetter}`,
  });

  const rows = dataResponse.data.values ?? [];
  const emails = new Set<string>();

  for (const row of rows) {
    const raw = row[0];
    if (!raw) continue;
    const normalized = normalizeEmail(String(raw));
    if (normalized) {
      emails.add(normalized);
    }
  }

  return { emails, columnLetter };
}
