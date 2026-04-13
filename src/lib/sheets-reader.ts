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
  tabName: string,
  emailColumnHint: string = "auto"
): Promise<EmailReadResult> {
  const sheets = getSheetsClient(refreshToken);

  // If no tab specified, resolve to the first tab in the spreadsheet
  let resolvedTab = tabName;
  if (!resolvedTab) {
    const tabNames = await getTabNames(refreshToken, spreadsheetId);
    resolvedTab = tabNames[0] ?? "Sheet1";
  }

  // Wrap tab name in single quotes to handle spaces/special chars
  const safeTab = `'${resolvedTab.replace(/'/g, "''")}'`;

  let columnLetter: string;

  if (emailColumnHint === "auto") {
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeTab}!1:1`,
    });

    const headers = (headerResponse.data.values?.[0] ?? []) as string[];
    const emailColIndex = findEmailColumnIndex(headers);

    if (emailColIndex === -1) {
      throw new Error(
        `No email column found in "${resolvedTab}" of sheet ${spreadsheetId}. Headers: ${headers.join(", ")}`
      );
    }

    columnLetter = columnIndexToLetter(emailColIndex);
  } else {
    columnLetter = emailColumnHint.toUpperCase();
  }

  const dataResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${safeTab}!${columnLetter}2:${columnLetter}`,
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

export async function getTabNames(
  refreshToken: string,
  spreadsheetId: string
): Promise<string[]> {
  const sheets = getSheetsClient(refreshToken);
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId,
  });

  return (spreadsheet.data.sheets ?? [])
    .map((s) => s.properties?.title ?? "")
    .filter(Boolean);
}
