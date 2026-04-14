import { getSheetsClient } from "./google-auth";
import {
  normalizeEmail,
  findEmailColumnIndex,
  columnIndexToLetter,
} from "./email-utils";
import { EmailReadResult } from "@/types";

const PRESENT_IN_HEADER = "Present In";

export async function readEmailsFromSheet(
  refreshToken: string,
  spreadsheetId: string,
  tabName: string,
  emailColumnHint: string = "auto"
): Promise<EmailReadResult> {
  const sheets = getSheetsClient(refreshToken);

  // Resolve tab + get sheetId
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const allTabs = spreadsheet.data.sheets ?? [];

  let resolvedTab = tabName;
  let sheetId: number | undefined;

  if (resolvedTab) {
    const match = allTabs.find((t) => t.properties?.title === resolvedTab);
    sheetId = match?.properties?.sheetId ?? undefined;
  }

  if (!resolvedTab || sheetId === undefined) {
    const firstTab = allTabs[0];
    resolvedTab = firstTab?.properties?.title ?? "Sheet1";
    sheetId = firstTab?.properties?.sheetId ?? 0;
  }

  const safeTab = `'${resolvedTab.replace(/'/g, "''")}'`;

  // Read full header row to find email column, Present In column, and last filled column
  const headerResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${safeTab}!1:1`,
  });

  const headers = (headerResponse.data.values?.[0] ?? []) as string[];

  let emailColIndex: number;
  if (emailColumnHint === "auto") {
    emailColIndex = findEmailColumnIndex(headers);
    if (emailColIndex === -1) {
      throw new Error(
        `No email column found in "${resolvedTab}" of sheet ${spreadsheetId}. Headers: ${headers.join(", ")}`
      );
    }
  } else {
    // Convert letter to index (A=0, B=1, ...)
    const upper = emailColumnHint.toUpperCase();
    emailColIndex = 0;
    for (let i = 0; i < upper.length; i++) {
      emailColIndex = emailColIndex * 26 + (upper.charCodeAt(i) - 64);
    }
    emailColIndex -= 1;
  }

  const columnLetter = columnIndexToLetter(emailColIndex);

  // Find existing "Present In" column
  const presentInColumnIndex = headers.findIndex(
    (h) => h?.toString().trim().toLowerCase() === PRESENT_IN_HEADER.toLowerCase()
  );

  // Last filled header column index (0-based)
  let lastColumnIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]?.toString().trim()) lastColumnIndex = i;
  }

  // Read email column data (starting from row 2)
  const dataResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${safeTab}!${columnLetter}2:${columnLetter}`,
  });

  const rows = dataResponse.data.values ?? [];
  const emails = new Map<string, number[]>();

  for (let i = 0; i < rows.length; i++) {
    const raw = rows[i]?.[0];
    if (!raw) continue;
    const normalized = normalizeEmail(String(raw));
    if (!normalized) continue;
    const rowNumber = i + 2; // 1-based, +1 for header
    const existing = emails.get(normalized);
    if (existing) {
      existing.push(rowNumber);
    } else {
      emails.set(normalized, [rowNumber]);
    }
  }

  return {
    emails,
    columnLetter,
    presentInColumnIndex: presentInColumnIndex >= 0 ? presentInColumnIndex : null,
    lastColumnIndex,
    sheetId,
    resolvedTabName: resolvedTab,
  };
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
