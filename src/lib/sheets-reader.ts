import { getSheetsClient } from "./google-auth";
import {
  normalizeEmail,
  findEmailColumnIndex,
  findNameColumns,
  columnIndexToLetter,
} from "./email-utils";
import { withRetry } from "./retry";
import { EmailReadResult } from "@/types";
import { sheets_v4 } from "googleapis";

const PRESENT_IN_HEADER = "Present In";

type TabInfo = sheets_v4.Schema$Sheet;

export async function readEmailsFromSheet(
  refreshToken: string,
  spreadsheetId: string,
  tabName: string,
  emailColumnHint: string = "auto",
  cachedTabs?: TabInfo[]
): Promise<EmailReadResult> {
  const sheets = getSheetsClient(refreshToken);

  // Resolve tab + get sheetId (skip API call if caller provided cached tabs)
  let allTabs: TabInfo[];
  if (cachedTabs) {
    allTabs = cachedTabs;
  } else {
    const spreadsheet = await withRetry(() =>
      sheets.spreadsheets.get({ spreadsheetId })
    );
    allTabs = spreadsheet.data.sheets ?? [];
  }

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

  // Read the header row (row 1) in one API call
  const headerResponse = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeTab}!1:1`,
    })
  );

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
    const upper = emailColumnHint.toUpperCase();
    emailColIndex = 0;
    for (let i = 0; i < upper.length; i++) {
      emailColIndex = emailColIndex * 26 + (upper.charCodeAt(i) - 64);
    }
    emailColIndex -= 1;
  }

  const columnLetter = columnIndexToLetter(emailColIndex);

  const presentInColumnIndex = headers.findIndex(
    (h) => h?.toString().trim().toLowerCase() === PRESENT_IN_HEADER.toLowerCase()
  );

  let lastColumnIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]?.toString().trim()) lastColumnIndex = i;
  }

  const nameCols = findNameColumns(headers);

  // Build batch ranges for data (email + optional name columns)
  const ranges: string[] = [`${safeTab}!${columnLetter}2:${columnLetter}`];
  const rangeIdx: {
    email: number;
    fullName?: number;
    firstName?: number;
    lastName?: number;
  } = { email: 0 };

  if (nameCols.fullName !== undefined) {
    rangeIdx.fullName = ranges.length;
    const letter = columnIndexToLetter(nameCols.fullName);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }
  if (nameCols.firstName !== undefined) {
    rangeIdx.firstName = ranges.length;
    const letter = columnIndexToLetter(nameCols.firstName);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }
  if (nameCols.lastName !== undefined) {
    rangeIdx.lastName = ranges.length;
    const letter = columnIndexToLetter(nameCols.lastName);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }

  const batchResponse = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges,
    })
  );

  const valueRanges = batchResponse.data.valueRanges ?? [];
  const emailRows = valueRanges[rangeIdx.email]?.values ?? [];
  const fullNameRows =
    rangeIdx.fullName !== undefined
      ? valueRanges[rangeIdx.fullName]?.values ?? []
      : [];
  const firstNameRows =
    rangeIdx.firstName !== undefined
      ? valueRanges[rangeIdx.firstName]?.values ?? []
      : [];
  const lastNameRows =
    rangeIdx.lastName !== undefined
      ? valueRanges[rangeIdx.lastName]?.values ?? []
      : [];

  const emails = new Map<string, number[]>();
  const names = new Map<string, string[]>();

  for (let i = 0; i < emailRows.length; i++) {
    const rawEmail = emailRows[i]?.[0];
    if (!rawEmail) continue;
    const normalized = normalizeEmail(String(rawEmail));
    if (!normalized) continue;

    const rowNumber = i + 2;
    const existing = emails.get(normalized);
    if (existing) {
      existing.push(rowNumber);
    } else {
      emails.set(normalized, [rowNumber]);
    }

    // Build name for this row
    const full = fullNameRows[i]?.[0];
    const first = firstNameRows[i]?.[0];
    const last = lastNameRows[i]?.[0];
    let name = "";
    if (full && String(full).trim()) {
      name = String(full).trim();
    } else {
      const firstStr = first ? String(first).trim() : "";
      const lastStr = last ? String(last).trim() : "";
      // Guard: if first and last are the same (case-insensitive), use only one
      if (
        firstStr &&
        lastStr &&
        firstStr.toLowerCase() === lastStr.toLowerCase()
      ) {
        name = firstStr;
      } else if (firstStr || lastStr) {
        name = [firstStr, lastStr].filter(Boolean).join(" ");
      }
    }

    if (name) {
      const list = names.get(normalized);
      if (list) {
        list.push(name);
      } else {
        names.set(normalized, [name]);
      }
    }
  }

  return {
    emails,
    names,
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
  const spreadsheet = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId })
  );
  return (spreadsheet.data.sheets ?? [])
    .map((s) => s.properties?.title ?? "")
    .filter(Boolean);
}
