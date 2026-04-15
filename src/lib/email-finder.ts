import { getSheetsClient } from "./google-auth";
import { columnIndexToLetter, findNameColumns } from "./email-utils";
import { withRetry } from "./retry";
import { sheets_v4 } from "googleapis";

const EMAIL_HEADER = "Email";
const RED_BG = { red: 1.0, green: 0.82, blue: 0.82 };
const WHITE_BG = { red: 1.0, green: 1.0, blue: 1.0 };

export interface EmailFinderTabResult {
  tabName: string;
  totalRows: number;
  matched: number;
  unmatched: number;
  error?: string;
}

export interface EmailFinderResult {
  spreadsheetUrl: string;
  tabs: EmailFinderTabResult[];
}

function normalizeName(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, " ");
}

/**
 * Reads the master sheet's Name + Email columns to build a name→email lookup.
 * Master is expected to have Name in column A and Email in column B
 * (matching the output of the main sync engine).
 */
async function loadMasterNameMap(
  sheets: sheets_v4.Sheets,
  masterSheetId: string
): Promise<Map<string, string>> {
  const res = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId: masterSheetId,
      ranges: ["Master!A2:A", "Master!B2:B"],
    })
  );

  const nameRows = res.data.valueRanges?.[0]?.values ?? [];
  const emailRows = res.data.valueRanges?.[1]?.values ?? [];

  const map = new Map<string, string>();
  const rowCount = Math.max(nameRows.length, emailRows.length);
  for (let i = 0; i < rowCount; i++) {
    const rawName = nameRows[i]?.[0];
    const rawEmail = emailRows[i]?.[0];
    if (!rawName || !rawEmail) continue;
    const key = normalizeName(String(rawName));
    if (!key) continue;
    if (!map.has(key)) {
      map.set(key, String(rawEmail).trim());
    }
  }
  return map;
}

export async function findEmailsForNames(
  refreshToken: string,
  masterSheetId: string,
  targetSpreadsheetId: string,
  tabs: string[]
): Promise<EmailFinderResult> {
  const sheets = getSheetsClient(refreshToken);

  const nameToEmail = await loadMasterNameMap(sheets, masterSheetId);

  // Fetch target spreadsheet metadata once
  const targetMeta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: targetSpreadsheetId })
  );
  const allTabs = targetMeta.data.sheets ?? [];

  const results: EmailFinderTabResult[] = [];

  for (const tabName of tabs) {
    try {
      const result = await processTab(
        sheets,
        targetSpreadsheetId,
        allTabs,
        tabName,
        nameToEmail
      );
      results.push(result);
    } catch (err) {
      results.push({
        tabName,
        totalRows: 0,
        matched: 0,
        unmatched: 0,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  return {
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${targetSpreadsheetId}/edit`,
    tabs: results,
  };
}

async function processTab(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  allTabs: sheets_v4.Schema$Sheet[],
  tabName: string,
  nameToEmail: Map<string, string>
): Promise<EmailFinderTabResult> {
  const tab = allTabs.find((t) => t.properties?.title === tabName);
  if (!tab?.properties) {
    throw new Error(`Tab "${tabName}" not found`);
  }
  const sheetId = tab.properties.sheetId;
  if (typeof sheetId !== "number") {
    throw new Error(`Tab "${tabName}" has no sheetId`);
  }

  const safeTab = `'${tabName.replace(/'/g, "''")}'`;

  // Read headers
  const headerRes = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeTab}!1:1`,
    })
  );
  const headers = (headerRes.data.values?.[0] ?? []) as string[];

  const nameCols = findNameColumns(headers);
  const nameColIndices = [
    nameCols.fullName,
    nameCols.firstName,
    nameCols.lastName,
  ].filter((v): v is number => typeof v === "number");

  if (nameColIndices.length === 0) {
    throw new Error(
      `No name column found in "${tabName}" (expected Name / Full Name / First Name + Last Name / Nome + Cognome)`
    );
  }

  const rightmostNameCol = Math.max(...nameColIndices);
  const emailHeaderIdx = rightmostNameCol + 1;
  const existingEmailAtInsert =
    headers[emailHeaderIdx]?.toString().trim().toLowerCase() ===
    EMAIL_HEADER.toLowerCase();

  // Read name column data
  const nameRanges: string[] = [];
  const rangeMap: { fullName?: number; firstName?: number; lastName?: number } = {};

  if (nameCols.fullName !== undefined) {
    rangeMap.fullName = nameRanges.length;
    const letter = columnIndexToLetter(nameCols.fullName);
    nameRanges.push(`${safeTab}!${letter}2:${letter}`);
  }
  if (nameCols.firstName !== undefined) {
    rangeMap.firstName = nameRanges.length;
    const letter = columnIndexToLetter(nameCols.firstName);
    nameRanges.push(`${safeTab}!${letter}2:${letter}`);
  }
  if (nameCols.lastName !== undefined) {
    rangeMap.lastName = nameRanges.length;
    const letter = columnIndexToLetter(nameCols.lastName);
    nameRanges.push(`${safeTab}!${letter}2:${letter}`);
  }

  const dataRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges: nameRanges,
    })
  );
  const valueRanges = dataRes.data.valueRanges ?? [];
  const fullRows =
    rangeMap.fullName !== undefined
      ? valueRanges[rangeMap.fullName]?.values ?? []
      : [];
  const firstRows =
    rangeMap.firstName !== undefined
      ? valueRanges[rangeMap.firstName]?.values ?? []
      : [];
  const lastRows =
    rangeMap.lastName !== undefined
      ? valueRanges[rangeMap.lastName]?.values ?? []
      : [];

  const rowCount = Math.max(fullRows.length, firstRows.length, lastRows.length);

  // Build per-row outcome: email (or null) and whether this row has a name
  interface RowOutcome {
    hasName: boolean;
    email: string | null;
  }
  const outcomes: RowOutcome[] = [];
  let matched = 0;
  let unmatched = 0;
  let nameRowCount = 0;

  for (let i = 0; i < rowCount; i++) {
    const full = fullRows[i]?.[0];
    const first = firstRows[i]?.[0];
    const last = lastRows[i]?.[0];

    let name = "";
    if (full && String(full).trim()) {
      name = String(full).trim();
    } else if ((first && String(first).trim()) || (last && String(last).trim())) {
      name = [first, last]
        .filter((v) => v && String(v).trim())
        .map((v) => String(v).trim())
        .join(" ");
    }

    if (!name) {
      outcomes.push({ hasName: false, email: null });
      continue;
    }

    nameRowCount++;
    const email = nameToEmail.get(normalizeName(name)) ?? null;
    if (email) matched++;
    else unmatched++;
    outcomes.push({ hasName: true, email });
  }

  // Build batchUpdate requests
  const requests: sheets_v4.Schema$Request[] = [];

  // 1. Insert new column (if not already there)
  if (!existingEmailAtInsert) {
    requests.push({
      insertDimension: {
        range: {
          sheetId,
          dimension: "COLUMNS",
          startIndex: emailHeaderIdx,
          endIndex: emailHeaderIdx + 1,
        },
        inheritFromBefore: false,
      },
    });
  }

  // 2. Write "Email" header (bold + subtle bg)
  requests.push({
    updateCells: {
      rows: [
        {
          values: [
            {
              userEnteredValue: { stringValue: EMAIL_HEADER },
              userEnteredFormat: {
                textFormat: { bold: true },
                backgroundColor: { red: 0.93, green: 0.93, blue: 0.96 },
              },
            },
          ],
        },
      ],
      fields:
        "userEnteredValue,userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor",
      start: { sheetId, rowIndex: 0, columnIndex: emailHeaderIdx },
    },
  });

  // 3. Write email values in one big updateCells (contiguous rows starting at row 2)
  if (outcomes.length > 0) {
    requests.push({
      updateCells: {
        rows: outcomes.map((o) => ({
          values: [
            {
              userEnteredValue: { stringValue: o.email ?? "" },
            },
          ],
        })),
        fields: "userEnteredValue",
        start: { sheetId, rowIndex: 1, columnIndex: emailHeaderIdx },
      },
    });
  }

  // 4. Set backgrounds on each name column (red for unmatched, white for matched, white for no-name)
  for (const colIdx of nameColIndices) {
    requests.push({
      updateCells: {
        rows: outcomes.map((o) => ({
          values: [
            {
              userEnteredFormat: {
                backgroundColor:
                  o.hasName && !o.email ? RED_BG : WHITE_BG,
              },
            },
          ],
        })),
        fields: "userEnteredFormat.backgroundColor",
        start: { sheetId, rowIndex: 1, columnIndex: colIdx },
      },
    });
  }

  // Execute in chunks
  const CHUNK_SIZE = 10;
  for (let i = 0; i < requests.length; i += CHUNK_SIZE) {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: requests.slice(i, i + CHUNK_SIZE) },
      })
    );
  }

  return {
    tabName,
    totalRows: nameRowCount,
    matched,
    unmatched,
  };
}
