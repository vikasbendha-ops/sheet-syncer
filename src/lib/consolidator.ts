import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { normalizeEmail, columnIndexToLetter } from "./email-utils";
import { sheets_v4 } from "googleapis";

/**
 * Consolidator
 *
 * Reads selected tabs from a single source spreadsheet, extracts
 * (Nome, Cognome, Email, Telefono Cellulare) columns from each row,
 * dedupes across all rows by Email (case-insensitive), and writes a single
 * "Consolidated" tab into the master sheet.
 *
 * Dedupe rule:
 *  - Key: lowercased Email. Rows without email pass through unmerged
 *    (we cannot match them safely).
 *  - When two rows share an email, phone wins: if one row has a phone and
 *    the other doesn't, keep the one with a phone. Otherwise keep the
 *    existing row and only fill blanks from the newcomer.
 *  - For Name / Surname: prefer non-blank, first-seen as tiebreak.
 */

const OUTPUT_TAB = "Consolidated";

// Hardcoded Italian source column names (case-insensitive header match)
const SRC_NAME = "Nome";
const SRC_SURNAME = "Cognome";
const SRC_EMAIL = "Email";
const SRC_PHONE = "Telefono Cellulare";

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

export interface ConsolidatorTabResult {
  tabName: string;
  totalRows: number;
  rowsWithEmail: number;
  rowsWithoutEmail: number;
  missingColumns: string[];
  error?: string;
}

export interface ConsolidatorResult {
  spreadsheetUrl: string; // link to the master sheet's Consolidated tab
  totalSourceRows: number;
  uniqueRows: number;
  duplicatesMerged: number;
  rowsWithoutEmail: number;
  tabs: ConsolidatorTabResult[];
}

interface ConsolidatedRow {
  name: string;
  surname: string;
  email: string; // normalized lowercase (or "" for unmatched rows)
  phone: string;
  // bookkeeping
  firstSeenIdx: number;
}

function normalizeHeader(s: unknown): string {
  return String(s ?? "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ");
}

function findHeaderIndex(headers: string[], target: string): number {
  const want = normalizeHeader(target);
  return headers.findIndex((h) => normalizeHeader(h) === want);
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

async function ensureOutputTab(
  sheets: sheets_v4.Sheets,
  masterSheetId: string
): Promise<number> {
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );
  const existing = (meta.data.sheets ?? []).find(
    (t) => t.properties?.title === OUTPUT_TAB
  );
  if (existing?.properties?.sheetId !== undefined && existing.properties.sheetId !== null) {
    return existing.properties.sheetId;
  }

  const res = await withRetry(() =>
    sheets.spreadsheets.batchUpdate({
      spreadsheetId: masterSheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: OUTPUT_TAB } } }],
      },
    })
  );
  const newSheetId = res.data.replies?.[0]?.addSheet?.properties?.sheetId;
  if (newSheetId === undefined || newSheetId === null) {
    throw new Error("Failed to create Consolidated tab in master sheet");
  }
  return newSheetId;
}

export async function runConsolidator(
  refreshToken: string,
  masterSheetId: string,
  sourceSpreadsheetId: string,
  sourceTabs: string[]
): Promise<ConsolidatorResult> {
  const sheets = getSheetsClient(refreshToken);

  // Read all selected tabs from source spreadsheet
  const tabResults: ConsolidatorTabResult[] = [];

  // Batched read: one batchGet for all tabs' full ranges (header + data)
  const ranges = sourceTabs.map((tab) => {
    const safeTab = `'${tab.replace(/'/g, "''")}'`;
    return `${safeTab}!A:ZZ`;
  });

  const batchRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId: sourceSpreadsheetId,
      ranges,
    })
  );

  const valueRanges = batchRes.data.valueRanges ?? [];

  // Build a global, ordered list of rows; key by email when present.
  const byEmail = new Map<string, ConsolidatedRow>();
  const noEmail: ConsolidatedRow[] = [];
  let nextIdx = 0;
  let totalSourceRows = 0;
  let duplicatesMerged = 0;

  for (let t = 0; t < sourceTabs.length; t++) {
    const tabName = sourceTabs[t];
    const valueRange = valueRanges[t];
    const tabResult: ConsolidatorTabResult = {
      tabName,
      totalRows: 0,
      rowsWithEmail: 0,
      rowsWithoutEmail: 0,
      missingColumns: [],
    };

    try {
      const rows = (valueRange?.values ?? []) as string[][];
      if (rows.length === 0) {
        tabResults.push(tabResult);
        continue;
      }

      const headers = (rows[0] ?? []).map((h) => String(h ?? ""));

      const idxName = findHeaderIndex(headers, SRC_NAME);
      const idxSurname = findHeaderIndex(headers, SRC_SURNAME);
      const idxEmail = findHeaderIndex(headers, SRC_EMAIL);
      const idxPhone = findHeaderIndex(headers, SRC_PHONE);

      const missing: string[] = [];
      if (idxName === -1) missing.push(SRC_NAME);
      if (idxSurname === -1) missing.push(SRC_SURNAME);
      if (idxEmail === -1) missing.push(SRC_EMAIL);
      if (idxPhone === -1) missing.push(SRC_PHONE);
      tabResult.missingColumns = missing;

      // If literally none of the four columns are present, skip whole tab.
      // Otherwise pull whatever we can, leaving missing fields blank.
      if (
        idxName === -1 &&
        idxSurname === -1 &&
        idxEmail === -1 &&
        idxPhone === -1
      ) {
        tabResult.error = `None of the expected columns found. Headers: ${headers
          .filter(Boolean)
          .join(", ")}`;
        tabResults.push(tabResult);
        continue;
      }

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r] ?? [];
        const name = idxName >= 0 ? clean(row[idxName]) : "";
        const surname = idxSurname >= 0 ? clean(row[idxSurname]) : "";
        const rawEmail = idxEmail >= 0 ? clean(row[idxEmail]) : "";
        const phone = idxPhone >= 0 ? clean(row[idxPhone]) : "";

        // Skip wholly empty rows
        if (!name && !surname && !rawEmail && !phone) continue;

        tabResult.totalRows++;
        totalSourceRows++;

        const normalizedEmail = rawEmail ? normalizeEmail(rawEmail) : null;

        if (!normalizedEmail) {
          tabResult.rowsWithoutEmail++;
          noEmail.push({
            name,
            surname,
            email: rawEmail, // keep raw (may be invalid format) so user sees it
            phone,
            firstSeenIdx: nextIdx++,
          });
          continue;
        }

        tabResult.rowsWithEmail++;

        const existing = byEmail.get(normalizedEmail);
        if (!existing) {
          byEmail.set(normalizedEmail, {
            name,
            surname,
            email: normalizedEmail,
            phone,
            firstSeenIdx: nextIdx++,
          });
          continue;
        }

        // Merge: phone wins. Within name/surname: prefer non-blank, keep existing as tiebreak.
        duplicatesMerged++;

        // Phone: if existing lacks phone and new has phone, take new's phone.
        // If both have phones, keep existing (first-seen).
        if (!existing.phone && phone) {
          existing.phone = phone;
        }
        // Name / surname: fill blanks from newcomer
        if (!existing.name && name) existing.name = name;
        if (!existing.surname && surname) existing.surname = surname;
      }

      tabResults.push(tabResult);
    } catch (err) {
      tabResult.error =
        err instanceof Error ? err.message : "Failed to read tab";
      tabResults.push(tabResult);
    }
  }

  // Final row order: deduped-by-email first (in first-seen order), then no-email rows
  const dedupedRows = Array.from(byEmail.values()).sort(
    (a, b) => a.firstSeenIdx - b.firstSeenIdx
  );
  const allRows = [...dedupedRows, ...noEmail];

  // Write to master sheet's Consolidated tab
  const outputSheetId = await ensureOutputTab(sheets, masterSheetId);

  // Clear existing content first (full overwrite each run)
  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${OUTPUT_TAB}!A:D`,
    })
  );

  const headerRow = ["Name", "Surname", "Email", "Phone"];
  const dataRows = allRows.map((r) => [r.name, r.surname, r.email, r.phone]);
  const valueRows = [headerRow, ...dataRows];

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${OUTPUT_TAB}!A1:${columnIndexToLetter(headerRow.length - 1)}${valueRows.length}`,
      valueInputOption: "RAW",
      requestBody: { values: valueRows },
    })
  );

  // Style header row (bold + light bg) — best effort, do not fail the run if styling fails
  try {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId: masterSheetId,
        requestBody: {
          requests: [
            {
              repeatCell: {
                range: {
                  sheetId: outputSheetId,
                  startRowIndex: 0,
                  endRowIndex: 1,
                  startColumnIndex: 0,
                  endColumnIndex: headerRow.length,
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: HEADER_BG,
                    textFormat: { bold: true },
                  },
                },
                fields: "userEnteredFormat(backgroundColor,textFormat)",
              },
            },
            {
              updateSheetProperties: {
                properties: {
                  sheetId: outputSheetId,
                  gridProperties: { frozenRowCount: 1 },
                },
                fields: "gridProperties.frozenRowCount",
              },
            },
          ],
        },
      })
    );
  } catch {
    // ignore styling failure
  }

  return {
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${masterSheetId}/edit#gid=${outputSheetId}`,
    totalSourceRows,
    uniqueRows: allRows.length,
    duplicatesMerged,
    rowsWithoutEmail: noEmail.length,
    tabs: tabResults,
  };
}
