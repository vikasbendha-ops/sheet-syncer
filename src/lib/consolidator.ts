import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { normalizeEmail, columnIndexToLetter } from "./email-utils";
import { extractSpreadsheetId } from "./url-parser";
import { sheets_v4 } from "googleapis";

/**
 * Consolidator
 *
 * A "section" merges rows from one or more source spreadsheets (each with
 * one or more picked tabs) and writes a single output tab into a user-chosen
 * output spreadsheet.
 *
 * Multiple sections can be configured on the page and run sequentially.
 *
 * For each section:
 *  - Extract (Nome, Cognome, Email, Telefono Cellulare) from every picked
 *    tab across every picked source spreadsheet.
 *  - Dedupe across ALL of those rows by lowercased Email.
 *      * Rows with no valid email pass through unmerged.
 *      * When two rows share an email, "phone wins": if existing has no
 *        phone and newcomer does, replace phone. Otherwise keep existing.
 *        Blank Name / Surname fields are filled from the newcomer.
 *  - Write header [Name, Surname, Email, Phone] + deduped rows into the
 *    output spreadsheet's chosen tab (overwrite each run).
 */

// Hardcoded Italian source column names (case-insensitive header match)
const SRC_NAME = "Nome";
const SRC_SURNAME = "Cognome";
const SRC_EMAIL = "Email";
const SRC_PHONE = "Telefono Cellulare";

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

export interface ConsolidatorSourceConfig {
  url: string;
  tabs: string[];
}

export interface ConsolidatorSection {
  id: string;
  name: string;
  sources: ConsolidatorSourceConfig[];
  outputUrl: string;
  outputTabName: string;
}

export interface ConsolidatorTabResult {
  tabName: string;
  totalRows: number;
  rowsWithEmail: number;
  rowsWithoutEmail: number;
  missingColumns: string[];
  error?: string;
}

export interface ConsolidatorSourceResult {
  spreadsheetUrl: string;
  tabs: ConsolidatorTabResult[];
  error?: string;
}

export interface ConsolidatorSectionResult {
  sectionId: string;
  sectionName: string;
  outputSpreadsheetUrl: string;
  outputTabName: string;
  totalSourceRows: number;
  uniqueRows: number;
  duplicatesMerged: number;
  rowsWithoutEmail: number;
  sources: ConsolidatorSourceResult[];
  error?: string; // fatal: whole section bailed
}

interface ConsolidatedRow {
  name: string;
  surname: string;
  email: string; // normalized lowercase, or raw string when invalid
  phone: string;
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
  spreadsheetId: string,
  tabName: string
): Promise<number> {
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId })
  );
  const existing = (meta.data.sheets ?? []).find(
    (t) => t.properties?.title === tabName
  );
  if (
    existing?.properties?.sheetId !== undefined &&
    existing.properties.sheetId !== null
  ) {
    return existing.properties.sheetId;
  }

  const res = await withRetry(() =>
    sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: tabName } } }],
      },
    })
  );
  const newSheetId = res.data.replies?.[0]?.addSheet?.properties?.sheetId;
  if (newSheetId === undefined || newSheetId === null) {
    throw new Error(`Failed to create "${tabName}" tab`);
  }
  return newSheetId;
}

/**
 * Reads picked tabs from one source spreadsheet and accumulates into the
 * shared byEmail / noEmail collections. Returns per-tab read results plus a
 * counter delta for the caller to fold into totals.
 */
async function readSourceSpreadsheet(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabs: string[],
  byEmail: Map<string, ConsolidatedRow>,
  noEmail: ConsolidatedRow[],
  counters: { nextIdx: number; totalSourceRows: number; duplicatesMerged: number }
): Promise<ConsolidatorTabResult[]> {
  const ranges = tabs.map((tab) => {
    const safe = `'${tab.replace(/'/g, "''")}'`;
    return `${safe}!A:ZZ`;
  });

  const batchRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges,
    })
  );

  const valueRanges = batchRes.data.valueRanges ?? [];
  const results: ConsolidatorTabResult[] = [];

  for (let t = 0; t < tabs.length; t++) {
    const tabName = tabs[t];
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
        results.push(tabResult);
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

      if (
        idxName === -1 &&
        idxSurname === -1 &&
        idxEmail === -1 &&
        idxPhone === -1
      ) {
        tabResult.error = `None of the expected columns found. Headers: ${headers
          .filter(Boolean)
          .join(", ")}`;
        results.push(tabResult);
        continue;
      }

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r] ?? [];
        const name = idxName >= 0 ? clean(row[idxName]) : "";
        const surname = idxSurname >= 0 ? clean(row[idxSurname]) : "";
        const rawEmail = idxEmail >= 0 ? clean(row[idxEmail]) : "";
        const phone = idxPhone >= 0 ? clean(row[idxPhone]) : "";

        if (!name && !surname && !rawEmail && !phone) continue;

        tabResult.totalRows++;
        counters.totalSourceRows++;

        const normalizedEmail = rawEmail ? normalizeEmail(rawEmail) : null;

        if (!normalizedEmail) {
          tabResult.rowsWithoutEmail++;
          noEmail.push({
            name,
            surname,
            email: rawEmail,
            phone,
            firstSeenIdx: counters.nextIdx++,
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
            firstSeenIdx: counters.nextIdx++,
          });
          continue;
        }

        counters.duplicatesMerged++;
        if (!existing.phone && phone) existing.phone = phone;
        if (!existing.name && name) existing.name = name;
        if (!existing.surname && surname) existing.surname = surname;
      }

      results.push(tabResult);
    } catch (err) {
      tabResult.error =
        err instanceof Error ? err.message : "Failed to read tab";
      results.push(tabResult);
    }
  }

  return results;
}

export async function runConsolidatorSection(
  refreshToken: string,
  section: ConsolidatorSection
): Promise<ConsolidatorSectionResult> {
  const sectionResult: ConsolidatorSectionResult = {
    sectionId: section.id,
    sectionName: section.name,
    outputSpreadsheetUrl: "",
    outputTabName: section.outputTabName || "Consolidated",
    totalSourceRows: 0,
    uniqueRows: 0,
    duplicatesMerged: 0,
    rowsWithoutEmail: 0,
    sources: [],
  };

  const sheets = getSheetsClient(refreshToken);

  // Resolve output spreadsheet first (fail fast if URL is bogus)
  let outputSpreadsheetId: string;
  try {
    outputSpreadsheetId = extractSpreadsheetId(section.outputUrl);
  } catch (err) {
    sectionResult.error =
      err instanceof Error ? err.message : "Invalid output spreadsheet URL";
    return sectionResult;
  }

  if (section.sources.length === 0) {
    sectionResult.error = "Section has no source spreadsheets configured.";
    return sectionResult;
  }

  const byEmail = new Map<string, ConsolidatedRow>();
  const noEmail: ConsolidatedRow[] = [];
  const counters = { nextIdx: 0, totalSourceRows: 0, duplicatesMerged: 0 };

  for (const src of section.sources) {
    const sourceResult: ConsolidatorSourceResult = {
      spreadsheetUrl: src.url,
      tabs: [],
    };

    try {
      const spreadsheetId = extractSpreadsheetId(src.url);
      if (!src.tabs.length) {
        sourceResult.error = "No tabs selected for this source.";
      } else {
        sourceResult.tabs = await readSourceSpreadsheet(
          sheets,
          spreadsheetId,
          src.tabs,
          byEmail,
          noEmail,
          counters
        );
      }
    } catch (err) {
      sourceResult.error =
        err instanceof Error ? err.message : "Failed to read source";
    }

    sectionResult.sources.push(sourceResult);
  }

  sectionResult.totalSourceRows = counters.totalSourceRows;
  sectionResult.duplicatesMerged = counters.duplicatesMerged;
  sectionResult.rowsWithoutEmail = noEmail.length;

  const dedupedRows = Array.from(byEmail.values()).sort(
    (a, b) => a.firstSeenIdx - b.firstSeenIdx
  );
  const allRows = [...dedupedRows, ...noEmail];
  sectionResult.uniqueRows = allRows.length;

  // Write output
  try {
    const tabName = sectionResult.outputTabName;
    const outputSheetId = await ensureOutputTab(
      sheets,
      outputSpreadsheetId,
      tabName
    );

    const safeTab = `'${tabName.replace(/'/g, "''")}'`;
    await withRetry(() =>
      sheets.spreadsheets.values.clear({
        spreadsheetId: outputSpreadsheetId,
        range: `${safeTab}!A:D`,
      })
    );

    const headerRow = ["Name", "Surname", "Email", "Phone"];
    const dataRows = allRows.map((r) => [r.name, r.surname, r.email, r.phone]);
    const valueRows = [headerRow, ...dataRows];

    await withRetry(() =>
      sheets.spreadsheets.values.update({
        spreadsheetId: outputSpreadsheetId,
        range: `${safeTab}!A1:${columnIndexToLetter(headerRow.length - 1)}${valueRows.length}`,
        valueInputOption: "RAW",
        requestBody: { values: valueRows },
      })
    );

    // Best-effort header styling
    try {
      await withRetry(() =>
        sheets.spreadsheets.batchUpdate({
          spreadsheetId: outputSpreadsheetId,
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

    sectionResult.outputSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${outputSpreadsheetId}/edit#gid=${outputSheetId}`;
  } catch (err) {
    sectionResult.error =
      err instanceof Error ? err.message : "Failed to write output";
  }

  return sectionResult;
}

/**
 * Runs every section sequentially. A section's failure does not abort the
 * rest — its error is returned in that section's result instead.
 */
export async function runConsolidatorBatch(
  refreshToken: string,
  sections: ConsolidatorSection[]
): Promise<ConsolidatorSectionResult[]> {
  const results: ConsolidatorSectionResult[] = [];
  for (const section of sections) {
    try {
      const r = await runConsolidatorSection(refreshToken, section);
      results.push(r);
    } catch (err) {
      results.push({
        sectionId: section.id,
        sectionName: section.name,
        outputSpreadsheetUrl: "",
        outputTabName: section.outputTabName || "Consolidated",
        totalSourceRows: 0,
        uniqueRows: 0,
        duplicatesMerged: 0,
        rowsWithoutEmail: 0,
        sources: [],
        error: err instanceof Error ? err.message : "Section failed",
      });
    }
  }
  return results;
}
