import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { normalizeEmail, findEmailColumnIndex, columnIndexToLetter } from "./email-utils";
import { extractSpreadsheetId } from "./url-parser";
import { sheets_v4 } from "googleapis";

/**
 * Consolidator
 *
 * A "section" merges rows from one or more source spreadsheets (each with
 * one or more picked tabs) into a user-chosen output spreadsheet + tab.
 *
 * Schema: the output is the UNION of every column header seen across all
 * picked tabs of all picked sources. Output column order = first-seen order:
 * source 1 / tab 1's columns come first, then any NEW columns introduced
 * by subsequent tabs/sources appended at the end. Headers are matched
 * case-insensitively with whitespace collapsed (so "Email" and "email "
 * fold into one column), and the first-seen casing is kept for the output
 * header text.
 *
 * Dedupe + merge rule (per section, across all sources combined):
 *  - Dedupe key: lowercased, validated Email. Rows whose Email column is
 *    blank or invalid pass through unmerged at the bottom of the output.
 *  - When two rows share an email, columns are merged with "first non-blank
 *    value wins" semantics: the existing row keeps every non-blank cell it
 *    already has, and only its blank cells get filled in from the newcomer.
 *    This is the generalization of the old "phone wins" rule — any column
 *    where the first source had no value picks up the second source's value.
 *  - For Name / Surname / Phone / any column: same uniform rule.
 *
 * Output destination: writes back into the user-chosen output spreadsheet
 * (the section's `outputUrl`), as a tab named `outputTabName`. Each run is
 * a full overwrite of that tab (clear A:ZZ then write headers + rows).
 */

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
  totalRows: number; // non-empty rows read
  rowsWithEmail: number;
  rowsWithoutEmail: number;
  columnsContributed: number; // headers this tab introduced into the union
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
  totalColumns: number; // size of the consolidated column union
  sources: ConsolidatorSourceResult[];
  error?: string; // fatal: whole section bailed
}

/**
 * Tracks the running set of columns we've seen across a section's sources.
 *  - `keys` is the insertion-ordered list of normalized column keys.
 *  - `headers` maps a normalized key to the first-seen ORIGINAL header text
 *    (preserving the user's casing for the output row).
 */
interface ColumnRegistry {
  keys: string[];
  headers: Map<string, string>;
}

interface ConsolidatedRow {
  values: Map<string, string>; // normalized key → cell value
  firstSeenIdx: number; // stable ordering across the run
}

function normalizeColumnKey(header: string): string {
  return header.trim().toLowerCase().replace(/\s+/g, " ");
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    // Avoid scientific notation for large integers (e.g. phone numbers
    // stored numerically would render as 3.93E+11 under default toString).
    return value.toLocaleString("en-US", {
      useGrouping: false,
      maximumFractionDigits: 20,
    });
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim();
}

/**
 * Ensures the output tab exists with enough columns to hold the union.
 * - New tab: created with `requiredCols` columns straight away.
 * - Existing tab: grows the grid via appendDimension if its current
 *   columnCount is short.
 */
async function ensureOutputTab(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabName: string,
  requiredCols: number
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
    const sheetId = existing.properties.sheetId;
    const currentCols =
      existing.properties.gridProperties?.columnCount ?? 0;
    if (requiredCols > currentCols) {
      await withRetry(() =>
        sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [
              {
                appendDimension: {
                  sheetId,
                  dimension: "COLUMNS",
                  length: requiredCols - currentCols,
                },
              },
            ],
          },
        })
      );
    }
    return sheetId;
  }

  const res = await withRetry(() =>
    sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: tabName,
                gridProperties: {
                  columnCount: Math.max(requiredCols, 26),
                  rowCount: 1000,
                },
              },
            },
          },
        ],
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
 * Reads picked tabs from one source spreadsheet, registering every column
 * header seen, and folding rows into the section's byEmail / noEmail
 * collections with "first non-blank value wins" merge semantics.
 */
async function readSourceSpreadsheet(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabs: string[],
  registry: ColumnRegistry,
  byEmail: Map<string, ConsolidatedRow>,
  noEmail: ConsolidatedRow[],
  counters: {
    nextIdx: number;
    totalSourceRows: number;
    duplicatesMerged: number;
  }
): Promise<ConsolidatorTabResult[]> {
  const ranges = tabs.map((tab) => {
    const safe = `'${tab.replace(/'/g, "''")}'`;
    return `${safe}!A:ZZ`;
  });

  // UNFORMATTED_VALUE so numeric cells (phone numbers, IDs) don't come
  // back as "3.93E+11"; clean() handles the JS number → safe string
  // conversion afterward.
  const batchRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges,
      valueRenderOption: "UNFORMATTED_VALUE",
    })
  );

  const valueRanges = batchRes.data.valueRanges ?? [];
  const results: ConsolidatorTabResult[] = [];

  for (let t = 0; t < tabs.length; t++) {
    const tabName = tabs[t];
    const tabResult: ConsolidatorTabResult = {
      tabName,
      totalRows: 0,
      rowsWithEmail: 0,
      rowsWithoutEmail: 0,
      columnsContributed: 0,
    };

    try {
      const rows = (valueRanges[t]?.values ?? []) as unknown[][];
      if (rows.length === 0) {
        results.push(tabResult);
        continue;
      }

      const headerRow = rows[0] ?? [];
      const headerStrings = headerRow.map((h) => clean(h));

      // Build per-column key array for this tab. Empty/whitespace-only
      // headers are skipped — their cells are unreferenced. Duplicate
      // headers within the same tab take the FIRST column's index; later
      // duplicates are treated as the same key (their values fold into the
      // first column's value via blank-fill).
      const tabColKeys: string[] = new Array(headerStrings.length);
      const seenInThisTab = new Set<string>();
      const beforeRegistrySize = registry.keys.length;
      for (let c = 0; c < headerStrings.length; c++) {
        const orig = headerStrings[c];
        if (!orig) {
          tabColKeys[c] = "";
          continue;
        }
        const key = normalizeColumnKey(orig);
        if (!key) {
          tabColKeys[c] = "";
          continue;
        }
        tabColKeys[c] = key;
        seenInThisTab.add(key);
        if (!registry.headers.has(key)) {
          registry.keys.push(key);
          registry.headers.set(key, orig);
        }
      }
      tabResult.columnsContributed =
        registry.keys.length - beforeRegistrySize;

      // Find email column (alias-rich detection from email-utils)
      const emailColIdx = findEmailColumnIndex(
        headerStrings.map((h) => h)
      );

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r] ?? [];

        // Pull each cell, indexed by normalized column key. Empty cells
        // are not stored — Map.get returns undefined later, and the
        // merge rule treats undefined / "" as "blank".
        const cellMap = new Map<string, string>();
        let nonEmpty = false;
        for (let c = 0; c < row.length; c++) {
          const key = tabColKeys[c];
          if (!key) continue;
          const val = clean(row[c]);
          if (!val) continue;
          // First non-blank within the SAME row wins (in case of duplicate
          // headers in a tab — see comment above).
          if (!cellMap.has(key)) {
            cellMap.set(key, val);
          }
          nonEmpty = true;
        }

        if (!nonEmpty) continue;
        tabResult.totalRows++;
        counters.totalSourceRows++;

        let normEmail: string | null = null;
        if (emailColIdx >= 0) {
          const rawEmail = clean(row[emailColIdx]);
          normEmail = rawEmail ? normalizeEmail(rawEmail) : null;
        }

        if (!normEmail) {
          tabResult.rowsWithoutEmail++;
          noEmail.push({
            values: cellMap,
            firstSeenIdx: counters.nextIdx++,
          });
          continue;
        }

        tabResult.rowsWithEmail++;

        const existing = byEmail.get(normEmail);
        if (!existing) {
          byEmail.set(normEmail, {
            values: cellMap,
            firstSeenIdx: counters.nextIdx++,
          });
          continue;
        }

        // Merge: fill blanks on existing from newcomer. Existing's
        // already-set columns are kept (first-seen non-blank wins).
        counters.duplicatesMerged++;
        for (const [key, val] of cellMap) {
          const prev = existing.values.get(key);
          if (!prev) {
            existing.values.set(key, val);
          }
        }
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
    totalColumns: 0,
    sources: [],
  };

  const sheets = getSheetsClient(refreshToken);

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

  const registry: ColumnRegistry = { keys: [], headers: new Map() };
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
          registry,
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
  sectionResult.totalColumns = registry.keys.length;

  const dedupedRows = Array.from(byEmail.values()).sort(
    (a, b) => a.firstSeenIdx - b.firstSeenIdx
  );
  const allRows = [...dedupedRows, ...noEmail];
  sectionResult.uniqueRows = allRows.length;

  // If we read nothing, still create / clear the output tab so the user
  // sees a fresh empty Consolidated.
  const headerKeys = registry.keys;
  const headerRow = headerKeys.map((k) => registry.headers.get(k) ?? k);
  const requiredCols = Math.max(headerRow.length, 1);

  try {
    const tabName = sectionResult.outputTabName;
    const outputSheetId = await ensureOutputTab(
      sheets,
      outputSpreadsheetId,
      tabName,
      requiredCols
    );

    const safeTab = `'${tabName.replace(/'/g, "''")}'`;
    await withRetry(() =>
      sheets.spreadsheets.values.clear({
        spreadsheetId: outputSpreadsheetId,
        range: `${safeTab}!A:ZZ`,
      })
    );

    if (headerRow.length > 0) {
      const dataRows = allRows.map((r) =>
        headerKeys.map((k) => r.values.get(k) ?? "")
      );
      const valueRows = [headerRow, ...dataRows];
      const lastColLetter = columnIndexToLetter(headerRow.length - 1);

      await withRetry(() =>
        sheets.spreadsheets.values.update({
          spreadsheetId: outputSpreadsheetId,
          range: `${safeTab}!A1:${lastColLetter}${valueRows.length}`,
          valueInputOption: "RAW",
          requestBody: { values: valueRows },
        })
      );

      // Best-effort header styling: bold + light bg + freeze first row.
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
                    fields:
                      "userEnteredFormat(backgroundColor,textFormat)",
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
        // styling failure shouldn't fail the section
      }
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
        totalColumns: 0,
        sources: [],
        error: err instanceof Error ? err.message : "Section failed",
      });
    }
  }
  return results;
}
