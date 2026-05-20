import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { normalizeEmail, columnIndexToLetter } from "./email-utils";
import { extractSpreadsheetId } from "./url-parser";
import { parseFlexibleDate, toSheetsDateSerial } from "./sheet-dates";
import { sheets_v4 } from "googleapis";

/**
 * Consolidator
 *
 * A "section" merges rows from one or more source spreadsheets (each with
 * one or more picked tabs) into a user-chosen output spreadsheet + tab.
 *
 * Schema: the output is the UNION of every column header seen across all
 * picked tabs of all picked sources. Output column order = first-seen order.
 * Headers are matched case-insensitively with whitespace collapsed (so
 * "Email" and "email " fold into one column), and the first-seen casing is
 * kept for the output header text.
 *
 * Row model: **every source row is preserved** — no email-keyed merge. If
 * the same person appears in multiple sources with different data per
 * source, you get one output row per source so no data is lost.
 *
 * Highlighting passes applied after the bulk write:
 *
 *  1. Email duplicates — values seen >1 time in the Email column get the
 *     first occurrence painted light green and every subsequent occurrence
 *     painted light red. Same convention as the Duplicate Finder feature.
 *
 *  2. Phone duplicates — same rule for the phone column (case-insensitive
 *     header alias match: Telefono Cellulare, Telefono, Cellulare, Phone,
 *     Mobile, …). Comparison strips all non-digits so "+39 333 1234567"
 *     and "393331234567" dedupe to the same key.
 *
 *  3. Renewal Date conditional formatting — if a Renewal Date column is
 *     present (alias-matched: "Renewal Date", "Renewal / Expiry Date",
 *     "Expiry", "Data rinnovo", …), the cells in that column are converted
 *     to real Sheets date values (number serial + DATE format) and four
 *     native conditional-format rules are installed on the section's row
 *     range, with the same four-tier semantics as the renewal-sync engine:
 *
 *       past      → dark red bg + white text
 *       0–4 days  → light red bg
 *       5–14 days → light yellow bg
 *       15–30 days→ light green bg
 *
 *     Conditional rules persist in the spreadsheet — Sheets re-evaluates
 *     them on every open and every cell edit, so highlighting refreshes
 *     daily without re-running consolidate.
 *
 * Output destination: writes back into the user-chosen output spreadsheet
 * (the section's `outputUrl`), as a tab named `outputTabName`. Each run is
 * a full overwrite of that tab.
 */

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

// Duplicate-finder palette
const DUP_GREEN_BG = { red: 0.78, green: 0.92, blue: 0.78 };
const DUP_RED_BG = { red: 1.0, green: 0.82, blue: 0.82 };

// Renewal-tier palette (matches renewal-sync exactly)
const LIGHT_GREEN_BG = { red: 0.78, green: 0.92, blue: 0.78 }; // 15-30 days
const LIGHT_YELLOW_BG = { red: 1.0, green: 0.95, blue: 0.7 }; // 5-14 days
const LIGHT_RED_BG = { red: 1.0, green: 0.82, blue: 0.82 }; // 0-4 days
const DARK_RED_BG = { red: 0.78, green: 0.1, blue: 0.1 }; // past
const BLACK_TEXT = { red: 0, green: 0, blue: 0 };
const WHITE_TEXT = { red: 1, green: 1, blue: 1 };

// Column-detection aliases (all values are pre-normalized lowercase,
// trimmed, whitespace-collapsed — matched against the registry keys which
// are also normalized).
const EMAIL_ALIASES = [
  "email",
  "e-mail",
  "email address",
  "emailaddress",
  "mail",
  "indirizzo email",
  "indirizzo e-mail",
  "indirizzo mail",
];

const PHONE_ALIASES = [
  "telefono cellulare",
  "telefono cellular",
  "cellulare",
  "telefono",
  "telefono mobile",
  "numero di telefono",
  "numero telefono",
  "phone number",
  "mobile number",
  "phone",
  "mobile",
  "cell",
  "cellphone",
  "tel",
  "telephone",
];

const PHONE_SUBSTRING_FALLBACKS = [
  "telefono",
  "cellulare",
  "phone",
  "mobile",
];

const RENEWAL_DATE_ALIASES = [
  "renewal date",
  "renewal",
  "data rinnovo",
  "rinnovo",
  "data di rinnovo",
  "renewal / expiry date",
  "renewal/expiry date",
  "expiry date",
  "expiry",
];

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
  columnsContributed: number;
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
  totalOutputRows: number;
  rowsWithoutEmail: number;
  totalColumns: number;
  emailDuplicateValues: number; // distinct email values seen >1 time
  emailDuplicateCells: number; // cells painted red (occurrences - 1 per value)
  phoneDuplicateValues: number;
  phoneDuplicateCells: number;
  renewalRulesInstalled: boolean;
  sources: ConsolidatorSourceResult[];
  error?: string;
}

interface ColumnRegistry {
  keys: string[];
  headers: Map<string, string>;
}

interface ConsolidatedRow {
  values: Map<string, string>; // normalized column key → cell value
}

function normalizeColumnKey(header: string): string {
  return header.trim().toLowerCase().replace(/\s+/g, " ");
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    return value.toLocaleString("en-US", {
      useGrouping: false,
      maximumFractionDigits: 20,
    });
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim();
}

function normalizePhoneKey(value: string): string | null {
  const digits = value.replace(/\D/g, "");
  return digits || null;
}

/** Find the registry key matching any of the aliases (exact normalized). */
function findRegistryKey(
  registry: ColumnRegistry,
  aliases: string[],
  substringFallbacks: string[] = []
): string | null {
  for (const alias of aliases) {
    const want = normalizeColumnKey(alias);
    const hit = registry.keys.find((k) => k === want);
    if (hit) return hit;
  }
  for (const sub of substringFallbacks) {
    const wantSub = normalizeColumnKey(sub);
    const hit = registry.keys.find((k) => k.includes(wantSub));
    if (hit) return hit;
  }
  return null;
}

/**
 * Ensures the output tab exists with enough columns to hold the union.
 * Grows the grid via appendDimension if an existing tab is too narrow.
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
 * Reads picked tabs from one source spreadsheet, registers headers, and
 * appends every non-empty row to allRows. No merging.
 */
async function readSourceSpreadsheet(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabs: string[],
  registry: ColumnRegistry,
  allRows: ConsolidatedRow[],
  emailColumnExisted: { value: boolean },
  counters: { totalSourceRows: number; rowsWithoutEmail: number }
): Promise<ConsolidatorTabResult[]> {
  const ranges = tabs.map((tab) => {
    const safe = `'${tab.replace(/'/g, "''")}'`;
    return `${safe}!A:ZZ`;
  });

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

      // Map column index → normalized registry key (or "" if blank header).
      const tabColKeys: string[] = new Array(headerStrings.length);
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
        if (!registry.headers.has(key)) {
          registry.keys.push(key);
          registry.headers.set(key, orig);
        }
      }
      tabResult.columnsContributed =
        registry.keys.length - beforeRegistrySize;

      // Note whether this tab even has an email column — we use this later
      // to know if the email-dupe highlighting pass is worth running.
      const tabHasEmail = tabColKeys.some((k) =>
        EMAIL_ALIASES.includes(k)
      );
      if (tabHasEmail) emailColumnExisted.value = true;

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r] ?? [];
        const cellMap = new Map<string, string>();
        let nonEmpty = false;
        for (let c = 0; c < row.length; c++) {
          const key = tabColKeys[c];
          if (!key) continue;
          const val = clean(row[c]);
          if (!val) continue;
          if (!cellMap.has(key)) {
            cellMap.set(key, val);
          }
          nonEmpty = true;
        }

        if (!nonEmpty) continue;
        tabResult.totalRows++;
        counters.totalSourceRows++;

        // Track rows that have an email vs not (purely for reporting)
        let hasValidEmail = false;
        for (const k of EMAIL_ALIASES) {
          const v = cellMap.get(k);
          if (v && normalizeEmail(v)) {
            hasValidEmail = true;
            break;
          }
        }
        if (hasValidEmail) tabResult.rowsWithEmail++;
        else {
          tabResult.rowsWithoutEmail++;
          counters.rowsWithoutEmail++;
        }

        allRows.push({ values: cellMap });
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
    totalOutputRows: 0,
    rowsWithoutEmail: 0,
    totalColumns: 0,
    emailDuplicateValues: 0,
    emailDuplicateCells: 0,
    phoneDuplicateValues: 0,
    phoneDuplicateCells: 0,
    renewalRulesInstalled: false,
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
  const allRows: ConsolidatedRow[] = [];
  const emailColumnExisted = { value: false };
  const counters = { totalSourceRows: 0, rowsWithoutEmail: 0 };

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
          allRows,
          emailColumnExisted,
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
  sectionResult.rowsWithoutEmail = counters.rowsWithoutEmail;
  sectionResult.totalColumns = registry.keys.length;
  sectionResult.totalOutputRows = allRows.length;

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

    // Wipe the tab before writing the new union.
    await withRetry(() =>
      sheets.spreadsheets.values.clear({
        spreadsheetId: outputSpreadsheetId,
        range: `${safeTab}!A:ZZ`,
      })
    );

    if (headerRow.length === 0) {
      sectionResult.outputSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${outputSpreadsheetId}/edit#gid=${outputSheetId}`;
      return sectionResult;
    }

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

    // Build the per-feature follow-up requests (header styling, format
    // resets, duplicate highlights, renewal date conversion + conditional
    // rules). Sent as one batchUpdate at the end.
    const formatRequests: sheets_v4.Schema$Request[] = [];

    // (a) Header styling: bold + light bg + freeze first row.
    formatRequests.push({
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
    });
    formatRequests.push({
      updateSheetProperties: {
        properties: {
          sheetId: outputSheetId,
          gridProperties: { frozenRowCount: 1 },
        },
        fields: "gridProperties.frozenRowCount",
      },
    });

    // (b) Wipe any conditional-format rules left over on this tab from a
    //     previous run (or whatever the user had before). We always own
    //     this tab — full overwrite.
    const outputTabInfo = (
      await withRetry(() =>
        sheets.spreadsheets.get({ spreadsheetId: outputSpreadsheetId })
      )
    ).data.sheets?.find((t) => t.properties?.sheetId === outputSheetId);
    const existingRules = outputTabInfo?.conditionalFormats ?? [];
    for (let i = existingRules.length - 1; i >= 0; i--) {
      formatRequests.push({
        deleteConditionalFormatRule: { sheetId: outputSheetId, index: i },
      });
    }

    // (c) Reset background to white + text to black across the entire data
    //     range (rows 2..N, cols 0..rowEndCol). Clears any stale formatting
    //     from a previous run; conditional rules + duplicate highlights are
    //     applied on top.
    if (allRows.length > 0) {
      formatRequests.push({
        repeatCell: {
          range: {
            sheetId: outputSheetId,
            startRowIndex: 1,
            endRowIndex: allRows.length + 1,
            startColumnIndex: 0,
            endColumnIndex: headerRow.length,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 1, green: 1, blue: 1 },
              textFormat: { foregroundColor: BLACK_TEXT },
            },
          },
          fields:
            "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.foregroundColor",
        },
      });
    }

    // (d) Email duplicate highlighting (first green, subsequent red).
    const emailKey = findRegistryKey(registry, EMAIL_ALIASES);
    if (emailKey && allRows.length > 0) {
      const emailColIdx = registry.keys.indexOf(emailKey);
      const emailDupeMap = new Map<string, number[]>();
      for (let r = 0; r < allRows.length; r++) {
        const raw = allRows[r].values.get(emailKey);
        if (!raw) continue;
        const norm = normalizeEmail(raw);
        if (!norm) continue;
        const list = emailDupeMap.get(norm);
        if (list) list.push(r);
        else emailDupeMap.set(norm, [r]);
      }
      for (const [, rowIndices] of emailDupeMap) {
        if (rowIndices.length < 2) continue;
        sectionResult.emailDuplicateValues++;
        // First = green
        formatRequests.push(
          paintCellRequest(
            outputSheetId,
            rowIndices[0] + 1,
            emailColIdx,
            DUP_GREEN_BG
          )
        );
        for (let i = 1; i < rowIndices.length; i++) {
          formatRequests.push(
            paintCellRequest(
              outputSheetId,
              rowIndices[i] + 1,
              emailColIdx,
              DUP_RED_BG
            )
          );
          sectionResult.emailDuplicateCells++;
        }
      }
    }

    // (e) Phone duplicate highlighting (same rule, strip non-digits).
    const phoneKey = findRegistryKey(
      registry,
      PHONE_ALIASES,
      PHONE_SUBSTRING_FALLBACKS
    );
    if (phoneKey && allRows.length > 0) {
      const phoneColIdx = registry.keys.indexOf(phoneKey);
      const phoneDupeMap = new Map<string, number[]>();
      for (let r = 0; r < allRows.length; r++) {
        const raw = allRows[r].values.get(phoneKey);
        if (!raw) continue;
        const norm = normalizePhoneKey(raw);
        if (!norm) continue;
        const list = phoneDupeMap.get(norm);
        if (list) list.push(r);
        else phoneDupeMap.set(norm, [r]);
      }
      for (const [, rowIndices] of phoneDupeMap) {
        if (rowIndices.length < 2) continue;
        sectionResult.phoneDuplicateValues++;
        formatRequests.push(
          paintCellRequest(
            outputSheetId,
            rowIndices[0] + 1,
            phoneColIdx,
            DUP_GREEN_BG
          )
        );
        for (let i = 1; i < rowIndices.length; i++) {
          formatRequests.push(
            paintCellRequest(
              outputSheetId,
              rowIndices[i] + 1,
              phoneColIdx,
              DUP_RED_BG
            )
          );
          sectionResult.phoneDuplicateCells++;
        }
      }
    }

    // (f) Renewal Date column: convert parseable values to real Sheets
    //     dates (numberValue + DATE format), then install the four tier
    //     conditional-format rules.
    const renewalKey = findRegistryKey(registry, RENEWAL_DATE_ALIASES);
    if (renewalKey && allRows.length > 0) {
      const renewalColIdx = registry.keys.indexOf(renewalKey);

      // Build one updateCells per parseable date cell. Skip unparseable
      // cells (they stay as the text we already wrote).
      for (let r = 0; r < allRows.length; r++) {
        const raw = allRows[r].values.get(renewalKey) ?? "";
        if (!raw) continue;
        const parsed = parseFlexibleDate(raw);
        if (!parsed) continue;
        formatRequests.push({
          updateCells: {
            rows: [
              {
                values: [
                  {
                    userEnteredValue: {
                      numberValue: toSheetsDateSerial(parsed),
                    },
                    userEnteredFormat: {
                      numberFormat: {
                        type: "DATE",
                        pattern: "dd/mm/yyyy",
                      },
                    },
                  },
                ],
              },
            ],
            fields: "userEnteredValue,userEnteredFormat.numberFormat",
            start: {
              sheetId: outputSheetId,
              rowIndex: r + 1,
              columnIndex: renewalColIdx,
            },
          },
        });
      }

      // Install conditional-format rules over every row in the sheet's
      // grid (covers future rows too).
      const sheetRowCount =
        outputTabInfo?.properties?.gridProperties?.rowCount ??
        Math.max(allRows.length + 1, 1000);
      const renewalColLetter = columnIndexToLetter(renewalColIdx);
      const ref = `$${renewalColLetter}2`;
      const range: sheets_v4.Schema$GridRange = {
        sheetId: outputSheetId,
        startRowIndex: 1,
        endRowIndex: sheetRowCount,
        startColumnIndex: 0,
        endColumnIndex: headerRow.length,
      };

      const tierRules = [
        {
          formula: `=AND(ISNUMBER(${ref}), ${ref}<TODAY())`,
          bg: DARK_RED_BG,
          fg: WHITE_TEXT,
        },
        {
          formula: `=AND(ISNUMBER(${ref}), ${ref}>=TODAY(), ${ref}<=TODAY()+4)`,
          bg: LIGHT_RED_BG,
          fg: BLACK_TEXT,
        },
        {
          formula: `=AND(ISNUMBER(${ref}), ${ref}>=TODAY()+5, ${ref}<=TODAY()+14)`,
          bg: LIGHT_YELLOW_BG,
          fg: BLACK_TEXT,
        },
        {
          formula: `=AND(ISNUMBER(${ref}), ${ref}>=TODAY()+15, ${ref}<=TODAY()+30)`,
          bg: LIGHT_GREEN_BG,
          fg: BLACK_TEXT,
        },
      ];

      tierRules.forEach((r, idx) => {
        formatRequests.push({
          addConditionalFormatRule: {
            rule: {
              ranges: [range],
              booleanRule: {
                condition: {
                  type: "CUSTOM_FORMULA",
                  values: [{ userEnteredValue: r.formula }],
                },
                format: {
                  backgroundColor: r.bg,
                  textFormat: { foregroundColor: r.fg },
                },
              },
            },
            index: idx,
          },
        });
      });

      sectionResult.renewalRulesInstalled = true;
    }

    // Send all format work in chunked batches.
    const CHUNK_SIZE = 50;
    for (let i = 0; i < formatRequests.length; i += CHUNK_SIZE) {
      const slice = formatRequests.slice(i, i + CHUNK_SIZE);
      if (slice.length === 0) continue;
      try {
        await withRetry(() =>
          sheets.spreadsheets.batchUpdate({
            spreadsheetId: outputSpreadsheetId,
            requestBody: { requests: slice },
          })
        );
      } catch (err) {
        // Don't fail the whole section over a styling glitch — surface
        // it on the section result instead and keep going.
        const msg = err instanceof Error ? err.message : "format error";
        sectionResult.error = sectionResult.error
          ? `${sectionResult.error}; ${msg}`
          : `Formatting partially applied: ${msg}`;
        break;
      }
    }

    sectionResult.outputSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${outputSpreadsheetId}/edit#gid=${outputSheetId}`;
  } catch (err) {
    sectionResult.error =
      err instanceof Error ? err.message : "Failed to write output";
  }

  return sectionResult;
}

function paintCellRequest(
  sheetId: number,
  rowIndex: number,
  colIndex: number,
  bg: { red: number; green: number; blue: number }
): sheets_v4.Schema$Request {
  return {
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: colIndex,
        endColumnIndex: colIndex + 1,
      },
      cell: {
        userEnteredFormat: { backgroundColor: bg },
      },
      fields: "userEnteredFormat.backgroundColor",
    },
  };
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
        totalOutputRows: 0,
        rowsWithoutEmail: 0,
        totalColumns: 0,
        emailDuplicateValues: 0,
        emailDuplicateCells: 0,
        phoneDuplicateValues: 0,
        phoneDuplicateCells: 0,
        renewalRulesInstalled: false,
        sources: [],
        error: err instanceof Error ? err.message : "Section failed",
      });
    }
  }
  return results;
}
