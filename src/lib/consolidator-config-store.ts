import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";

/**
 * Per-user config for the Consolidator feature, persisted as a hidden
 * `_consolidator_config` tab inside the user's master spreadsheet.
 *
 * Stored shape: one row, columns [sourceUrl, sourceTabs(JSON)].
 */

const TAB = "_consolidator_config";
const HEADERS = ["sourceUrl", "sourceTabs"];

export interface ConsolidatorConfig {
  sourceUrl: string;
  sourceTabs: string[];
}

const EMPTY: ConsolidatorConfig = {
  sourceUrl: "",
  sourceTabs: [],
};

async function ensureTab(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );
  const exists = (meta.data.sheets ?? []).some(
    (t) => t.properties?.title === TAB
  );
  if (exists) return;

  await withRetry(() =>
    sheets.spreadsheets.batchUpdate({
      spreadsheetId: masterSheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: TAB } } }],
      },
    })
  );
  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A1:B1`,
      valueInputOption: "RAW",
      requestBody: { values: [HEADERS] },
    })
  );
}

function parseTabs(raw: unknown): string[] {
  if (typeof raw !== "string" || !raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.filter((t): t is string => typeof t === "string");
    }
  } catch {
    // not JSON
  }
  return [];
}

export async function getConsolidatorConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<ConsolidatorConfig> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:B2`,
    })
  );

  const row = res.data.values?.[0] ?? [];
  if (!row.length) return { ...EMPTY };

  return {
    sourceUrl: typeof row[0] === "string" ? row[0] : "",
    sourceTabs: parseTabs(row[1]),
  };
}

export async function saveConsolidatorConfig(
  refreshToken: string,
  masterSheetId: string,
  config: ConsolidatorConfig
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:B2`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[config.sourceUrl, JSON.stringify(config.sourceTabs)]],
      },
    })
  );
}

export async function clearConsolidatorConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:B2`,
    })
  );
}
