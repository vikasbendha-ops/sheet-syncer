import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";

const TAB = "_report_sync_config";
const HEADERS = ["role", "url", "tabs"];

export interface ReportDestinationConfig {
  url: string;
  tabs: string[];
}

export interface ReportConfig {
  sourceUrl: string;
  sourceTabs: string[];
  destinations: ReportDestinationConfig[];
}

const EMPTY: ReportConfig = {
  sourceUrl: "",
  sourceTabs: [],
  destinations: [],
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
      range: `${TAB}!A1:C1`,
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
    // not JSON — fall through
  }
  return [];
}

export async function getReportConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<ReportConfig> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:C`,
    })
  );

  const rows = res.data.values ?? [];
  const config: ReportConfig = {
    sourceUrl: EMPTY.sourceUrl,
    sourceTabs: [...EMPTY.sourceTabs],
    destinations: [],
  };
  for (const row of rows) {
    const role = row[0];
    const url = typeof row[1] === "string" ? row[1] : "";
    const tabs = parseTabs(row[2]);
    if (role === "source") {
      config.sourceUrl = url;
      config.sourceTabs = tabs;
    } else if (role === "destination") {
      config.destinations.push({ url, tabs });
    }
  }
  return config;
}

export async function saveReportConfig(
  refreshToken: string,
  masterSheetId: string,
  config: ReportConfig
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const rows: string[][] = [];
  rows.push([
    "source",
    config.sourceUrl,
    JSON.stringify(config.sourceTabs),
  ]);
  for (const dest of config.destinations) {
    rows.push(["destination", dest.url, JSON.stringify(dest.tabs)]);
  }

  // Clear any prior rows (variable length list of destinations)
  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:C`,
    })
  );

  if (rows.length > 0) {
    await withRetry(() =>
      sheets.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${TAB}!A2:C${rows.length + 1}`,
        valueInputOption: "RAW",
        requestBody: { values: rows },
      })
    );
  }
}

export async function clearReportConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:C`,
    })
  );
}
