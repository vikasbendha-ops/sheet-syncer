import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";

const TAB = "_biz_tutor_config";
const HEADERS = ["role", "url", "tabs"];

export interface BizTutorConfig {
  sourceUrl: string;
  sourceTabs: string[];
  lookupUrl: string;
  lookupTabs: string[];
}

const EMPTY: BizTutorConfig = {
  sourceUrl: "",
  sourceTabs: [],
  lookupUrl: "",
  lookupTabs: [],
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

export async function getBizTutorConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<BizTutorConfig> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:C3`,
    })
  );

  const rows = res.data.values ?? [];
  const config: BizTutorConfig = { ...EMPTY };
  for (const row of rows) {
    const role = row[0];
    const url = typeof row[1] === "string" ? row[1] : "";
    const tabs = parseTabs(row[2]);
    if (role === "source") {
      config.sourceUrl = url;
      config.sourceTabs = tabs;
    } else if (role === "lookup") {
      config.lookupUrl = url;
      config.lookupTabs = tabs;
    }
  }
  return config;
}

export async function saveBizTutorConfig(
  refreshToken: string,
  masterSheetId: string,
  config: BizTutorConfig
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const rows = [
    ["source", config.sourceUrl, JSON.stringify(config.sourceTabs)],
    ["lookup", config.lookupUrl, JSON.stringify(config.lookupTabs)],
  ];

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:C3`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    })
  );
}

export async function clearBizTutorConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:C3`,
    })
  );
}
