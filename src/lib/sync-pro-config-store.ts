// src/lib/sync-pro-config-store.ts
import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import type {
  ProLinkedSheet,
  ProPropagateColumn,
  ProSection,
} from "./sync-pro-types";

/**
 * Hidden config tab in the user's master spreadsheet. One row per Pro
 * section. The two array columns (`propagateColumns`, `linkedSheets`)
 * are stored as JSON strings to keep the schema flat — same pattern
 * the consolidator uses for its `sources` column.
 */

const TAB = "_sync_pro_config";
const HEADERS = [
  "sectionId",
  "sectionName",
  "masterTabName",
  "writePresentIn",
  "propagateColumns",
  "linkedSheets",
];

export interface SyncProConfig {
  sections: ProSection[];
}

const EMPTY: SyncProConfig = { sections: [] };

function genId(): string {
  return `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

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
      range: `${TAB}!A1:F1`,
      valueInputOption: "RAW",
      requestBody: { values: [HEADERS] },
    })
  );
}

function parsePropagateColumns(raw: unknown): ProPropagateColumn[] {
  if (typeof raw !== "string" || !raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .filter(
        (c): c is { name?: unknown } =>
          typeof c === "object" && c !== null
      )
      .map((c) => ({ name: typeof c.name === "string" ? c.name : "" }))
      .filter((c) => c.name);
  } catch {
    return [];
  }
}

function parseLinkedSheets(raw: unknown): ProLinkedSheet[] {
  if (typeof raw !== "string" || !raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .filter(
        (s): s is Record<string, unknown> =>
          typeof s === "object" && s !== null
      )
      .map((s) => ({
        url: typeof s.url === "string" ? s.url : "",
        nickname: typeof s.nickname === "string" ? s.nickname : "",
        tabName: typeof s.tabName === "string" ? s.tabName : "",
        emailColumn:
          typeof s.emailColumn === "string" ? s.emailColumn : "auto",
        columnMapping:
          typeof s.columnMapping === "object" && s.columnMapping !== null
            ? Object.fromEntries(
                Object.entries(s.columnMapping as Record<string, unknown>).map(
                  ([k, v]) => [
                    k,
                    typeof v === "string" ? v : v === null ? null : null,
                  ]
                )
              )
            : {},
      }));
  } catch {
    return [];
  }
}

export async function getSyncProConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<SyncProConfig> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:F1000`,
    })
  );
  const rows = res.data.values ?? [];
  if (!rows.length) return { ...EMPTY };

  const sections: ProSection[] = [];
  for (const row of rows) {
    const id = typeof row[0] === "string" && row[0] ? row[0] : genId();
    const name = typeof row[1] === "string" ? row[1] : "";
    const masterTabName = typeof row[2] === "string" ? row[2] : "";
    const writePresentIn =
      typeof row[3] === "string"
        ? row[3].toLowerCase() !== "false"
        : true;
    const propagateColumns = parsePropagateColumns(row[4]);
    const linkedSheets = parseLinkedSheets(row[5]);

    if (
      !name &&
      !masterTabName &&
      linkedSheets.length === 0 &&
      propagateColumns.length === 0
    ) {
      continue;
    }

    sections.push({
      id,
      name,
      masterTabName,
      writePresentIn,
      propagateColumns,
      linkedSheets,
    });
  }
  return { sections };
}

export async function saveSyncProConfig(
  refreshToken: string,
  masterSheetId: string,
  config: SyncProConfig
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A:F`,
    })
  );

  const rows: string[][] = [HEADERS];
  for (const s of config.sections) {
    rows.push([
      s.id || genId(),
      s.name ?? "",
      s.masterTabName ?? "",
      s.writePresentIn ? "true" : "false",
      JSON.stringify(s.propagateColumns ?? []),
      JSON.stringify(s.linkedSheets ?? []),
    ]);
  }

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A1:F${rows.length}`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    })
  );
}

export async function clearSyncProConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:F1000`,
    })
  );
}
