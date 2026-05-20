import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";

/**
 * Per-user Consolidator config, persisted in a hidden `_consolidator_config`
 * tab inside the master spreadsheet.
 *
 * Schema (current):
 *   headers: [sectionId, name, outputUrl, outputTabName, sources]
 *   one row per section; `sources` cell holds a JSON array of {url, tabs[]}.
 *
 * Legacy schema (single-section, what was shipped first):
 *   headers: [sourceUrl, sourceTabs]
 *   one row holding a single source URL + tabs.
 * `ensureTab` detects the legacy headers on read and migrates them in place
 * into a one-section new-schema record so the UI sees a continuous history.
 */

const TAB = "_consolidator_config";
const HEADERS = [
  "sectionId",
  "name",
  "outputUrl",
  "outputTabName",
  "sources",
];
const LEGACY_HEADERS = ["sourceUrl", "sourceTabs"];

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

export interface ConsolidatorConfig {
  sections: ConsolidatorSection[];
}

const EMPTY: ConsolidatorConfig = { sections: [] };

function genId(): string {
  return `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function parseSources(raw: unknown): ConsolidatorSourceConfig[] {
  if (typeof raw !== "string" || !raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .filter(
        (s): s is { url?: unknown; tabs?: unknown } =>
          typeof s === "object" && s !== null
      )
      .map((s) => ({
        url: typeof s.url === "string" ? s.url : "",
        tabs: Array.isArray(s.tabs)
          ? (s.tabs as unknown[]).filter(
              (t): t is string => typeof t === "string"
            )
          : [],
      }));
  } catch {
    return [];
  }
}

function parseLegacyTabs(raw: unknown): string[] {
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

/**
 * Ensures the tab exists and that its headers are the current schema.
 * Returns the detected pre-existing headers so callers can migrate row data.
 */
async function ensureTab(
  refreshToken: string,
  masterSheetId: string
): Promise<{ existed: boolean; previousHeaders: string[] }> {
  const sheets = getSheetsClient(refreshToken);
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );
  const tabPresent = (meta.data.sheets ?? []).some(
    (t) => t.properties?.title === TAB
  );

  if (!tabPresent) {
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
        range: `${TAB}!A1:E1`,
        valueInputOption: "RAW",
        requestBody: { values: [HEADERS] },
      })
    );
    return { existed: false, previousHeaders: [] };
  }

  const headerRes = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A1:E1`,
    })
  );
  const previousHeaders = (headerRes.data.values?.[0] ?? []).map((h) =>
    String(h ?? "")
  );

  return { existed: true, previousHeaders };
}

export async function getConsolidatorConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<ConsolidatorConfig> {
  const { previousHeaders } = await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const isLegacy =
    previousHeaders[0] === LEGACY_HEADERS[0] &&
    previousHeaders[1] === LEGACY_HEADERS[1];

  if (isLegacy) {
    const res = await withRetry(() =>
      sheets.spreadsheets.values.get({
        spreadsheetId: masterSheetId,
        range: `${TAB}!A2:B2`,
      })
    );
    const row = res.data.values?.[0] ?? [];
    const sourceUrl = typeof row[0] === "string" ? row[0] : "";
    const tabs = parseLegacyTabs(row[1]);

    if (!sourceUrl && tabs.length === 0) {
      // Empty legacy row → just rewrite headers, no migrated section.
      await withRetry(() =>
        sheets.spreadsheets.values.clear({
          spreadsheetId: masterSheetId,
          range: `${TAB}!A:E`,
        })
      );
      await withRetry(() =>
        sheets.spreadsheets.values.update({
          spreadsheetId: masterSheetId,
          range: `${TAB}!A1:E1`,
          valueInputOption: "RAW",
          requestBody: { values: [HEADERS] },
        })
      );
      return { ...EMPTY };
    }

    const migrated: ConsolidatorSection = {
      id: genId(),
      name: "Section 1",
      sources: [{ url: sourceUrl, tabs }],
      outputUrl: sourceUrl, // legacy wrote into the same source spreadsheet
      outputTabName: "Consolidated",
    };

    // Rewrite the sheet with the new schema in place.
    await saveConsolidatorConfig(refreshToken, masterSheetId, {
      sections: [migrated],
    });
    return { sections: [migrated] };
  }

  // New schema read.
  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:E1000`,
    })
  );
  const rows = res.data.values ?? [];
  const sections: ConsolidatorSection[] = [];

  for (const row of rows) {
    const id = typeof row[0] === "string" && row[0] ? row[0] : genId();
    const name = typeof row[1] === "string" ? row[1] : "";
    const outputUrl = typeof row[2] === "string" ? row[2] : "";
    const outputTabName =
      typeof row[3] === "string" && row[3] ? row[3] : "Consolidated";
    const sources = parseSources(row[4]);

    if (!outputUrl && sources.length === 0 && !name) continue;

    sections.push({ id, name, sources, outputUrl, outputTabName });
  }

  return { sections };
}

export async function saveConsolidatorConfig(
  refreshToken: string,
  masterSheetId: string,
  config: ConsolidatorConfig
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  // Clear everything except header, then rewrite.
  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A:E`,
    })
  );

  const rows: string[][] = [HEADERS];
  for (const s of config.sections) {
    rows.push([
      s.id || genId(),
      s.name ?? "",
      s.outputUrl ?? "",
      s.outputTabName || "Consolidated",
      JSON.stringify(s.sources ?? []),
    ]);
  }

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A1:E${rows.length}`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
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
      range: `${TAB}!A2:E1000`,
    })
  );
}
