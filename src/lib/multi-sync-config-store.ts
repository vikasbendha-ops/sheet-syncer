// src/lib/multi-sync-config-store.ts
import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import type { MultiLinkedSheet, MultiSyncSection } from "./multi-sync-types";

/**
 * Two hidden tabs in the user's master spreadsheet:
 *
 *   `_multi_sync_meta`    — single integer (the next slot to assign)
 *   `_multi_sync_config`  — one row per section
 *
 * The slot counter is monotonic — `claimSlot` increments it and returns
 * the claimed value. Deleting a section does NOT decrement it, so a slot
 * number once assigned is permanent and never reused. This guarantees
 * the user's "Present In - N" column names never collide across the
 * lifetime of the master sheet.
 */

const CONFIG_TAB = "_multi_sync_config";
const META_TAB = "_multi_sync_meta";
const CONFIG_HEADERS = [
  "sectionId",
  "sectionName",
  "slot",
  "masterTabName",
  "presentInColumnName",
  "linkedSheets",
];
const META_HEADERS = ["nextSlot"];

export interface MultiSyncConfig {
  sections: MultiSyncSection[];
}

const EMPTY: MultiSyncConfig = { sections: [] };

function genId(): string {
  return `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

async function ensureMetaTab(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );
  const exists = (meta.data.sheets ?? []).some(
    (t) => t.properties?.title === META_TAB
  );
  if (exists) return;

  await withRetry(() =>
    sheets.spreadsheets.batchUpdate({
      spreadsheetId: masterSheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: META_TAB } } }],
      },
    })
  );
  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${META_TAB}!A1:A2`,
      valueInputOption: "RAW",
      requestBody: { values: [META_HEADERS, ["1"]] },
    })
  );
}

async function ensureConfigTab(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );
  const exists = (meta.data.sheets ?? []).some(
    (t) => t.properties?.title === CONFIG_TAB
  );
  if (exists) return;

  await withRetry(() =>
    sheets.spreadsheets.batchUpdate({
      spreadsheetId: masterSheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: CONFIG_TAB } } }],
      },
    })
  );
  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A1:F1`,
      valueInputOption: "RAW",
      requestBody: { values: [CONFIG_HEADERS] },
    })
  );
}

/**
 * Read the current nextSlot value (or 1 if the meta tab is missing /
 * unparseable), increment it, write back, and return the slot the
 * caller should use. Never returns the same number twice over the
 * lifetime of the master spreadsheet.
 */
export async function claimSlot(
  refreshToken: string,
  masterSheetId: string
): Promise<number> {
  await ensureMetaTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${META_TAB}!A2`,
    })
  );
  const raw = res.data.values?.[0]?.[0];
  let current = Number(raw);
  if (!Number.isFinite(current) || current < 1) current = 1;

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${META_TAB}!A2`,
      valueInputOption: "RAW",
      requestBody: { values: [[String(current + 1)]] },
    })
  );
  return current;
}

function parseLinkedSheets(raw: unknown): MultiLinkedSheet[] {
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
        lastSynced: typeof s.lastSynced === "string" ? s.lastSynced : "",
      }));
  } catch {
    return [];
  }
}

export async function getMultiSyncConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<MultiSyncConfig> {
  await ensureConfigTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:F1000`,
    })
  );
  const rows = res.data.values ?? [];
  if (!rows.length) return { ...EMPTY };

  const sections: MultiSyncSection[] = [];
  for (const row of rows) {
    const id = typeof row[0] === "string" && row[0] ? row[0] : genId();
    const name = typeof row[1] === "string" ? row[1] : "";
    const slotRaw = Number(row[2]);
    const slot = Number.isFinite(slotRaw) && slotRaw > 0 ? slotRaw : 1;
    const masterTabName = typeof row[3] === "string" ? row[3] : "";
    const presentInColumnName =
      typeof row[4] === "string" ? row[4] : "";
    const linkedSheets = parseLinkedSheets(row[5]);

    if (
      !name &&
      !masterTabName &&
      !presentInColumnName &&
      linkedSheets.length === 0
    ) {
      continue;
    }

    sections.push({
      id,
      name,
      slot,
      masterTabName,
      presentInColumnName,
      linkedSheets,
    });
  }
  return { sections };
}

/**
 * Validates that no two sections share the same `presentInColumnName`
 * (case-insensitive trim). Throws on collision so the API can surface a
 * 400. Empty strings are caught by the standard "required" checks; this
 * only guards against duplicates among non-empty values.
 */
function assertUniqueColumnNames(config: MultiSyncConfig): void {
  const seen = new Map<string, string>(); // normalized → sectionId
  for (const s of config.sections) {
    const key = (s.presentInColumnName ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
    if (!key) continue;
    const prev = seen.get(key);
    if (prev && prev !== s.id) {
      throw new Error(
        `Two sections share the Present In column name "${s.presentInColumnName}". Each section must have a unique column name.`
      );
    }
    seen.set(key, s.id);
  }
}

export async function saveMultiSyncConfig(
  refreshToken: string,
  masterSheetId: string,
  config: MultiSyncConfig
): Promise<void> {
  assertUniqueColumnNames(config);
  await ensureConfigTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A:F`,
    })
  );

  const rows: string[][] = [CONFIG_HEADERS];
  for (const s of config.sections) {
    rows.push([
      s.id || genId(),
      s.name ?? "",
      String(s.slot),
      s.masterTabName ?? "",
      s.presentInColumnName ?? "",
      JSON.stringify(s.linkedSheets ?? []),
    ]);
  }

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A1:F${rows.length}`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    })
  );
}

export async function clearMultiSyncConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  await ensureConfigTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);
  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:F1000`,
    })
  );
}

/**
 * Updates the lastSynced timestamp for one linked sheet inside one
 * section. Called by the engine after each successful read. Re-reads the
 * config to find the row, mutates the JSON, then writes back. Cheap for
 * a handful of sections; if this becomes hot we can switch to a single
 * batchUpdate per section run.
 */
export async function updateLinkedSheetLastSynced(
  refreshToken: string,
  masterSheetId: string,
  sectionId: string,
  nickname: string,
  url: string,
  timestamp: string
): Promise<void> {
  const config = await getMultiSyncConfig(refreshToken, masterSheetId);
  const sec = config.sections.find((s) => s.id === sectionId);
  if (!sec) return;
  const linked = sec.linkedSheets.find(
    (l) => l.nickname === nickname && l.url === url
  );
  if (!linked) return;
  linked.lastSynced = timestamp;
  await saveMultiSyncConfig(refreshToken, masterSheetId, config);
}
