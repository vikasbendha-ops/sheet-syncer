// src/lib/multi-sync-engine.ts
import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import {
  normalizeEmail,
  findEmailColumnIndex,
  findNameColumns,
  columnIndexToLetter,
} from "./email-utils";
import { extractSpreadsheetId } from "./url-parser";
import { writePresentInColumn } from "./present-in-writer";
import { updateLinkedSheetLastSynced } from "./multi-sync-config-store";
import { sheets_v4 } from "googleapis";
import type {
  MultiLinkedSheet,
  MultiSyncBatchResult,
  MultiSyncSection,
  MultiSyncSectionResult,
  MultiTabResult,
} from "./multi-sync-types";

/**
 * Multi Sync engine — basic Present In sync, executed per section.
 *
 * For each section:
 *   1. Read every linked sheet's email + name columns (parallel, fan-out
 *      under MAX_CONCURRENCY).
 *   2. Build a cross-reference map (email → which sheets contain it).
 *   3. Write the section's UNIQUELY-NAMED Present In column into each
 *      source sheet (only touches the column whose header exactly equals
 *      section.presentInColumnName — never the unsuffixed "Present In"
 *      from basic sync, never other sections' columns).
 *   4. Write the section's Master tab into the user's master spreadsheet
 *      with header [Name, Email, <nick1>, ...] and ✅/❌ per linked sheet.
 *   5. Update lastSynced on each linked sheet's config row.
 *
 * Name prefix-folding logic is DUPLICATED from sync-engine.ts (per spec,
 * to keep basic sync 100% untouched).
 */

const MAX_CONCURRENCY = 3;
const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

async function withConcurrencyLimit<T>(
  tasks: (() => Promise<T>)[],
  limit: number
): Promise<T[]> {
  const results: T[] = [];
  const executing = new Set<Promise<void>>();
  for (const task of tasks) {
    const p = task().then((result) => {
      results.push(result);
      executing.delete(p);
    });
    executing.add(p);
    if (executing.size >= limit) {
      await Promise.race(executing);
    }
  }
  await Promise.all(executing);
  return results;
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

function letterToIndex(letter: string): number {
  if (!letter || letter.toLowerCase() === "auto") return -1;
  const upper = letter.toUpperCase();
  let idx = 0;
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    idx = idx * 26 + (code - 64);
  }
  return idx - 1;
}

/**
 * Prefix-folding name merge — duplicated from sync-engine.ts to keep
 * the basic sync untouched (per spec direction). If a bug is found,
 * fix both copies.
 */
function mergeNamesForEmail(
  namesByEmail: Map<string, string[]>
): Map<string, string> {
  const out = new Map<string, string>();
  for (const [email, names] of namesByEmail) {
    interface Candidate {
      name: string;
      words: string[];
      count: number;
      mergedCount: number;
      firstSeenIdx: number;
    }
    const candidates = new Map<string, Candidate>();
    let nextIdx = 0;
    for (const name of names) {
      const key = name.toLowerCase();
      const existing = candidates.get(key);
      if (existing) {
        existing.count++;
        existing.mergedCount++;
      } else {
        candidates.set(key, {
          name,
          words: key.split(/\s+/).filter(Boolean),
          count: 1,
          mergedCount: 1,
          firstSeenIdx: nextIdx++,
        });
      }
    }
    if (candidates.size === 0) continue;

    const list = Array.from(candidates.values());
    for (const a of list) {
      for (const b of list) {
        if (a === b) continue;
        if (b.words.length <= a.words.length) continue;
        const isPrefix = a.words.every((w, i) => b.words[i] === w);
        if (isPrefix) b.mergedCount += a.count;
      }
    }
    list.sort((a, b) => {
      if (b.mergedCount !== a.mergedCount) return b.mergedCount - a.mergedCount;
      if (b.words.length !== a.words.length) return b.words.length - a.words.length;
      return a.firstSeenIdx - b.firstSeenIdx;
    });
    out.set(email, list[0].name);
  }
  return out;
}

interface SheetReadRow {
  rowNumber: number;
  name: string;
}

interface SheetReadOutcome {
  result: MultiTabResult;
  spreadsheetId: string;
  sheetId: number;
  /** Map<normalizedEmail, SheetReadRow>. First-seen row wins per email. */
  byEmail: Map<string, SheetReadRow>;
  resolvedTabName: string;
}

async function readLinkedSheet(
  sheets: sheets_v4.Sheets,
  linked: MultiLinkedSheet
): Promise<SheetReadOutcome> {
  const tabResult: MultiTabResult = {
    nickname: linked.nickname,
    url: linked.url,
    tabName: linked.tabName,
    rowsRead: 0,
    emailsFound: 0,
  };
  const byEmail = new Map<string, SheetReadRow>();

  const spreadsheetId = extractSpreadsheetId(linked.url);

  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId })
  );
  const allTabs = meta.data.sheets ?? [];
  const matched = allTabs.find(
    (t) => t.properties?.title === linked.tabName
  );
  const sheetId = matched?.properties?.sheetId ?? -1;
  const resolvedTab = matched?.properties?.title ?? linked.tabName;
  if (sheetId < 0) {
    tabResult.error = `Tab "${linked.tabName}" not found`;
    return {
      result: tabResult,
      spreadsheetId,
      sheetId: -1,
      byEmail,
      resolvedTabName: resolvedTab,
    };
  }

  const safeTab = `'${resolvedTab.replace(/'/g, "''")}'`;

  const headerRes = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeTab}!1:1`,
    })
  );
  const headers = (headerRes.data.values?.[0] ?? []).map((h) => clean(h));

  let emailColIdx = letterToIndex(linked.emailColumn);
  if (emailColIdx < 0) emailColIdx = findEmailColumnIndex(headers);
  if (emailColIdx < 0) {
    tabResult.error = `No email column found in "${resolvedTab}". Headers: ${headers.filter(Boolean).join(", ")}`;
    return {
      result: tabResult,
      spreadsheetId,
      sheetId,
      byEmail,
      resolvedTabName: resolvedTab,
    };
  }

  const nameCols = findNameColumns(headers);

  const ranges: string[] = [];
  const rangeForCol = new Map<number, number>();
  function addColumnRange(colIdx: number) {
    if (colIdx < 0 || rangeForCol.has(colIdx)) return;
    rangeForCol.set(colIdx, ranges.length);
    const letter = columnIndexToLetter(colIdx);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }
  addColumnRange(emailColIdx);
  if (nameCols.fullName !== undefined) addColumnRange(nameCols.fullName);
  if (nameCols.firstName !== undefined) addColumnRange(nameCols.firstName);
  if (nameCols.lastName !== undefined) addColumnRange(nameCols.lastName);

  const batchRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges,
      valueRenderOption: "UNFORMATTED_VALUE",
    })
  );
  const valueRanges = batchRes.data.valueRanges ?? [];
  function valuesFor(colIdx: number): unknown[][] {
    if (colIdx < 0) return [];
    const ri = rangeForCol.get(colIdx);
    if (ri === undefined) return [];
    return (valueRanges[ri]?.values ?? []) as unknown[][];
  }

  const emailRows = valuesFor(emailColIdx);
  const fullNameRows =
    nameCols.fullName !== undefined ? valuesFor(nameCols.fullName) : [];
  const firstNameRows =
    nameCols.firstName !== undefined ? valuesFor(nameCols.firstName) : [];
  const lastNameRows =
    nameCols.lastName !== undefined ? valuesFor(nameCols.lastName) : [];

  for (let i = 0; i < emailRows.length; i++) {
    const rawEmail = emailRows[i]?.[0];
    if (rawEmail === null || rawEmail === undefined || rawEmail === "")
      continue;
    tabResult.rowsRead++;
    const normalized = normalizeEmail(String(rawEmail));
    if (!normalized) continue;
    tabResult.emailsFound++;

    const full = clean(fullNameRows[i]?.[0]);
    const first = clean(firstNameRows[i]?.[0]);
    const last = clean(lastNameRows[i]?.[0]);
    let name = "";
    if (full) {
      name = full;
    } else if (first || last) {
      if (
        first &&
        last &&
        first.toLowerCase() === last.toLowerCase()
      ) {
        name = first;
      } else {
        name = [first, last].filter(Boolean).join(" ");
      }
    }

    if (!byEmail.has(normalized)) {
      byEmail.set(normalized, { rowNumber: i + 2, name });
    }
  }

  return {
    result: tabResult,
    spreadsheetId,
    sheetId,
    byEmail,
    resolvedTabName: resolvedTab,
  };
}

export async function runMultiSyncSection(
  refreshToken: string,
  masterSheetId: string,
  section: MultiSyncSection
): Promise<MultiSyncSectionResult> {
  const timestamp = new Date().toISOString();
  const masterTabName =
    section.masterTabName || `Multi Master - ${section.slot}`;
  const presentInColumnName =
    section.presentInColumnName || `Present In - ${section.slot}`;

  const result: MultiSyncSectionResult = {
    sectionId: section.id,
    sectionName: section.name,
    slot: section.slot,
    masterSpreadsheetUrl: "",
    masterTabName,
    presentInColumnName,
    totalUniqueEmails: 0,
    linkedSheets: [],
    presentInWritten: false,
    timestamp,
  };

  if (!section.linkedSheets.length) {
    result.error = "Section needs at least one linked sheet.";
    return result;
  }

  const sheetsClient = getSheetsClient(refreshToken);

  // ---- Read phase ----
  const readTasks = section.linkedSheets.map(
    (linked) => async () => {
      try {
        return await readLinkedSheet(sheetsClient, linked);
      } catch (err) {
        const tabResult: MultiTabResult = {
          nickname: linked.nickname,
          url: linked.url,
          tabName: linked.tabName,
          rowsRead: 0,
          emailsFound: 0,
          error: err instanceof Error ? err.message : "Failed to read sheet",
        };
        return {
          result: tabResult,
          spreadsheetId: "",
          sheetId: -1,
          byEmail: new Map<string, SheetReadRow>(),
          resolvedTabName: linked.tabName,
        } as SheetReadOutcome;
      }
    }
  );
  const outcomes = await withConcurrencyLimit(readTasks, MAX_CONCURRENCY);
  outcomes.sort(
    (a, b) =>
      section.linkedSheets.findIndex((l) => l.nickname === a.result.nickname) -
      section.linkedSheets.findIndex((l) => l.nickname === b.result.nickname)
  );
  result.linkedSheets = outcomes.map((o) => o.result);

  // ---- Cross-reference ----
  const crossRef = new Map<
    string,
    Array<{
      linkedSheetIdx: number;
      sheetId: number;
      rowNumber: number;
      spreadsheetId: string;
    }>
  >();
  for (let i = 0; i < outcomes.length; i++) {
    const o = outcomes[i];
    for (const [email, row] of o.byEmail) {
      const list = crossRef.get(email) ?? [];
      list.push({
        linkedSheetIdx: i,
        sheetId: o.sheetId,
        rowNumber: row.rowNumber,
        spreadsheetId: o.spreadsheetId,
      });
      crossRef.set(email, list);
    }
  }
  result.totalUniqueEmails = crossRef.size;

  // ---- Present In write-back, per source sheet ----
  // Each section writes only its OWN column (matched by exact header
  // equality on section.presentInColumnName) — never the unsuffixed
  // "Present In" from basic sync, never another section's column.
  const presentInTasks: Array<() => Promise<void>> = [];
  for (let i = 0; i < outcomes.length; i++) {
    const outcome = outcomes[i];
    if (outcome.sheetId < 0) continue;
    const tabResult = result.linkedSheets[i];
    const safeTab = `'${outcome.resolvedTabName.replace(/'/g, "''")}'`;

    let presentInColumnIndex: number | null = null;
    let lastColumnIndex = -1;
    try {
      const headerRes = await withRetry(() =>
        sheetsClient.spreadsheets.values.get({
          spreadsheetId: outcome.spreadsheetId,
          range: `${safeTab}!1:1`,
        })
      );
      const headers = (headerRes.data.values?.[0] ?? []).map((h) =>
        clean(h)
      );
      const target = presentInColumnName.toLowerCase().trim();
      for (let c = 0; c < headers.length; c++) {
        if (headers[c]) lastColumnIndex = c;
        if (headers[c]?.toLowerCase().trim() === target) {
          presentInColumnIndex = c;
        }
      }
    } catch (err) {
      tabResult.error = `${tabResult.error ? tabResult.error + "; " : ""}Failed to read headers: ${err instanceof Error ? err.message : String(err)}`;
      continue;
    }

    const cellData: Array<{
      rowIndex: number;
      links: Array<{ text: string; url: string }>;
    }> = [];
    for (const [email, row] of outcome.byEmail) {
      const others = (crossRef.get(email) ?? []).filter(
        (x) => x.linkedSheetIdx !== i
      );
      if (others.length === 0) continue;
      cellData.push({
        rowIndex: row.rowNumber - 1,
        links: others.map((o) => ({
          text: section.linkedSheets[o.linkedSheetIdx]?.nickname ?? "?",
          url: `https://docs.google.com/spreadsheets/d/${o.spreadsheetId}/edit#gid=${o.sheetId}&range=A${o.rowNumber}`,
        })),
      });
    }

    const colIdx =
      presentInColumnIndex !== null
        ? presentInColumnIndex
        : lastColumnIndex + 1;
    const headerNeeded = presentInColumnIndex === null;

    presentInTasks.push(async () => {
      try {
        await writePresentInColumn(
          refreshToken,
          outcome.spreadsheetId,
          outcome.sheetId,
          colIdx,
          headerNeeded,
          cellData,
          presentInColumnName
        );
      } catch (err) {
        tabResult.error = `${tabResult.error ? tabResult.error + "; " : ""}Failed to write Present In: ${err instanceof Error ? err.message : String(err)}`;
      }
    });
  }
  await withConcurrencyLimit(presentInTasks, MAX_CONCURRENCY);
  result.presentInWritten = true;

  // ---- Master tab write ----
  try {
    const masterMeta = await withRetry(() =>
      sheetsClient.spreadsheets.get({ spreadsheetId: masterSheetId })
    );
    const allMasterTabs = masterMeta.data.sheets ?? [];
    const existing = allMasterTabs.find(
      (t) => t.properties?.title === masterTabName
    );
    let masterSheetIdNumber: number;
    if (
      existing?.properties?.sheetId !== undefined &&
      existing.properties.sheetId !== null
    ) {
      masterSheetIdNumber = existing.properties.sheetId;
    } else {
      const addRes = await withRetry(() =>
        sheetsClient.spreadsheets.batchUpdate({
          spreadsheetId: masterSheetId,
          requestBody: {
            requests: [
              {
                addSheet: { properties: { title: masterTabName } },
              },
            ],
          },
        })
      );
      masterSheetIdNumber =
        addRes.data.replies?.[0]?.addSheet?.properties?.sheetId ?? -1;
      if (masterSheetIdNumber < 0) {
        throw new Error(`Could not create master tab "${masterTabName}"`);
      }
    }

    const safeMaster = `'${masterTabName.replace(/'/g, "''")}'`;
    await withRetry(() =>
      sheetsClient.spreadsheets.values.clear({
        spreadsheetId: masterSheetId,
        range: `${safeMaster}!A:ZZ`,
      })
    );

    const namesByEmail = new Map<string, string[]>();
    for (const o of outcomes) {
      for (const [email, row] of o.byEmail) {
        if (!row.name) continue;
        const list = namesByEmail.get(email) ?? [];
        list.push(row.name);
        namesByEmail.set(email, list);
      }
    }
    const mergedNames = mergeNamesForEmail(namesByEmail);

    const sortedEmails = Array.from(crossRef.keys()).sort();
    const headerRow = [
      "Name",
      "Email",
      ...section.linkedSheets.map((l) => l.nickname || l.tabName),
    ];
    const dataRows = sortedEmails.map((email) => {
      const presence = section.linkedSheets.map((_, idx) =>
        outcomes[idx]?.byEmail.has(email) ? "✅" : "❌"
      );
      return [mergedNames.get(email) ?? "", email, ...presence];
    });
    const valueRows = [headerRow, ...dataRows];
    const lastColLetter = columnIndexToLetter(headerRow.length - 1);

    await withRetry(() =>
      sheetsClient.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${safeMaster}!A1:${lastColLetter}${valueRows.length}`,
        valueInputOption: "RAW",
        requestBody: { values: valueRows },
      })
    );

    // Best-effort header styling
    try {
      await withRetry(() =>
        sheetsClient.spreadsheets.batchUpdate({
          spreadsheetId: masterSheetId,
          requestBody: {
            requests: [
              {
                repeatCell: {
                  range: {
                    sheetId: masterSheetIdNumber,
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
                    sheetId: masterSheetIdNumber,
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
      // ignore styling failures
    }

    result.masterSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${masterSheetId}/edit#gid=${masterSheetIdNumber}`;
  } catch (err) {
    result.error = `${result.error ? result.error + "; " : ""}Failed to write Master tab: ${err instanceof Error ? err.message : String(err)}`;
  }

  // ---- Per-section, per-linked-sheet lastSynced update ----
  // Best-effort: failures here don't fail the section. Done sequentially
  // because updateLinkedSheetLastSynced re-reads + rewrites the whole
  // config tab (a handful of writes is fine).
  for (let i = 0; i < section.linkedSheets.length; i++) {
    const linked = section.linkedSheets[i];
    const tabResult = result.linkedSheets[i];
    if (tabResult?.error) continue; // only stamp on successful reads
    try {
      await updateLinkedSheetLastSynced(
        refreshToken,
        masterSheetId,
        section.id,
        linked.nickname,
        linked.url,
        timestamp
      );
    } catch {
      // ignore — non-fatal
    }
  }

  return result;
}

export async function runMultiSyncBatch(
  refreshToken: string,
  masterSheetId: string,
  sections: MultiSyncSection[]
): Promise<MultiSyncBatchResult> {
  const results: MultiSyncSectionResult[] = [];
  for (const section of sections) {
    try {
      results.push(
        await runMultiSyncSection(refreshToken, masterSheetId, section)
      );
    } catch (err) {
      results.push({
        sectionId: section.id,
        sectionName: section.name,
        slot: section.slot,
        masterSpreadsheetUrl: "",
        masterTabName:
          section.masterTabName || `Multi Master - ${section.slot}`,
        presentInColumnName:
          section.presentInColumnName || `Present In - ${section.slot}`,
        totalUniqueEmails: 0,
        linkedSheets: [],
        presentInWritten: false,
        timestamp: new Date().toISOString(),
        error: err instanceof Error ? err.message : "Section failed",
      });
    }
  }
  return { sections: results };
}
