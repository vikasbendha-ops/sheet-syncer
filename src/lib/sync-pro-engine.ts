// src/lib/sync-pro-engine.ts
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
import { sheets_v4 } from "googleapis";
import type {
  ProBatchResult,
  ProColumnStats,
  ProConflict,
  ProLinkedSheet,
  ProSection,
  ProSectionResult,
  ProTabResult,
} from "./sync-pro-types";

/**
 * Sheet Syncer Pro engine.
 *
 * Each section is independent. Within a section:
 *   1. Read every linked sheet's email + name + mapped propagate columns.
 *   2. Cross-reference emails to build {email: which sheets contain it}.
 *   3. Propagation pass: for each logical propagate column, fill blanks
 *      from non-blank values. Never overwrite. Surface conflicts when
 *      multiple sheets have different non-blank values.
 *   4. Apply per-sheet writes (one batchUpdate per sheet).
 *   5. Write Present In column into each source sheet (optional).
 *   6. Write Master ✅/❌ tab into the master sheet.
 */

// Lowered from 5 to stay under Google Sheets' 60 read/min per-user quota,
// matching the main sync engine.
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

/**
 * Resolve a 1-based column letter (or "auto") to a 0-based column index.
 * `auto` returns -1 and the caller should fall back to header detection.
 */
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
 * Merge names across sheets for the Master tab.
 *
 * DUPLICATED FROM src/lib/sync-engine.ts so the basic sync stays 100%
 * untouched (per spec / user direction). If a bug is found in either
 * copy, fix both.
 *
 * Strategy:
 *  1. Count every name occurrence (case-insensitive).
 *  2. If name A is a word-prefix of name B (e.g. "Laura" ⊂ "Laura
 *     Pegoraro"), fold A's count into B so the more complete name wins.
 *  3. Pick winner by mergedCount → words.length → firstSeenIdx.
 */
function mergeNamesForEmail(namesByEmail: Map<string, string[]>): Map<string, string> {
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

// ===========================================================================
// Read phase
// ===========================================================================

/** A single read row from one linked sheet, keyed by normalized email. */
interface SheetRowRecord {
  /** 1-based row number in the source tab. */
  rowNumber: number;
  /** Display name extracted from the sheet's name column(s), if any. */
  name: string;
  /**
   * Map from logical column name → raw cell value (cleaned). Only contains
   * columns this linked sheet mapped to a real header.
   */
  perColumn: Map<string, string>;
}

interface SheetReadOutcome {
  result: ProTabResult;
  /** Resolved Google Sheets numeric sheetId for the picked tab. */
  sheetId: number;
  /** Resolved spreadsheetId. */
  spreadsheetId: string;
  /** Map<normalizedEmail, SheetRowRecord>. First-seen row wins per email. */
  byEmail: Map<string, SheetRowRecord>;
  /**
   * The actual column index in this sheet for each logical column the
   * section requested. -1 = not mapped / not found in this sheet.
   */
  logicalColIndex: Map<string, number>;
  /** 0-based index of the email column in this sheet. */
  emailColIdx: number;
}

/**
 * Reads one linked sheet. Resolves the email column (auto-detect or letter),
 * resolves each mapped logical column to a 0-based index, then bulk-reads
 * the values via batchGet.
 */
async function readLinkedSheet(
  sheets: sheets_v4.Sheets,
  linked: ProLinkedSheet,
  propagateColumns: string[]
): Promise<SheetReadOutcome> {
  const tabResult: ProTabResult = {
    nickname: linked.nickname,
    url: linked.url,
    tabName: linked.tabName,
    rowsRead: 0,
    emailsFound: 0,
    cellsFilled: 0,
  };
  const byEmail = new Map<string, SheetRowRecord>();
  const logicalColIndex = new Map<string, number>();

  const spreadsheetId = extractSpreadsheetId(linked.url);

  // Resolve tab + sheetId
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
      sheetId: -1,
      spreadsheetId,
      byEmail,
      logicalColIndex,
      emailColIdx: -1,
    };
  }

  const safeTab = `'${resolvedTab.replace(/'/g, "''")}'`;

  // Read header row to resolve column indices
  const headerRes = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeTab}!1:1`,
    })
  );
  const headers = (headerRes.data.values?.[0] ?? []).map((h) =>
    clean(h)
  );

  // Email column
  let emailColIdx = letterToIndex(linked.emailColumn);
  if (emailColIdx < 0) {
    emailColIdx = findEmailColumnIndex(headers);
  }
  if (emailColIdx < 0) {
    tabResult.error = `No email column found in "${resolvedTab}". Headers: ${headers.filter(Boolean).join(", ")}`;
    return {
      result: tabResult,
      sheetId,
      spreadsheetId,
      byEmail,
      logicalColIndex,
      emailColIdx: -1,
    };
  }

  // Name columns (for the Master tab's Name field)
  const nameCols = findNameColumns(headers);

  // Logical propagate columns → 0-based column indices (case-insensitive)
  const normalizedHeaders = headers.map((h) =>
    h.toLowerCase().trim().replace(/\s+/g, " ")
  );
  for (const logical of propagateColumns) {
    const mapped = linked.columnMapping[logical];
    if (!mapped) {
      logicalColIndex.set(logical, -1);
      continue;
    }
    const want = mapped.toLowerCase().trim().replace(/\s+/g, " ");
    const idx = normalizedHeaders.findIndex((h) => h === want);
    logicalColIndex.set(logical, idx);
  }

  // Build batch ranges: email + name columns + every mapped (idx >= 0) logical column
  const ranges: string[] = [];
  const rangeForCol = new Map<number, number>(); // col idx → ranges[] index
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
  for (const [, idx] of logicalColIndex) addColumnRange(idx);

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

    // Build display name (same logic as sync-engine reader)
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

    // Collect mapped column values
    const perColumn = new Map<string, string>();
    for (const [logical, idx] of logicalColIndex) {
      if (idx < 0) continue;
      const colRows = valuesFor(idx);
      const v = clean(colRows[i]?.[0]);
      if (v) perColumn.set(logical, v);
    }

    // First-seen row wins per email within a single sheet
    if (!byEmail.has(normalized)) {
      byEmail.set(normalized, {
        rowNumber: i + 2,
        name,
        perColumn,
      });
    }
  }

  return {
    result: tabResult,
    sheetId,
    spreadsheetId,
    byEmail,
    logicalColIndex,
    emailColIdx,
  };
}

export async function runSyncProSection(
  refreshToken: string,
  masterSheetId: string,
  section: ProSection
): Promise<ProSectionResult> {
  const masterTabName = section.masterTabName || `Pro: ${section.name}`;
  const result: ProSectionResult = {
    sectionId: section.id,
    sectionName: section.name,
    masterSpreadsheetUrl: "",
    masterTabName,
    totalUniqueEmails: 0,
    totalCellsFilled: 0,
    totalConflicts: 0,
    linkedSheets: [],
    columnStats: section.propagateColumns.map((c) => ({
      name: c.name,
      cellsFilled: 0,
      conflicts: 0,
      skippedSheets: 0,
    })),
    conflicts: [],
    presentInWritten: false,
  };

  if (!section.linkedSheets.length) {
    result.error = "Section needs at least one linked sheet.";
    return result;
  }

  const sheetsClient = getSheetsClient(refreshToken);
  const propagateColumnNames = section.propagateColumns.map((c) => c.name);

  // Read every linked sheet in parallel under MAX_CONCURRENCY
  const readTasks = section.linkedSheets.map(
    (linked) => async () => {
      try {
        return await readLinkedSheet(sheetsClient, linked, propagateColumnNames);
      } catch (err) {
        const tabResult: ProTabResult = {
          nickname: linked.nickname,
          url: linked.url,
          tabName: linked.tabName,
          rowsRead: 0,
          emailsFound: 0,
          cellsFilled: 0,
          error: err instanceof Error ? err.message : "Failed to read sheet",
        };
        return {
          result: tabResult,
          sheetId: -1,
          spreadsheetId: "",
          byEmail: new Map<string, SheetRowRecord>(),
          logicalColIndex: new Map<string, number>(),
          emailColIdx: -1,
        } as SheetReadOutcome;
      }
    }
  );
  const outcomes = await withConcurrencyLimit(readTasks, MAX_CONCURRENCY);
  // Preserve config order in the result list
  outcomes.sort(
    (a, b) =>
      section.linkedSheets.findIndex((l) => l.nickname === a.result.nickname) -
      section.linkedSheets.findIndex((l) => l.nickname === b.result.nickname)
  );

  result.linkedSheets = outcomes.map((o) => o.result);

  // Skipped-sheets count per logical column (sheets that mapped to null or
  // whose header lookup failed end up with idx=-1)
  for (const col of result.columnStats) {
    col.skippedSheets = outcomes.filter(
      (o) => (o.logicalColIndex.get(col.name) ?? -1) < 0
    ).length;
  }

  // Build cross-reference: email → array of {linkedSheetIdx, sheetId, rowNumber}
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
    for (const [email, record] of o.byEmail) {
      const list = crossRef.get(email) ?? [];
      list.push({
        linkedSheetIdx: i,
        sheetId: o.sheetId,
        rowNumber: record.rowNumber,
        spreadsheetId: o.spreadsheetId,
      });
      crossRef.set(email, list);
    }
  }
  result.totalUniqueEmails = crossRef.size;

  // Stash outcomes + crossRef on a local object that subsequent tasks
  // (propagation + present-in + master tab writes) will consume. For now
  // we early-return with read stats only; Tasks 6-7 fill in the rest.
  void crossRef;
  void writePresentInColumn;
  void mergeNamesForEmail;
  void HEADER_BG;
  void masterSheetId;
  void ({} as ProColumnStats);
  void ({} as ProConflict);
  return result;
}

export async function runSyncProBatch(
  refreshToken: string,
  masterSheetId: string,
  sections: ProSection[]
): Promise<ProBatchResult> {
  const results: ProSectionResult[] = [];
  for (const section of sections) {
    try {
      results.push(await runSyncProSection(refreshToken, masterSheetId, section));
    } catch (err) {
      results.push({
        sectionId: section.id,
        sectionName: section.name,
        masterSpreadsheetUrl: "",
        masterTabName: section.masterTabName || `Pro: ${section.name}`,
        totalUniqueEmails: 0,
        totalCellsFilled: 0,
        totalConflicts: 0,
        linkedSheets: [],
        columnStats: [],
        conflicts: [],
        presentInWritten: false,
        error: err instanceof Error ? err.message : "Section failed",
      });
    }
  }
  return { sections: results };
}
