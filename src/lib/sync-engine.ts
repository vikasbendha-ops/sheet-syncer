import { getLinkedSheets, updateLastSyncedBatch } from "./config-store";
import { readEmailsFromSheet } from "./sheets-reader";
import { writeMasterSheet } from "./sheets-writer";
import { writePresentInColumn, PresentInCell } from "./present-in-writer";
import { extractSpreadsheetId } from "./url-parser";
import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { SyncResult, MasterSheetData } from "@/types";
import { sheets_v4 } from "googleapis";

// Lowered from 5 to stay under Google Sheets' 60 read/min per-user quota
const MAX_CONCURRENCY = 3;

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

interface SheetSyncData {
  key: string; // stable identifier: `${url}#${tabName}`
  nickname: string;
  url: string;
  tabName: string;
  spreadsheetId: string;
  sheetId: number;
  emails: Map<string, number[]>;
  names: Map<string, string[]>;
  presentInColumnIndex: number | null;
  lastColumnIndex: number;
  index: number; // original config index for sorting + updateLastSynced
}

export async function runFullSync(
  refreshToken: string,
  masterSheetId: string
): Promise<SyncResult> {
  const timestamp = new Date().toISOString();
  const errors: string[] = [];

  const linkedSheets = await getLinkedSheets(refreshToken, masterSheetId);

  if (linkedSheets.length === 0) {
    return {
      success: true,
      sheetsProcessed: 0,
      totalEmails: 0,
      errors: [],
      timestamp,
    };
  }

  const sheetEmails: SheetSyncData[] = [];

  // Pre-fetch tab metadata once per unique spreadsheet (big perf win + avoids Vercel timeout)
  const tabMetadataCache = new Map<string, sheets_v4.Schema$Sheet[]>();
  const uniqueSpreadsheetIds = new Set<string>();
  for (const sheet of linkedSheets) {
    try {
      uniqueSpreadsheetIds.add(extractSpreadsheetId(sheet.url));
    } catch {
      // ignore invalid URLs here; they'll error in the read task
    }
  }

  const sheetsClient = getSheetsClient(refreshToken);
  await Promise.all(
    Array.from(uniqueSpreadsheetIds).map(async (sid) => {
      try {
        const res = await withRetry(() =>
          sheetsClient.spreadsheets.get({ spreadsheetId: sid })
        );
        tabMetadataCache.set(sid, res.data.sheets ?? []);
      } catch {
        // missing cache → reader will fetch on its own
      }
    })
  );

  const readTasks = linkedSheets.map((sheet, index) => async () => {
    try {
      const spreadsheetId = extractSpreadsheetId(sheet.url);
      const result = await readEmailsFromSheet(
        refreshToken,
        spreadsheetId,
        sheet.tabName,
        sheet.emailColumn,
        tabMetadataCache.get(spreadsheetId)
      );
      sheetEmails.push({
        key: `${sheet.url}#${result.resolvedTabName}`,
        nickname: sheet.nickname,
        url: sheet.url,
        tabName: result.resolvedTabName,
        spreadsheetId,
        sheetId: result.sheetId,
        emails: result.emails,
        names: result.names,
        presentInColumnIndex: result.presentInColumnIndex,
        lastColumnIndex: result.lastColumnIndex,
        index,
      });
    } catch (err) {
      const msg = `Failed to read "${sheet.nickname}": ${err instanceof Error ? err.message : String(err)}`;
      errors.push(msg);
    }
  });

  await withConcurrencyLimit(readTasks, MAX_CONCURRENCY);

  sheetEmails.sort((a, b) => a.index - b.index);

  // Batch update lastSynced for all successfully read sheets in a single API call
  try {
    const successIndices = sheetEmails.map((s) => s.index);
    await updateLastSyncedBatch(
      refreshToken,
      masterSheetId,
      successIndices,
      timestamp
    );
  } catch (err) {
    errors.push(
      `Failed to update lastSynced timestamps: ${err instanceof Error ? err.message : String(err)}`
    );
  }

  // Build cross-reference map: email → sheets that contain it.
  // Track the target sheetId + row so hyperlinks can deep-link to the exact row.
  const crossRef = new Map<
    string,
    Array<{
      nickname: string;
      sheetKey: string;
      spreadsheetId: string;
      sheetId: number;
      rowNumber: number; // first row where this email appears in that sheet
    }>
  >();
  for (const sheet of sheetEmails) {
    for (const [email, rowNumbers] of sheet.emails.entries()) {
      let list = crossRef.get(email);
      if (!list) {
        list = [];
        crossRef.set(email, list);
      }
      list.push({
        nickname: sheet.nickname,
        sheetKey: sheet.key,
        spreadsheetId: sheet.spreadsheetId,
        sheetId: sheet.sheetId,
        rowNumber: rowNumbers[0],
      });
    }
  }

  // Write "Present In" column back to each source sheet
  const writeTasks = sheetEmails.map((sheet) => async () => {
    try {
      const cellData: PresentInCell[] = [];
      for (const [email, rowNumbers] of sheet.emails.entries()) {
        const others = (crossRef.get(email) ?? []).filter(
          (x) => x.sheetKey !== sheet.key
        );
        if (others.length === 0) continue;
        for (const rowNumber of rowNumbers) {
          cellData.push({
            rowIndex: rowNumber - 1,
            links: others.map((o) => ({
              text: o.nickname,
              url: `https://docs.google.com/spreadsheets/d/${o.spreadsheetId}/edit#gid=${o.sheetId}&range=A${o.rowNumber}`,
            })),
          });
        }
      }

      const columnIndex =
        sheet.presentInColumnIndex !== null
          ? sheet.presentInColumnIndex
          : sheet.lastColumnIndex + 1;
      const headerNeeded = sheet.presentInColumnIndex === null;

      await writePresentInColumn(
        refreshToken,
        sheet.spreadsheetId,
        sheet.sheetId,
        columnIndex,
        headerNeeded,
        cellData
      );
    } catch (err) {
      const msg = `Failed to write "Present In" to "${sheet.nickname}": ${err instanceof Error ? err.message : String(err)}`;
      errors.push(msg);
    }
  });

  await withConcurrencyLimit(writeTasks, MAX_CONCURRENCY);

  // Write master sheet
  const allEmails = new Set<string>();
  for (const sheet of sheetEmails) {
    for (const email of sheet.emails.keys()) {
      allEmails.add(email);
    }
  }

  // Merge names across source sheets.
  // Strategy:
  //  1. Count every name occurrence (case-insensitive) across all sheets
  //  2. If name A is a word-prefix of name B (e.g. "Laura" ⊂ "Laura Pegoraro"),
  //     fold A's count into B so the more complete name wins
  //  3. Pick the winner by count → most words → first-seen order
  const mergedNames = new Map<string, string>();
  for (const email of allEmails) {
    interface Candidate {
      name: string;
      words: string[];
      count: number;
      mergedCount: number;
      firstSeenIdx: number;
    }
    const candidates = new Map<string, Candidate>();
    let nextIdx = 0;

    for (const sheet of sheetEmails) {
      const list = sheet.names.get(email) ?? [];
      for (const name of list) {
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
    }

    if (candidates.size === 0) continue;

    // Fold prefix-candidates into their supersets (e.g. "Laura" counts toward "Laura Pegoraro")
    const list = Array.from(candidates.values());
    for (const a of list) {
      for (const b of list) {
        if (a === b) continue;
        if (b.words.length <= a.words.length) continue;
        const isPrefix = a.words.every((w, i) => b.words[i] === w);
        if (isPrefix) {
          b.mergedCount += a.count;
        }
      }
    }

    list.sort((a, b) => {
      if (b.mergedCount !== a.mergedCount) return b.mergedCount - a.mergedCount;
      if (b.words.length !== a.words.length) return b.words.length - a.words.length;
      return a.firstSeenIdx - b.firstSeenIdx;
    });

    mergedNames.set(email, list[0].name);
  }

  const sortedEmails = Array.from(allEmails).sort();
  const headers = ["Name", "Email", ...sheetEmails.map((s) => s.nickname)];
  const rows = sortedEmails.map((email) => [
    mergedNames.get(email) ?? "",
    email,
    ...sheetEmails.map((s) => (s.emails.has(email) ? "✅" : "❌")),
  ]);

  const masterData: MasterSheetData = { headers, rows };
  try {
    await writeMasterSheet(refreshToken, masterSheetId, masterData);
  } catch (err) {
    const msg = `Failed to write master sheet: ${err instanceof Error ? err.message : String(err)}`;
    errors.push(msg);
  }

  return {
    success: errors.length === 0,
    sheetsProcessed: sheetEmails.length,
    totalEmails: sortedEmails.length,
    errors,
    timestamp,
  };
}
