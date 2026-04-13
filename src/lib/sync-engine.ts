import { getLinkedSheets, updateLastSynced } from "./config-store";
import { readEmailsFromSheet } from "./sheets-reader";
import { writeMasterSheet } from "./sheets-writer";
import { extractSpreadsheetId } from "./url-parser";
import { SyncResult, MasterSheetData } from "@/types";

const MAX_CONCURRENCY = 5;

async function withConcurrencyLimit<T>(
  tasks: (() => Promise<T>)[],
  limit: number
): Promise<T[]> {
  const results: T[] = [];
  const executing: Promise<void>[] = [];

  for (const task of tasks) {
    const p = task().then((result) => {
      results.push(result);
    });
    executing.push(p);

    if (executing.length >= limit) {
      await Promise.race(executing);
      executing.splice(
        executing.findIndex((e) => e === p),
        1
      );
    }
  }

  await Promise.all(executing);
  return results;
}

export async function runFullSync(): Promise<SyncResult> {
  const timestamp = new Date().toISOString();
  const errors: string[] = [];

  // 1. Read linked sheets config
  const linkedSheets = await getLinkedSheets();

  if (linkedSheets.length === 0) {
    return {
      success: true,
      sheetsProcessed: 0,
      totalEmails: 0,
      errors: [],
      timestamp,
    };
  }

  // 2. Read emails from each linked sheet
  const sheetEmails: { nickname: string; emails: Set<string>; index: number }[] = [];

  const tasks = linkedSheets.map((sheet, index) => async () => {
    try {
      const spreadsheetId = extractSpreadsheetId(sheet.url);
      const result = await readEmailsFromSheet(spreadsheetId, sheet.emailColumn);
      sheetEmails.push({
        nickname: sheet.nickname,
        emails: result.emails,
        index,
      });
      await updateLastSynced(index, timestamp);
    } catch (err) {
      const msg = `Failed to read "${sheet.nickname}": ${err instanceof Error ? err.message : String(err)}`;
      errors.push(msg);
    }
  });

  await withConcurrencyLimit(tasks, MAX_CONCURRENCY);

  // Sort by original index to maintain column order
  sheetEmails.sort((a, b) => a.index - b.index);

  // 3. Merge — build master data
  const allEmails = new Set<string>();
  for (const sheet of sheetEmails) {
    for (const email of sheet.emails) {
      allEmails.add(email);
    }
  }

  const sortedEmails = Array.from(allEmails).sort();

  const headers = ["Email", ...sheetEmails.map((s) => s.nickname)];
  const rows = sortedEmails.map((email) => [
    email,
    ...sheetEmails.map((s) => (s.emails.has(email) ? "YES" : "NO")),
  ]);

  const masterData: MasterSheetData = { headers, rows };

  // 4. Write master sheet
  await writeMasterSheet(masterData);

  return {
    success: errors.length === 0,
    sheetsProcessed: sheetEmails.length,
    totalEmails: sortedEmails.length,
    errors,
    timestamp,
  };
}
