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

export async function runFullSync(refreshToken: string): Promise<SyncResult> {
  const timestamp = new Date().toISOString();
  const errors: string[] = [];

  const linkedSheets = await getLinkedSheets(refreshToken);

  if (linkedSheets.length === 0) {
    return {
      success: true,
      sheetsProcessed: 0,
      totalEmails: 0,
      errors: [],
      timestamp,
    };
  }

  const sheetEmails: { nickname: string; emails: Set<string>; index: number }[] = [];

  const tasks = linkedSheets.map((sheet, index) => async () => {
    try {
      const spreadsheetId = extractSpreadsheetId(sheet.url);
      const result = await readEmailsFromSheet(
        refreshToken,
        spreadsheetId,
        sheet.tabName,
        sheet.emailColumn
      );
      sheetEmails.push({
        nickname: sheet.nickname,
        emails: result.emails,
        index,
      });
      await updateLastSynced(refreshToken, index, timestamp);
    } catch (err) {
      const msg = `Failed to read "${sheet.nickname}": ${err instanceof Error ? err.message : String(err)}`;
      errors.push(msg);
    }
  });

  await withConcurrencyLimit(tasks, MAX_CONCURRENCY);

  sheetEmails.sort((a, b) => a.index - b.index);

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
    ...sheetEmails.map((s) => (s.emails.has(email) ? "✅" : "❌")),
  ]);

  const masterData: MasterSheetData = { headers, rows };

  await writeMasterSheet(refreshToken, masterData);

  return {
    success: errors.length === 0,
    sheetsProcessed: sheetEmails.length,
    totalEmails: sortedEmails.length,
    errors,
    timestamp,
  };
}
