import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { normalizeEmail } from "./email-utils";
import { sheets_v4 } from "googleapis";

/**
 * Duplicate Finder
 *
 * Scans selected tabs of a single source spreadsheet for duplicate values
 * in `Email` and `Telefono Cellulare` columns. Treats all picked tabs as
 * one combined dataset (a value seen in tab A then tab B is a duplicate).
 *
 * Highlighting rule:
 *  - First occurrence of a duplicated value  → light GREEN cell background.
 *  - Every subsequent occurrence            → light RED cell background.
 *  - Values that appear only once           → no formatting change.
 *
 * Only the offending CELL is colored (the email cell or the phone cell),
 * not the whole row. A row can be flagged twice (once on its email cell,
 * once on its phone cell) if both fields collide with other rows.
 *
 * Phone matching strips non-digits so "+39 333 1234567" and "393331234567"
 * dedupe to the same key. The cell value itself is never modified.
 *
 * Re-running clears any previous formatting in the Email and Phone columns
 * (across the picked tabs only) before repainting, so stale highlights from
 * previous runs do not linger.
 */

const SRC_EMAIL = "Email";
const SRC_PHONE = "Telefono Cellulare";

const GREEN_BG = { red: 0.78, green: 0.92, blue: 0.78 };
const RED_BG = { red: 1.0, green: 0.82, blue: 0.82 };
const WHITE_BG = { red: 1.0, green: 1.0, blue: 1.0 };

export interface DuplicateFinderTabResult {
  tabName: string;
  totalRows: number;
  emailDuplicateCells: number; // count of duplicate (non-first) email cells in this tab
  phoneDuplicateCells: number; // same for phone
  missingColumns: string[];
  error?: string;
}

export interface DuplicateFinderResult {
  spreadsheetUrl: string;
  totalDuplicateEmails: number; // distinct emails with >1 occurrence
  totalDuplicatePhones: number; // distinct phones with >1 occurrence
  totalDuplicateCells: number; // sum of (occurrences - 1) across all keys
  tabs: DuplicateFinderTabResult[];
}

interface CellLocation {
  tabName: string;
  sheetId: number;
  rowIndex: number; // 0-based (1 = first data row, since row 0 is header)
  colIndex: number; // 0-based
}

function normalizeHeader(s: unknown): string {
  return String(s ?? "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ");
}

function findHeaderIndex(headers: string[], target: string): number {
  const want = normalizeHeader(target);
  return headers.findIndex((h) => normalizeHeader(h) === want);
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizePhone(raw: string): string | null {
  const digits = raw.replace(/\D/g, "");
  if (!digits) return null;
  return digits;
}

function paintCellRequest(
  loc: CellLocation,
  bg: { red: number; green: number; blue: number }
): sheets_v4.Schema$Request {
  return {
    repeatCell: {
      range: {
        sheetId: loc.sheetId,
        startRowIndex: loc.rowIndex,
        endRowIndex: loc.rowIndex + 1,
        startColumnIndex: loc.colIndex,
        endColumnIndex: loc.colIndex + 1,
      },
      cell: {
        userEnteredFormat: { backgroundColor: bg },
      },
      fields: "userEnteredFormat.backgroundColor",
    },
  };
}

function clearColumnRequest(
  sheetId: number,
  colIndex: number,
  totalRows: number
): sheets_v4.Schema$Request | null {
  if (colIndex < 0 || totalRows < 2) return null;
  return {
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: 1, // skip header row
        endRowIndex: totalRows,
        startColumnIndex: colIndex,
        endColumnIndex: colIndex + 1,
      },
      cell: {
        userEnteredFormat: { backgroundColor: WHITE_BG },
      },
      fields: "userEnteredFormat.backgroundColor",
    },
  };
}

export async function runDuplicateFinder(
  refreshToken: string,
  spreadsheetId: string,
  tabs: string[]
): Promise<DuplicateFinderResult> {
  const sheets = getSheetsClient(refreshToken);

  // Resolve sheetId for every picked tab
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId })
  );
  const allTabs = meta.data.sheets ?? [];
  const tabInfoByName = new Map<string, number>();
  for (const t of tabs) {
    const found = allTabs.find((s) => s.properties?.title === t);
    if (
      !found ||
      found.properties?.sheetId === undefined ||
      found.properties?.sheetId === null
    ) {
      throw new Error(`Tab "${t}" not found in the spreadsheet`);
    }
    tabInfoByName.set(t, found.properties.sheetId);
  }

  // Batch-read every picked tab in one call
  const ranges = tabs.map((tab) => {
    const safe = `'${tab.replace(/'/g, "''")}'`;
    return `${safe}!A:ZZ`;
  });
  const batchRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges })
  );
  const valueRanges = batchRes.data.valueRanges ?? [];

  // Per-tab parsed snapshot (we need it twice: once to build dedupe maps,
  // once to issue clear-formatting requests over the full data range)
  interface TabSnapshot {
    tabName: string;
    sheetId: number;
    idxEmail: number;
    idxPhone: number;
    totalRows: number; // includes header
    missing: string[];
    result: DuplicateFinderTabResult;
  }

  const snapshots: TabSnapshot[] = [];
  const emailMap = new Map<string, CellLocation[]>();
  const phoneMap = new Map<string, CellLocation[]>();

  for (let i = 0; i < tabs.length; i++) {
    const tabName = tabs[i];
    const sheetId = tabInfoByName.get(tabName) as number;
    const rows = (valueRanges[i]?.values ?? []) as string[][];

    const tabResult: DuplicateFinderTabResult = {
      tabName,
      totalRows: 0,
      emailDuplicateCells: 0,
      phoneDuplicateCells: 0,
      missingColumns: [],
    };

    if (rows.length === 0) {
      snapshots.push({
        tabName,
        sheetId,
        idxEmail: -1,
        idxPhone: -1,
        totalRows: 0,
        missing: [SRC_EMAIL, SRC_PHONE],
        result: tabResult,
      });
      tabResult.missingColumns = [SRC_EMAIL, SRC_PHONE];
      continue;
    }

    const headers = (rows[0] ?? []).map((h) => String(h ?? ""));
    const idxEmail = findHeaderIndex(headers, SRC_EMAIL);
    const idxPhone = findHeaderIndex(headers, SRC_PHONE);

    const missing: string[] = [];
    if (idxEmail === -1) missing.push(SRC_EMAIL);
    if (idxPhone === -1) missing.push(SRC_PHONE);
    tabResult.missingColumns = missing;

    if (idxEmail === -1 && idxPhone === -1) {
      tabResult.error = `Neither "${SRC_EMAIL}" nor "${SRC_PHONE}" column found. Headers: ${headers
        .filter(Boolean)
        .join(", ")}`;
      snapshots.push({
        tabName,
        sheetId,
        idxEmail,
        idxPhone,
        totalRows: rows.length,
        missing,
        result: tabResult,
      });
      continue;
    }

    tabResult.totalRows = Math.max(0, rows.length - 1);

    // Scan every data row, push every valid email / phone into global maps.
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] ?? [];

      if (idxEmail >= 0) {
        const rawEmail = clean(row[idxEmail]);
        const norm = rawEmail ? normalizeEmail(rawEmail) : null;
        if (norm) {
          const list = emailMap.get(norm);
          const loc: CellLocation = {
            tabName,
            sheetId,
            rowIndex: r,
            colIndex: idxEmail,
          };
          if (list) list.push(loc);
          else emailMap.set(norm, [loc]);
        }
      }

      if (idxPhone >= 0) {
        const rawPhone = clean(row[idxPhone]);
        const phoneKey = rawPhone ? normalizePhone(rawPhone) : null;
        if (phoneKey) {
          const list = phoneMap.get(phoneKey);
          const loc: CellLocation = {
            tabName,
            sheetId,
            rowIndex: r,
            colIndex: idxPhone,
          };
          if (list) list.push(loc);
          else phoneMap.set(phoneKey, [loc]);
        }
      }
    }

    snapshots.push({
      tabName,
      sheetId,
      idxEmail,
      idxPhone,
      totalRows: rows.length,
      missing,
      result: tabResult,
    });
  }

  // Build batchUpdate requests
  const requests: sheets_v4.Schema$Request[] = [];

  // 1) Clear previous formatting in Email + Phone columns across all picked tabs.
  for (const snap of snapshots) {
    const clearEmail = clearColumnRequest(
      snap.sheetId,
      snap.idxEmail,
      snap.totalRows
    );
    if (clearEmail) requests.push(clearEmail);
    const clearPhone = clearColumnRequest(
      snap.sheetId,
      snap.idxPhone,
      snap.totalRows
    );
    if (clearPhone) requests.push(clearPhone);
  }

  // 2) Paint duplicates: first = green, rest = red.
  let totalDuplicateEmails = 0;
  let totalDuplicatePhones = 0;
  let totalDuplicateCells = 0;
  const perTabEmailDupes = new Map<string, number>();
  const perTabPhoneDupes = new Map<string, number>();

  for (const [, locations] of emailMap) {
    if (locations.length < 2) continue;
    totalDuplicateEmails++;
    requests.push(paintCellRequest(locations[0], GREEN_BG));
    for (let i = 1; i < locations.length; i++) {
      requests.push(paintCellRequest(locations[i], RED_BG));
      totalDuplicateCells++;
      const k = locations[i].tabName;
      perTabEmailDupes.set(k, (perTabEmailDupes.get(k) ?? 0) + 1);
    }
  }

  for (const [, locations] of phoneMap) {
    if (locations.length < 2) continue;
    totalDuplicatePhones++;
    requests.push(paintCellRequest(locations[0], GREEN_BG));
    for (let i = 1; i < locations.length; i++) {
      requests.push(paintCellRequest(locations[i], RED_BG));
      totalDuplicateCells++;
      const k = locations[i].tabName;
      perTabPhoneDupes.set(k, (perTabPhoneDupes.get(k) ?? 0) + 1);
    }
  }

  // Apply formatting (only if there's something to do)
  if (requests.length > 0) {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      })
    );
  }

  // Fold per-tab counts back into snapshots
  for (const snap of snapshots) {
    snap.result.emailDuplicateCells = perTabEmailDupes.get(snap.tabName) ?? 0;
    snap.result.phoneDuplicateCells = perTabPhoneDupes.get(snap.tabName) ?? 0;
  }

  return {
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
    totalDuplicateEmails,
    totalDuplicatePhones,
    totalDuplicateCells,
    tabs: snapshots.map((s) => s.result),
  };
}
