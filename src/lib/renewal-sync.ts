import { getSheetsClient } from "./google-auth";
import {
  columnIndexToLetter,
  findEmailColumnIndex,
  normalizeEmail,
} from "./email-utils";
import { withRetry } from "./retry";
import { sheets_v4 } from "googleapis";

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };
const RED_BG = { red: 1.0, green: 0.82, blue: 0.82 };
const WHITE_BG = { red: 1.0, green: 1.0, blue: 1.0 };

const FIELDS = ["phone", "courseName", "startDate", "renewalDate", "setterAssigned"] as const;
type FieldKey = (typeof FIELDS)[number];

const CANONICAL_HEADERS: Record<FieldKey, string> = {
  phone: "Phone",
  courseName: "Course name",
  startDate: "Start date",
  renewalDate: "Renewal Date",
  setterAssigned: "Setter assigned",
};

// Accept common variants (English + Italian) when locating columns
const HEADER_ALIASES: Record<FieldKey, string[]> = {
  phone: ["phone", "phone number", "telefono", "mobile", "cellulare"],
  courseName: ["course name", "course", "corso", "nome corso"],
  startDate: ["start date", "start", "data inizio", "inizio", "data di inizio"],
  renewalDate: ["renewal date", "renewal", "data rinnovo", "rinnovo", "data di rinnovo"],
  setterAssigned: ["setter assigned", "setter", "assigned setter", "assegnato"],
};

type LookupData = Record<FieldKey, string>;

export interface RenewalSyncTabResult {
  tabName: string;
  totalRows: number;
  matched: number;
  unmatched: number;
  pastRenewals: number;
  error?: string;
}

export interface RenewalSyncResult {
  spreadsheetUrl: string;
  tabs: RenewalSyncTabResult[];
}

function normalizeHeader(s: string): string {
  return s.toLowerCase().trim().replace(/\s+/g, " ");
}

function findHeaderIndex(headers: string[], aliases: string[]): number {
  const norm = headers.map((h) => normalizeHeader(h?.toString() ?? ""));
  const normAliases = aliases.map(normalizeHeader);
  for (const alias of normAliases) {
    const idx = norm.indexOf(alias);
    if (idx !== -1) return idx;
  }
  return -1;
}

function parseFlexibleDate(raw: string): Date | null {
  const str = raw.trim();
  if (!str) return null;

  // ISO YYYY-MM-DD (or YYYY/MM/DD)
  const iso = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (iso) {
    const d = new Date(parseInt(iso[1]), parseInt(iso[2]) - 1, parseInt(iso[3]));
    if (!isNaN(d.getTime())) return d;
  }

  // DD/MM/YYYY (European default — user's data is Italian)
  // Falls back to MM/DD/YYYY if first part > 12
  const parts = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (parts) {
    let a = parseInt(parts[1]);
    let b = parseInt(parts[2]);
    let year = parseInt(parts[3]);
    if (year < 100) year += 2000;

    // If first > 12 → must be DD/MM. Otherwise default to DD/MM (Italian).
    // Only swap when first ≤ 12 AND second > 12 (unambiguously MM/DD).
    if (a <= 12 && b > 12) {
      [a, b] = [b, a];
    }
    const d = new Date(year, b - 1, a);
    if (!isNaN(d.getTime())) return d;
  }

  const fallback = new Date(str);
  if (!isNaN(fallback.getTime())) return fallback;

  return null;
}

function isPastDate(raw: string): boolean {
  const d = parseFlexibleDate(raw);
  if (!d) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return d.getTime() < today.getTime();
}

async function buildLookupMap(
  sheets: sheets_v4.Sheets,
  lookupSpreadsheetId: string,
  lookupTabs: string[]
): Promise<{ map: Map<string, LookupData>; missingHeaders: string[] }> {
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: lookupSpreadsheetId })
  );
  const allTabs = meta.data.sheets ?? [];

  const map = new Map<string, LookupData>();
  const missingHeaders = new Set<string>();

  for (const tabName of lookupTabs) {
    const tab = allTabs.find((t) => t.properties?.title === tabName);
    if (!tab) {
      throw new Error(`Lookup tab "${tabName}" not found`);
    }
    const safeTab = `'${tabName.replace(/'/g, "''")}'`;

    const headerRes = await withRetry(() =>
      sheets.spreadsheets.values.get({
        spreadsheetId: lookupSpreadsheetId,
        range: `${safeTab}!1:1`,
      })
    );
    const headers = (headerRes.data.values?.[0] ?? []) as string[];

    const emailIdx = findEmailColumnIndex(headers);
    if (emailIdx === -1) {
      throw new Error(
        `Lookup tab "${tabName}" has no Email column. Headers: ${headers.join(", ")}`
      );
    }

    const fieldIdx: Partial<Record<FieldKey, number>> = {};
    for (const f of FIELDS) {
      const idx = findHeaderIndex(headers, HEADER_ALIASES[f]);
      if (idx === -1) {
        missingHeaders.add(CANONICAL_HEADERS[f]);
      } else {
        fieldIdx[f] = idx;
      }
    }

    // Read all needed columns in one batchGet
    const ranges: string[] = [];
    const idxMap: Record<string, number> = {};
    idxMap.email = ranges.length;
    ranges.push(`${safeTab}!${columnIndexToLetter(emailIdx)}2:${columnIndexToLetter(emailIdx)}`);
    for (const f of FIELDS) {
      if (fieldIdx[f] !== undefined) {
        idxMap[f] = ranges.length;
        const letter = columnIndexToLetter(fieldIdx[f]!);
        ranges.push(`${safeTab}!${letter}2:${letter}`);
      }
    }

    const dataRes = await withRetry(() =>
      sheets.spreadsheets.values.batchGet({
        spreadsheetId: lookupSpreadsheetId,
        ranges,
      })
    );
    const valueRanges = dataRes.data.valueRanges ?? [];
    const emailRows = valueRanges[idxMap.email]?.values ?? [];

    const fieldRows: Partial<Record<FieldKey, string[][]>> = {};
    for (const f of FIELDS) {
      if (idxMap[f] !== undefined) {
        fieldRows[f] = (valueRanges[idxMap[f]]?.values ?? []) as string[][];
      }
    }

    for (let i = 0; i < emailRows.length; i++) {
      const rawEmail = emailRows[i]?.[0];
      if (!rawEmail) continue;
      const email = normalizeEmail(String(rawEmail));
      if (!email) continue;
      if (map.has(email)) continue; // first occurrence wins

      const data: LookupData = {
        phone: "",
        courseName: "",
        startDate: "",
        renewalDate: "",
        setterAssigned: "",
      };
      for (const f of FIELDS) {
        const cell = fieldRows[f]?.[i]?.[0];
        if (cell !== undefined && cell !== null) {
          data[f] = String(cell).trim();
        }
      }
      map.set(email, data);
    }
  }

  return { map, missingHeaders: Array.from(missingHeaders) };
}

async function processSourceTab(
  sheets: sheets_v4.Sheets,
  sourceSpreadsheetId: string,
  allSourceTabs: sheets_v4.Schema$Sheet[],
  tabName: string,
  lookupMap: Map<string, LookupData>
): Promise<RenewalSyncTabResult> {
  const tab = allSourceTabs.find((t) => t.properties?.title === tabName);
  if (!tab?.properties) {
    throw new Error(`Tab "${tabName}" not found`);
  }
  const sheetId = tab.properties.sheetId;
  if (typeof sheetId !== "number") {
    throw new Error(`Tab "${tabName}" has no sheetId`);
  }

  const safeTab = `'${tabName.replace(/'/g, "''")}'`;

  const headerRes = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: sourceSpreadsheetId,
      range: `${safeTab}!1:1`,
    })
  );
  const headers = (headerRes.data.values?.[0] ?? []) as string[];

  const emailIdx = findEmailColumnIndex(headers);
  if (emailIdx === -1) {
    throw new Error(
      `No Email column in "${tabName}". Headers: ${headers.join(", ")}`
    );
  }

  const emailLetter = columnIndexToLetter(emailIdx);
  const emailsRes = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: sourceSpreadsheetId,
      range: `${safeTab}!${emailLetter}2:${emailLetter}`,
    })
  );
  const emailRows = (emailsRes.data.values ?? []) as string[][];

  // Resolve column index for each target field: reuse existing column or append at end
  let lastFilledIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]?.toString().trim()) lastFilledIdx = i;
  }

  const targetCol: Record<FieldKey, number> = {} as Record<FieldKey, number>;
  const headerNeeded: Record<FieldKey, boolean> = {} as Record<FieldKey, boolean>;
  let nextAppendIdx = lastFilledIdx + 1;
  for (const f of FIELDS) {
    const existing = findHeaderIndex(headers, HEADER_ALIASES[f]);
    if (existing !== -1) {
      targetCol[f] = existing;
      headerNeeded[f] = false;
    } else {
      targetCol[f] = nextAppendIdx++;
      headerNeeded[f] = true;
    }
  }

  // Build per-row outcome
  interface RowOutcome {
    email: string | null;
    matched: boolean;
    data: LookupData | null;
    renewalPast: boolean;
  }
  const outcomes: RowOutcome[] = [];
  let matched = 0;
  let unmatched = 0;
  let pastRenewals = 0;
  let totalRows = 0;

  for (let i = 0; i < emailRows.length; i++) {
    const raw = emailRows[i]?.[0];
    if (!raw) {
      outcomes.push({ email: null, matched: false, data: null, renewalPast: false });
      continue;
    }
    const email = normalizeEmail(String(raw));
    if (!email) {
      outcomes.push({ email: null, matched: false, data: null, renewalPast: false });
      continue;
    }
    totalRows++;
    const data = lookupMap.get(email) ?? null;
    if (data) {
      matched++;
      const renewalPast = isPastDate(data.renewalDate);
      if (renewalPast) pastRenewals++;
      outcomes.push({ email, matched: true, data, renewalPast });
    } else {
      unmatched++;
      outcomes.push({ email, matched: false, data: null, renewalPast: false });
    }
  }

  const requests: sheets_v4.Schema$Request[] = [];

  // 1. Headers (only for newly-appended columns)
  for (const f of FIELDS) {
    if (!headerNeeded[f]) continue;
    requests.push({
      updateCells: {
        rows: [
          {
            values: [
              {
                userEnteredValue: { stringValue: CANONICAL_HEADERS[f] },
                userEnteredFormat: {
                  textFormat: { bold: true },
                  backgroundColor: HEADER_BG,
                },
              },
            ],
          },
        ],
        fields:
          "userEnteredValue,userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor",
        start: { sheetId, rowIndex: 0, columnIndex: targetCol[f] },
      },
    });
  }

  // 2. Per-field column writes (value + background)
  if (outcomes.length > 0) {
    for (const f of FIELDS) {
      requests.push({
        updateCells: {
          rows: outcomes.map((o) => {
            const value = o.data ? o.data[f] : "";
            const background =
              f === "renewalDate" && o.renewalPast ? RED_BG : WHITE_BG;
            return {
              values: [
                {
                  userEnteredValue: { stringValue: value },
                  userEnteredFormat: { backgroundColor: background },
                },
              ],
            };
          }),
          fields: "userEnteredValue,userEnteredFormat.backgroundColor",
          start: { sheetId, rowIndex: 1, columnIndex: targetCol[f] },
        },
      });
    }
  }

  const CHUNK_SIZE = 5;
  for (let i = 0; i < requests.length; i += CHUNK_SIZE) {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId: sourceSpreadsheetId,
        requestBody: { requests: requests.slice(i, i + CHUNK_SIZE) },
      })
    );
  }

  return {
    tabName,
    totalRows,
    matched,
    unmatched,
    pastRenewals,
  };
}

export async function runRenewalSync(
  refreshToken: string,
  sourceSpreadsheetId: string,
  sourceTabs: string[],
  lookupSpreadsheetId: string,
  lookupTabs: string[]
): Promise<RenewalSyncResult> {
  const sheets = getSheetsClient(refreshToken);

  const { map: lookupMap } = await buildLookupMap(
    sheets,
    lookupSpreadsheetId,
    lookupTabs
  );

  const sourceMeta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: sourceSpreadsheetId })
  );
  const allSourceTabs = sourceMeta.data.sheets ?? [];

  const results: RenewalSyncTabResult[] = [];
  for (const tabName of sourceTabs) {
    try {
      const result = await processSourceTab(
        sheets,
        sourceSpreadsheetId,
        allSourceTabs,
        tabName,
        lookupMap
      );
      results.push(result);
    } catch (err) {
      results.push({
        tabName,
        totalRows: 0,
        matched: 0,
        unmatched: 0,
        pastRenewals: 0,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  return {
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${sourceSpreadsheetId}/edit`,
    tabs: results,
  };
}
