import { getSheetsClient } from "./google-auth";
import {
  columnIndexToLetter,
  findEmailColumnIndex,
  normalizeEmail,
} from "./email-utils";
import { withRetry } from "./retry";
import { sheets_v4 } from "googleapis";

const BIZ_TUTOR_HEADER = "BIZ TUTOR";
const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

const BIZ_TUTOR_ALIASES = [
  "biz tutor",
  "biztutor",
  "biz_tutor",
  "business tutor",
];

const PHONE_ALIASES = [
  "phone",
  "mobile",
  "phone number",
  "mobile number",
  "telefono",
  "cellulare",
  "tel",
  "telephone",
  "numero di telefono",
];

export interface BizTutorTabResult {
  tabName: string;
  totalRows: number;
  matchedByEmail: number;
  matchedByPhone: number;
  unmatched: number;
  error?: string;
}

export interface BizTutorSyncResult {
  spreadsheetUrl: string;
  tabs: BizTutorTabResult[];
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

function phoneKeys(raw: string): string[] {
  const digits = raw.replace(/\D/g, "");
  if (digits.length < 7) return [];
  const set = new Set<string>([digits]);
  // Also key on last 10 digits to match across country-code variants
  if (digits.length > 10) set.add(digits.slice(-10));
  return Array.from(set);
}

async function buildLookupMap(
  sheets: sheets_v4.Sheets,
  lookupSpreadsheetId: string,
  lookupTabs: string[]
): Promise<{
  emailToTutor: Map<string, string>;
  phoneToTutor: Map<string, string>;
}> {
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: lookupSpreadsheetId })
  );
  const allTabs = meta.data.sheets ?? [];

  const emailToTutor = new Map<string, string>();
  const phoneToTutor = new Map<string, string>();

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

    const tutorIdx = findHeaderIndex(headers, BIZ_TUTOR_ALIASES);
    if (tutorIdx === -1) {
      throw new Error(
        `Lookup tab "${tabName}" has no BIZ TUTOR column. Headers: ${headers.join(", ")}`
      );
    }
    const emailIdx = findEmailColumnIndex(headers);
    const phoneIdx = findHeaderIndex(headers, PHONE_ALIASES);
    if (emailIdx === -1 && phoneIdx === -1) {
      throw new Error(
        `Lookup tab "${tabName}" has neither Email nor Phone column to match on`
      );
    }

    const ranges: string[] = [];
    const rangeIdx: { tutor: number; email?: number; phone?: number } = {
      tutor: 0,
    };
    ranges.push(
      `${safeTab}!${columnIndexToLetter(tutorIdx)}2:${columnIndexToLetter(tutorIdx)}`
    );
    if (emailIdx !== -1) {
      rangeIdx.email = ranges.length;
      const letter = columnIndexToLetter(emailIdx);
      ranges.push(`${safeTab}!${letter}2:${letter}`);
    }
    if (phoneIdx !== -1) {
      rangeIdx.phone = ranges.length;
      const letter = columnIndexToLetter(phoneIdx);
      ranges.push(`${safeTab}!${letter}2:${letter}`);
    }

    const dataRes = await withRetry(() =>
      sheets.spreadsheets.values.batchGet({
        spreadsheetId: lookupSpreadsheetId,
        ranges,
      })
    );
    const valueRanges = dataRes.data.valueRanges ?? [];
    const tutorRows = (valueRanges[rangeIdx.tutor]?.values ?? []) as string[][];
    const emailRows =
      rangeIdx.email !== undefined
        ? ((valueRanges[rangeIdx.email]?.values ?? []) as string[][])
        : [];
    const phoneRows =
      rangeIdx.phone !== undefined
        ? ((valueRanges[rangeIdx.phone]?.values ?? []) as string[][])
        : [];

    const rowCount = Math.max(
      tutorRows.length,
      emailRows.length,
      phoneRows.length
    );
    for (let i = 0; i < rowCount; i++) {
      const tutorRaw = tutorRows[i]?.[0];
      if (!tutorRaw) continue;
      const tutor = String(tutorRaw).trim();
      if (!tutor) continue;

      const emailRaw = emailRows[i]?.[0];
      if (emailRaw) {
        const email = normalizeEmail(String(emailRaw));
        if (email && !emailToTutor.has(email)) {
          emailToTutor.set(email, tutor);
        }
      }

      const phoneRaw = phoneRows[i]?.[0];
      if (phoneRaw) {
        for (const key of phoneKeys(String(phoneRaw))) {
          if (!phoneToTutor.has(key)) {
            phoneToTutor.set(key, tutor);
          }
        }
      }
    }
  }

  return { emailToTutor, phoneToTutor };
}

async function processSourceTab(
  sheets: sheets_v4.Sheets,
  sourceSpreadsheetId: string,
  allSourceTabs: sheets_v4.Schema$Sheet[],
  tabName: string,
  emailToTutor: Map<string, string>,
  phoneToTutor: Map<string, string>
): Promise<BizTutorTabResult> {
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
  const phoneIdx = findHeaderIndex(headers, PHONE_ALIASES);
  if (emailIdx === -1 && phoneIdx === -1) {
    throw new Error(
      `Source tab "${tabName}" has neither Email nor Phone column to look up`
    );
  }

  // Locate or append the BIZ TUTOR column
  let tutorIdx = findHeaderIndex(headers, BIZ_TUTOR_ALIASES);
  let headerNeeded = false;
  if (tutorIdx === -1) {
    let lastFilledIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i]?.toString().trim()) lastFilledIdx = i;
    }
    tutorIdx = lastFilledIdx + 1;
    headerNeeded = true;
  }

  // Read source columns: email, phone, existing tutor (to preserve unmatched values)
  const ranges: string[] = [];
  const rangeIdx: { email?: number; phone?: number; tutor: number } = {
    tutor: 0,
  };
  const tutorLetter = columnIndexToLetter(tutorIdx);
  ranges.push(`${safeTab}!${tutorLetter}2:${tutorLetter}`);
  if (emailIdx !== -1) {
    rangeIdx.email = ranges.length;
    const letter = columnIndexToLetter(emailIdx);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }
  if (phoneIdx !== -1) {
    rangeIdx.phone = ranges.length;
    const letter = columnIndexToLetter(phoneIdx);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }

  const dataRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId: sourceSpreadsheetId,
      ranges,
    })
  );
  const valueRanges = dataRes.data.valueRanges ?? [];
  const existingTutorRows = (valueRanges[rangeIdx.tutor]?.values ?? []) as string[][];
  const emailRows =
    rangeIdx.email !== undefined
      ? ((valueRanges[rangeIdx.email]?.values ?? []) as string[][])
      : [];
  const phoneRows =
    rangeIdx.phone !== undefined
      ? ((valueRanges[rangeIdx.phone]?.values ?? []) as string[][])
      : [];

  const rowCount = Math.max(
    existingTutorRows.length,
    emailRows.length,
    phoneRows.length
  );

  let matchedByEmail = 0;
  let matchedByPhone = 0;
  let unmatched = 0;
  let totalRows = 0;

  const outputs: string[] = new Array(rowCount).fill("");

  for (let i = 0; i < rowCount; i++) {
    const existing = existingTutorRows[i]?.[0]
      ? String(existingTutorRows[i][0])
      : "";

    const emailRaw = emailRows[i]?.[0];
    const phoneRaw = phoneRows[i]?.[0];
    const hasIdentifier = !!emailRaw || !!phoneRaw;
    if (hasIdentifier) totalRows++;

    let resolved: string | null = null;
    if (emailRaw) {
      const email = normalizeEmail(String(emailRaw));
      if (email && emailToTutor.has(email)) {
        resolved = emailToTutor.get(email)!;
        matchedByEmail++;
      }
    }
    if (resolved === null && phoneRaw) {
      for (const key of phoneKeys(String(phoneRaw))) {
        if (phoneToTutor.has(key)) {
          resolved = phoneToTutor.get(key)!;
          matchedByPhone++;
          break;
        }
      }
    }

    if (resolved !== null) {
      outputs[i] = resolved;
    } else {
      // Preserve any existing value rather than clobbering with empty
      outputs[i] = existing;
      if (hasIdentifier) unmatched++;
    }
  }

  const requests: sheets_v4.Schema$Request[] = [];

  if (headerNeeded) {
    requests.push({
      updateCells: {
        rows: [
          {
            values: [
              {
                userEnteredValue: { stringValue: BIZ_TUTOR_HEADER },
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
        start: { sheetId, rowIndex: 0, columnIndex: tutorIdx },
      },
    });
  }

  if (rowCount > 0) {
    requests.push({
      updateCells: {
        rows: outputs.map((value) => ({
          values: [{ userEnteredValue: { stringValue: value } }],
        })),
        fields: "userEnteredValue",
        start: { sheetId, rowIndex: 1, columnIndex: tutorIdx },
      },
    });
  }

  if (requests.length > 0) {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId: sourceSpreadsheetId,
        requestBody: { requests },
      })
    );
  }

  return {
    tabName,
    totalRows,
    matchedByEmail,
    matchedByPhone,
    unmatched,
  };
}

export async function runBizTutorSync(
  refreshToken: string,
  sourceSpreadsheetId: string,
  sourceTabs: string[],
  lookupSpreadsheetId: string,
  lookupTabs: string[]
): Promise<BizTutorSyncResult> {
  const sheets = getSheetsClient(refreshToken);

  const { emailToTutor, phoneToTutor } = await buildLookupMap(
    sheets,
    lookupSpreadsheetId,
    lookupTabs
  );

  const sourceMeta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: sourceSpreadsheetId })
  );
  const allSourceTabs = sourceMeta.data.sheets ?? [];

  const results: BizTutorTabResult[] = [];
  for (const tabName of sourceTabs) {
    try {
      const result = await processSourceTab(
        sheets,
        sourceSpreadsheetId,
        allSourceTabs,
        tabName,
        emailToTutor,
        phoneToTutor
      );
      results.push(result);
    } catch (err) {
      results.push({
        tabName,
        totalRows: 0,
        matchedByEmail: 0,
        matchedByPhone: 0,
        unmatched: 0,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  return {
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${sourceSpreadsheetId}/edit`,
    tabs: results,
  };
}
