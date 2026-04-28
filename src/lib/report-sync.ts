import { getSheetsClient } from "./google-auth";
import {
  columnIndexToLetter,
  findEmailColumnIndex,
  normalizeEmail,
} from "./email-utils";
import { withRetry } from "./retry";
import { sheets_v4 } from "googleapis";

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

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

interface FieldDef {
  key: "esito" | "motivazione" | "noteVenditrice" | "dataRecall";
  header: string;
  aliases: string[];
}

const FIELDS: FieldDef[] = [
  {
    key: "esito",
    header: "ESITO",
    aliases: ["esito"],
  },
  {
    key: "motivazione",
    header: "MOTIVAZIONE",
    aliases: ["motivazione"],
  },
  {
    key: "noteVenditrice",
    header: "NOTE VENDITRICE",
    aliases: ["note venditrice"],
  },
  {
    key: "dataRecall",
    header: "DATA RECALL",
    aliases: ["data recall"],
  },
];

type FieldValues = Record<FieldDef["key"], string>;

const EMPTY_FIELDS: FieldValues = {
  esito: "",
  motivazione: "",
  noteVenditrice: "",
  dataRecall: "",
};

export interface ReportTabResult {
  tabName: string;
  totalRows: number;
  matchedByEmail: number;
  matchedByPhone: number;
  unmatched: number;
  error?: string;
}

export interface ReportDestinationResult {
  spreadsheetUrl: string;
  tabs: ReportTabResult[];
  error?: string;
}

export interface ReportSyncResult {
  destinations: ReportDestinationResult[];
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
  if (digits.length > 10) set.add(digits.slice(-10));
  return Array.from(set);
}

async function buildLookupMap(
  sheets: sheets_v4.Sheets,
  sourceSpreadsheetId: string,
  sourceTabs: string[]
): Promise<{
  emailToFields: Map<string, FieldValues>;
  phoneToFields: Map<string, FieldValues>;
}> {
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: sourceSpreadsheetId })
  );
  const allTabs = meta.data.sheets ?? [];

  const emailToFields = new Map<string, FieldValues>();
  const phoneToFields = new Map<string, FieldValues>();

  for (const tabName of sourceTabs) {
    const tab = allTabs.find((t) => t.properties?.title === tabName);
    if (!tab) {
      throw new Error(`Source tab "${tabName}" not found`);
    }
    const safeTab = `'${tabName.replace(/'/g, "''")}'`;

    const headerRes = await withRetry(() =>
      sheets.spreadsheets.values.get({
        spreadsheetId: sourceSpreadsheetId,
        range: `${safeTab}!1:1`,
      })
    );
    const headers = (headerRes.data.values?.[0] ?? []) as string[];

    const fieldIdx: Record<FieldDef["key"], number> = {
      esito: -1,
      motivazione: -1,
      noteVenditrice: -1,
      dataRecall: -1,
    };
    for (const f of FIELDS) {
      fieldIdx[f.key] = findHeaderIndex(headers, f.aliases);
    }
    const presentFields = FIELDS.filter((f) => fieldIdx[f.key] !== -1);
    if (presentFields.length === 0) {
      throw new Error(
        `Source tab "${tabName}" has none of the required columns (ESITO, MOTIVAZIONE, NOTE VENDITRICE, DATA RECALL). Headers: ${headers.join(", ")}`
      );
    }

    const emailIdx = findEmailColumnIndex(headers);
    const phoneIdx = findHeaderIndex(headers, PHONE_ALIASES);
    if (emailIdx === -1 && phoneIdx === -1) {
      throw new Error(
        `Source tab "${tabName}" has neither Email nor Phone column to match on`
      );
    }

    const ranges: string[] = [];
    const rangeMap: {
      email?: number;
      phone?: number;
      fields: Partial<Record<FieldDef["key"], number>>;
    } = { fields: {} };

    if (emailIdx !== -1) {
      rangeMap.email = ranges.length;
      const letter = columnIndexToLetter(emailIdx);
      ranges.push(`${safeTab}!${letter}2:${letter}`);
    }
    if (phoneIdx !== -1) {
      rangeMap.phone = ranges.length;
      const letter = columnIndexToLetter(phoneIdx);
      ranges.push(`${safeTab}!${letter}2:${letter}`);
    }
    for (const f of presentFields) {
      rangeMap.fields[f.key] = ranges.length;
      const letter = columnIndexToLetter(fieldIdx[f.key]);
      ranges.push(`${safeTab}!${letter}2:${letter}`);
    }

    const dataRes = await withRetry(() =>
      sheets.spreadsheets.values.batchGet({
        spreadsheetId: sourceSpreadsheetId,
        ranges,
      })
    );
    const valueRanges = dataRes.data.valueRanges ?? [];

    const emailRows =
      rangeMap.email !== undefined
        ? ((valueRanges[rangeMap.email]?.values ?? []) as string[][])
        : [];
    const phoneRows =
      rangeMap.phone !== undefined
        ? ((valueRanges[rangeMap.phone]?.values ?? []) as string[][])
        : [];

    const fieldRows: Partial<Record<FieldDef["key"], string[][]>> = {};
    for (const f of presentFields) {
      const idx = rangeMap.fields[f.key]!;
      fieldRows[f.key] = (valueRanges[idx]?.values ?? []) as string[][];
    }

    let rowCount = Math.max(emailRows.length, phoneRows.length);
    for (const f of presentFields) {
      rowCount = Math.max(rowCount, fieldRows[f.key]?.length ?? 0);
    }

    for (let i = 0; i < rowCount; i++) {
      const values: FieldValues = { ...EMPTY_FIELDS };
      let hasAnyField = false;
      for (const f of presentFields) {
        const raw = fieldRows[f.key]?.[i]?.[0];
        if (raw !== undefined && raw !== null && String(raw).trim() !== "") {
          values[f.key] = String(raw);
          hasAnyField = true;
        }
      }
      if (!hasAnyField) continue;

      const emailRaw = emailRows[i]?.[0];
      if (emailRaw) {
        const email = normalizeEmail(String(emailRaw));
        if (email && !emailToFields.has(email)) {
          emailToFields.set(email, values);
        }
      }

      const phoneRaw = phoneRows[i]?.[0];
      if (phoneRaw) {
        for (const key of phoneKeys(String(phoneRaw))) {
          if (!phoneToFields.has(key)) {
            phoneToFields.set(key, values);
          }
        }
      }
    }
  }

  return { emailToFields, phoneToFields };
}

async function processDestinationTab(
  sheets: sheets_v4.Sheets,
  destinationSpreadsheetId: string,
  allDestinationTabs: sheets_v4.Schema$Sheet[],
  tabName: string,
  emailToFields: Map<string, FieldValues>,
  phoneToFields: Map<string, FieldValues>
): Promise<ReportTabResult> {
  const tab = allDestinationTabs.find((t) => t.properties?.title === tabName);
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
      spreadsheetId: destinationSpreadsheetId,
      range: `${safeTab}!1:1`,
    })
  );
  const headers = (headerRes.data.values?.[0] ?? []) as string[];

  const emailIdx = findEmailColumnIndex(headers);
  const phoneIdx = findHeaderIndex(headers, PHONE_ALIASES);
  if (emailIdx === -1 && phoneIdx === -1) {
    throw new Error(
      `Destination tab "${tabName}" has neither Email nor Phone column to look up`
    );
  }

  // Determine column index for each field — reuse existing or append
  let lastFilledIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]?.toString().trim()) lastFilledIdx = i;
  }
  let nextAppend = lastFilledIdx + 1;

  const fieldCol: Record<FieldDef["key"], { index: number; needsHeader: boolean }> = {
    esito: { index: -1, needsHeader: false },
    motivazione: { index: -1, needsHeader: false },
    noteVenditrice: { index: -1, needsHeader: false },
    dataRecall: { index: -1, needsHeader: false },
  };
  for (const f of FIELDS) {
    const existing = findHeaderIndex(headers, f.aliases);
    if (existing !== -1) {
      fieldCol[f.key] = { index: existing, needsHeader: false };
    } else {
      fieldCol[f.key] = { index: nextAppend, needsHeader: true };
      nextAppend++;
    }
  }

  const ranges: string[] = [];
  const rangeMap: {
    email?: number;
    phone?: number;
    existing: Record<FieldDef["key"], number>;
  } = {
    existing: { esito: -1, motivazione: -1, noteVenditrice: -1, dataRecall: -1 },
  };

  if (emailIdx !== -1) {
    rangeMap.email = ranges.length;
    const letter = columnIndexToLetter(emailIdx);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }
  if (phoneIdx !== -1) {
    rangeMap.phone = ranges.length;
    const letter = columnIndexToLetter(phoneIdx);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }
  for (const f of FIELDS) {
    if (!fieldCol[f.key].needsHeader) {
      rangeMap.existing[f.key] = ranges.length;
      const letter = columnIndexToLetter(fieldCol[f.key].index);
      ranges.push(`${safeTab}!${letter}2:${letter}`);
    }
  }

  const dataRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId: destinationSpreadsheetId,
      ranges,
    })
  );
  const valueRanges = dataRes.data.valueRanges ?? [];

  const emailRows =
    rangeMap.email !== undefined
      ? ((valueRanges[rangeMap.email]?.values ?? []) as string[][])
      : [];
  const phoneRows =
    rangeMap.phone !== undefined
      ? ((valueRanges[rangeMap.phone]?.values ?? []) as string[][])
      : [];

  const existingRows: Record<FieldDef["key"], string[][]> = {
    esito: [],
    motivazione: [],
    noteVenditrice: [],
    dataRecall: [],
  };
  for (const f of FIELDS) {
    const idx = rangeMap.existing[f.key];
    if (idx !== -1) {
      existingRows[f.key] = (valueRanges[idx]?.values ?? []) as string[][];
    }
  }

  let rowCount = Math.max(emailRows.length, phoneRows.length);
  for (const f of FIELDS) {
    rowCount = Math.max(rowCount, existingRows[f.key].length);
  }

  let matchedByEmail = 0;
  let matchedByPhone = 0;
  let unmatched = 0;
  let totalRows = 0;

  const outputs: Record<FieldDef["key"], string[]> = {
    esito: new Array(rowCount).fill(""),
    motivazione: new Array(rowCount).fill(""),
    noteVenditrice: new Array(rowCount).fill(""),
    dataRecall: new Array(rowCount).fill(""),
  };

  for (let i = 0; i < rowCount; i++) {
    const emailRaw = emailRows[i]?.[0];
    const phoneRaw = phoneRows[i]?.[0];
    const hasIdentifier = !!emailRaw || !!phoneRaw;
    if (hasIdentifier) totalRows++;

    let resolved: FieldValues | null = null;
    let matchKind: "email" | "phone" | null = null;

    if (emailRaw) {
      const email = normalizeEmail(String(emailRaw));
      if (email && emailToFields.has(email)) {
        resolved = emailToFields.get(email)!;
        matchKind = "email";
      }
    }
    if (resolved === null && phoneRaw) {
      for (const key of phoneKeys(String(phoneRaw))) {
        if (phoneToFields.has(key)) {
          resolved = phoneToFields.get(key)!;
          matchKind = "phone";
          break;
        }
      }
    }

    if (resolved !== null) {
      if (matchKind === "email") matchedByEmail++;
      else matchedByPhone++;
      for (const f of FIELDS) {
        outputs[f.key][i] = resolved[f.key] ?? "";
      }
    } else {
      // Preserve existing values when no match
      for (const f of FIELDS) {
        const existing = existingRows[f.key][i]?.[0];
        outputs[f.key][i] = existing ? String(existing) : "";
      }
      if (hasIdentifier) unmatched++;
    }
  }

  const requests: sheets_v4.Schema$Request[] = [];

  for (const f of FIELDS) {
    if (fieldCol[f.key].needsHeader) {
      requests.push({
        updateCells: {
          rows: [
            {
              values: [
                {
                  userEnteredValue: { stringValue: f.header },
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
          start: { sheetId, rowIndex: 0, columnIndex: fieldCol[f.key].index },
        },
      });
    }
  }

  if (rowCount > 0) {
    for (const f of FIELDS) {
      requests.push({
        updateCells: {
          rows: outputs[f.key].map((value) => ({
            values: [{ userEnteredValue: { stringValue: value } }],
          })),
          fields: "userEnteredValue",
          start: { sheetId, rowIndex: 1, columnIndex: fieldCol[f.key].index },
        },
      });
    }
  }

  if (requests.length > 0) {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId: destinationSpreadsheetId,
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

export async function runReportSync(
  refreshToken: string,
  sourceSpreadsheetId: string,
  sourceTabs: string[],
  destinations: Array<{ spreadsheetId: string; tabs: string[] }>
): Promise<ReportSyncResult> {
  const sheets = getSheetsClient(refreshToken);

  const { emailToFields, phoneToFields } = await buildLookupMap(
    sheets,
    sourceSpreadsheetId,
    sourceTabs
  );

  const destResults: ReportDestinationResult[] = [];

  for (const dest of destinations) {
    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${dest.spreadsheetId}/edit`;
    try {
      const meta = await withRetry(() =>
        sheets.spreadsheets.get({ spreadsheetId: dest.spreadsheetId })
      );
      const allTabs = meta.data.sheets ?? [];

      const tabResults: ReportTabResult[] = [];
      for (const tabName of dest.tabs) {
        try {
          const r = await processDestinationTab(
            sheets,
            dest.spreadsheetId,
            allTabs,
            tabName,
            emailToFields,
            phoneToFields
          );
          tabResults.push(r);
        } catch (err) {
          tabResults.push({
            tabName,
            totalRows: 0,
            matchedByEmail: 0,
            matchedByPhone: 0,
            unmatched: 0,
            error: err instanceof Error ? err.message : String(err),
          });
        }
      }
      destResults.push({ spreadsheetUrl, tabs: tabResults });
    } catch (err) {
      destResults.push({
        spreadsheetUrl,
        tabs: [],
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  return { destinations: destResults };
}
