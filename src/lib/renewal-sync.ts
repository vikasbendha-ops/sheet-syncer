import { getSheetsClient } from "./google-auth";
import {
  columnIndexToLetter,
  findEmailColumnIndex,
  normalizeEmail,
} from "./email-utils";
import { withRetry } from "./retry";
import { sheets_v4 } from "googleapis";

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };
const WHITE_BG = { red: 1.0, green: 1.0, blue: 1.0 };

// Renewal-proximity row tiers (background + text color)
const LIGHT_GREEN_BG = { red: 0.78, green: 0.92, blue: 0.78 }; // 15-30 days
const LIGHT_YELLOW_BG = { red: 1.0, green: 0.95, blue: 0.7 }; // 5-14 days
const LIGHT_RED_BG = { red: 1.0, green: 0.82, blue: 0.82 }; // 0-4 days
const DARK_RED_BG = { red: 0.78, green: 0.1, blue: 0.1 }; // past
const BLACK_TEXT = { red: 0, green: 0, blue: 0 };
const WHITE_TEXT = { red: 1, green: 1, blue: 1 };

// The tier classification is still computed per row so we can count
// pastRenewals in the result summary, but the actual coloring is now done
// by native Google Sheets conditional-format rules installed below — not by
// painting cell backgrounds directly. That means the highlighting refreshes
// automatically every day (and on every cell edit) without re-running sync.
type RenewalTier = "past" | "imminent" | "soon" | "later" | "none";

// Google Sheets serial date epoch: 1899-12-30 (anchored to absorb the
// historical Lotus 1-2-3 leap-year-1900 bug). Using UTC-based math so the
// server's timezone never shifts the result by ±1 day.
const SHEETS_DATE_EPOCH_UTC_MS = Date.UTC(1899, 11, 30);

function toSheetsDateSerial(d: Date): number {
  const y = d.getFullYear();
  const m = d.getMonth();
  const day = d.getDate();
  return Math.floor(
    (Date.UTC(y, m, day) - SHEETS_DATE_EPOCH_UTC_MS) / 86400000
  );
}

const FIELDS = [
  "phone",
  "courseName",
  "startDate",
  "renewalDate",
  "setterAssigned",
  "subscriptionType",
] as const;
type FieldKey = (typeof FIELDS)[number];

// Fields whose value is COMPUTED from other fields (not pulled from the
// lookup spreadsheet). buildLookupMap skips these when scanning for columns,
// and the post-processing step in buildLookupMap fills them.
const COMPUTED_FIELDS: Set<FieldKey> = new Set(["subscriptionType"]);

const CANONICAL_HEADERS: Record<FieldKey, string> = {
  phone: "Phone",
  courseName: "Course name",
  startDate: "Start date",
  renewalDate: "Renewal Date",
  setterAssigned: "Setter assigned",
  subscriptionType: "TYPE OF SUBSCRIPTION",
};

// Dropdown options for the TYPE OF SUBSCRIPTION column. Installed as a Google
// Sheets data validation rule (ONE_OF_LIST) on the column each sync run.
const SUBSCRIPTION_OPTIONS = [
  "BAL",
  "BAC",
  "ELITE",
  "GOLD",
  "NMM",
] as const;

/**
 * Map a Course name string to one of the SUBSCRIPTION_OPTIONS, or "" if
 * nothing matches (cell left empty / user can pick manually). Match is
 * case-insensitive substring.
 *
 * NMM has no auto-rule — it lives in the dropdown but only via manual pick.
 *
 * Order matters only for clarity; the rules are mutually exclusive in
 * practice (a Course name in the real data is one of the four families).
 */
function classifySubscription(courseName: string): string {
  if (!courseName) return "";
  const c = courseName.toLowerCase();
  if (c.includes("biz academy club")) return "BAC";
  if (c.includes("biz academy light")) return "BAL";
  if (c.includes("biz academy elite") || c.includes("elite")) return "ELITE";
  if (c.includes("biz academy gold") || c.includes("gold")) return "GOLD";
  return "";
}

// Accept common variants (English + Italian) when locating columns.
// Aliases are matched against the header after lowercasing, trimming, and
// collapsing internal whitespace — so casing and extra spaces don't matter.
const HEADER_ALIASES: Record<FieldKey, string[]> = {
  phone: ["phone", "phone number", "telefono", "mobile", "cellulare"],
  courseName: [
    "course name",
    "course",
    "corso",
    "nome corso",
    "subscription / course name", // exact literal header users have in their lookup sheets
    "subscription/course name", // no-space variant — defensive
    "subscription",
  ],
  startDate: ["start date", "start", "data inizio", "inizio", "data di inizio"],
  renewalDate: [
    "renewal date",
    "renewal",
    "data rinnovo",
    "rinnovo",
    "data di rinnovo",
    "renewal / expiry date", // literal header users have in lookup sheets
    "renewal/expiry date", // no-space variant
    "expiry date",
    "expiry",
  ],
  setterAssigned: ["setter assigned", "setter", "assigned setter", "assegnato"],
  subscriptionType: [
    "type of subscription",
    "subscription type",
    "tipo di abbonamento",
    "tipo abbonamento",
    "abbonamento",
  ],
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

/**
 * Classify a renewal date into one of five proximity tiers relative to today.
 *  past     — renewal already happened (date < today)
 *  imminent — renewal happens today or in the next 4 days (0-4 days)
 *  soon     — renewal in 5-14 days
 *  later    — renewal in 15-30 days
 *  none     — no parseable date, or more than 30 days away
 */
function computeRenewalTier(raw: string): RenewalTier {
  const d = parseFlexibleDate(raw);
  if (!d) return "none";
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(d);
  target.setHours(0, 0, 0, 0);
  const diffDays = Math.floor(
    (target.getTime() - today.getTime()) / 86400000
  );
  if (diffDays < 0) return "past";
  if (diffDays <= 4) return "imminent";
  if (diffDays <= 14) return "soon";
  if (diffDays <= 30) return "later";
  return "none";
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
      if (COMPUTED_FIELDS.has(f)) continue; // not pulled from lookup
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
        subscriptionType: "",
      };
      for (const f of FIELDS) {
        if (COMPUTED_FIELDS.has(f)) continue; // computed after lookup read
        const cell = fieldRows[f]?.[i]?.[0];
        if (cell !== undefined && cell !== null) {
          data[f] = String(cell).trim();
        }
      }
      // Derived field: classify subscription type from courseName
      data.subscriptionType = classifySubscription(data.courseName);
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
    tier: RenewalTier;
  }
  const outcomes: RowOutcome[] = [];
  let matched = 0;
  let unmatched = 0;
  let pastRenewals = 0;
  let totalRows = 0;

  for (let i = 0; i < emailRows.length; i++) {
    const raw = emailRows[i]?.[0];
    if (!raw) {
      outcomes.push({ email: null, matched: false, data: null, tier: "none" });
      continue;
    }
    const email = normalizeEmail(String(raw));
    if (!email) {
      outcomes.push({ email: null, matched: false, data: null, tier: "none" });
      continue;
    }
    totalRows++;
    const data = lookupMap.get(email) ?? null;
    if (data) {
      matched++;
      const tier = computeRenewalTier(data.renewalDate);
      if (tier === "past") pastRenewals++;
      outcomes.push({ email, matched: true, data, tier });
    } else {
      unmatched++;
      outcomes.push({ email, matched: false, data: null, tier: "none" });
    }
  }

  const requests: sheets_v4.Schema$Request[] = [];

  // Grid bookkeeping. `rowEndCol` is the exclusive rightmost column index
  // we will touch (covers existing data + any newly appended target cols).
  // `thisTabInfo` lets us inspect the sheet's grid + existing conditional
  // formats. Sheets enforces that updateCells / repeatCell stay within the
  // grid bounds, so if our writes extend past the current columnCount we
  // must grow the grid FIRST in the same batch.
  const rowEndCol = Math.max(nextAppendIdx, lastFilledIdx + 1);
  const thisTabInfo = allSourceTabs.find(
    (t) => t.properties?.sheetId === sheetId
  );
  const currentColumnCount =
    thisTabInfo?.properties?.gridProperties?.columnCount ?? 0;
  if (rowEndCol > currentColumnCount) {
    requests.push({
      appendDimension: {
        sheetId,
        dimension: "COLUMNS",
        length: rowEndCol - currentColumnCount,
      },
    });
  }

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

  // 2. Per-field column writes.
  //    Renewal Date is written as a REAL date (numberValue + DATE number
  //    format) when parseable, so the native conditional-formatting rules
  //    below can compare it directly with TODAY(). Unparseable values fall
  //    back to a string write. All other fields are written as strings.
  if (outcomes.length > 0) {
    for (const f of FIELDS) {
      const isRenewalDate = f === "renewalDate";
      requests.push({
        updateCells: {
          rows: outcomes.map((o) => {
            const raw = o.data ? o.data[f] : "";

            if (isRenewalDate && raw) {
              const parsed = parseFlexibleDate(raw);
              if (parsed) {
                return {
                  values: [
                    {
                      userEnteredValue: {
                        numberValue: toSheetsDateSerial(parsed),
                      },
                      userEnteredFormat: {
                        numberFormat: { type: "DATE", pattern: "dd/mm/yyyy" },
                      },
                    },
                  ],
                };
              }
            }

            return {
              values: [
                {
                  userEnteredValue: { stringValue: raw },
                  ...(isRenewalDate
                    ? {
                        // Clear date format so a previously formatted cell
                        // doesn't render "" as a number.
                        userEnteredFormat: { numberFormat: { type: "TEXT" } },
                      }
                    : {}),
                },
              ],
            };
          }),
          fields: isRenewalDate
            ? "userEnteredValue,userEnteredFormat.numberFormat"
            : "userEnteredValue",
          start: { sheetId, rowIndex: 1, columnIndex: targetCol[f] },
        },
      });
    }
  }

  // 3. Reset background to white + text to black across the data range.
  //    Native conditional rules can only ADD formatting — they can't undo
  //    a manual cell color from a previous static-paint run. Resetting
  //    once each sync run clears any stale red/yellow/green and lets the
  //    conditional rules below take full control of cell appearance.
  if (outcomes.length > 0 && rowEndCol > 0) {
    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: 1,
          endRowIndex: outcomes.length + 1,
          startColumnIndex: 0,
          endColumnIndex: rowEndCol,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: WHITE_BG,
            textFormat: { foregroundColor: BLACK_TEXT },
          },
        },
        fields:
          "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.foregroundColor",
      },
    });
  }

  // 4. Wipe any existing conditional-format rules on this sheet, then
  //    install the four tier rules. Rules persist in the spreadsheet and
  //    Google re-evaluates them on every open + every cell edit — so the
  //    highlighting stays accurate daily without ever re-running the sync.
  const existingRules = thisTabInfo?.conditionalFormats ?? [];
  // Delete from highest index down so each delete leaves the remaining
  // indices stable.
  for (let i = existingRules.length - 1; i >= 0; i--) {
    requests.push({
      deleteConditionalFormatRule: { sheetId, index: i },
    });
  }

  // Cover every row in the sheet's actual grid so newly added rows
  // automatically pick up the rules. Falls back to a generous range if
  // gridProperties is unavailable.
  const sheetRowCount =
    thisTabInfo?.properties?.gridProperties?.rowCount ??
    Math.max(outcomes.length + 1, 1000);

  const renewalColLetter = columnIndexToLetter(targetCol.renewalDate);
  // The conditional-format range starts at row index 1 (zero-based) = the
  // first data row, displayed as row 2 in the Sheets UI. Sheets evaluates
  // the formula with the column anchored ($-prefixed) and the row relative
  // to each cell in the range, so $X2 auto-shifts to $X3, $X4, … per row.
  const ref = `$${renewalColLetter}2`;

  const conditionalRange: sheets_v4.Schema$GridRange = {
    sheetId,
    startRowIndex: 1,
    endRowIndex: sheetRowCount,
    startColumnIndex: 0,
    endColumnIndex: rowEndCol,
  };

  const tierRules: Array<{
    formula: string;
    bg: { red: number; green: number; blue: number };
    fg: { red: number; green: number; blue: number };
  }> = [
    {
      // Past (must come first — covers everything strictly before today)
      formula: `=AND(ISNUMBER(${ref}), ${ref}<TODAY())`,
      bg: DARK_RED_BG,
      fg: WHITE_TEXT,
    },
    {
      // Imminent: today + next 4 days
      formula: `=AND(ISNUMBER(${ref}), ${ref}>=TODAY(), ${ref}<=TODAY()+4)`,
      bg: LIGHT_RED_BG,
      fg: BLACK_TEXT,
    },
    {
      // Soon: 5-14 days
      formula: `=AND(ISNUMBER(${ref}), ${ref}>=TODAY()+5, ${ref}<=TODAY()+14)`,
      bg: LIGHT_YELLOW_BG,
      fg: BLACK_TEXT,
    },
    {
      // Later: 15-30 days
      formula: `=AND(ISNUMBER(${ref}), ${ref}>=TODAY()+15, ${ref}<=TODAY()+30)`,
      bg: LIGHT_GREEN_BG,
      fg: BLACK_TEXT,
    },
  ];

  tierRules.forEach((r, idx) => {
    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [conditionalRange],
          booleanRule: {
            condition: {
              type: "CUSTOM_FORMULA",
              values: [{ userEnteredValue: r.formula }],
            },
            format: {
              backgroundColor: r.bg,
              textFormat: { foregroundColor: r.fg },
            },
          },
        },
        index: idx,
      },
    });
  });

  // 5. Install the TYPE OF SUBSCRIPTION dropdown (data validation) on the
  //    subscriptionType column. Re-run safe: setDataValidation overwrites
  //    any prior rule on the same range without piling them up. Per-value
  //    chip colors live on the cells themselves once the user sets them in
  //    the Sheets UI — that styling is NOT overwritten by setDataValidation.
  requests.push({
    setDataValidation: {
      range: {
        sheetId,
        startRowIndex: 1, // skip header
        endRowIndex: sheetRowCount,
        startColumnIndex: targetCol.subscriptionType,
        endColumnIndex: targetCol.subscriptionType + 1,
      },
      rule: {
        condition: {
          type: "ONE_OF_LIST",
          values: SUBSCRIPTION_OPTIONS.map((v) => ({
            userEnteredValue: v,
          })),
        },
        strict: true,
        showCustomUi: true,
      },
    },
  });

  // Row-paint adds N requests for an N-row tab; bumping chunk size keeps
  // total round-trips reasonable on large sheets without risking rate limits.
  const CHUNK_SIZE = 25;
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
