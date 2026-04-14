import { readEmailsFromSheet } from "./sheets-reader";
import { getSheetsClient } from "./google-auth";
import { sheets_v4 } from "googleapis";

const MARKER = "=== DOMAIN DISTRIBUTION ===";
const CHART_TITLE_PREFIX = "Domain Distribution";

export interface DomainCount {
  domain: string;
  count: number;
  percent: number;
}

export interface TabAnalysis {
  tabName: string;
  sheetId: number;
  totalEmails: number;
  uniqueEmails: number;
  domains: DomainCount[];
  error?: string;
}

export interface DomainAnalysisResult {
  spreadsheetId: string;
  tabs: TabAnalysis[];
}

function extractDomain(email: string): string | null {
  const at = email.lastIndexOf("@");
  if (at === -1 || at === email.length - 1) return null;
  return email.slice(at + 1).toLowerCase();
}

export async function analyzeDomainsForTabs(
  refreshToken: string,
  spreadsheetId: string,
  tabs: string[],
  emailColumn: string = "auto"
): Promise<DomainAnalysisResult> {
  const sheetsClient = getSheetsClient(refreshToken);
  let cachedTabs: sheets_v4.Schema$Sheet[] | undefined;
  try {
    const res = await sheetsClient.spreadsheets.get({ spreadsheetId });
    cachedTabs = res.data.sheets ?? [];
  } catch {
    // continue without cache
  }

  const results: TabAnalysis[] = await Promise.all(
    tabs.map(async (tabName): Promise<TabAnalysis> => {
      try {
        const result = await readEmailsFromSheet(
          refreshToken,
          spreadsheetId,
          tabName,
          emailColumn,
          cachedTabs
        );

        const domainCounts = new Map<string, number>();
        let totalEmails = 0;

        for (const [email, rowNumbers] of result.emails.entries()) {
          const domain = extractDomain(email);
          if (!domain) continue;
          const occurrences = rowNumbers.length;
          totalEmails += occurrences;
          domainCounts.set(domain, (domainCounts.get(domain) ?? 0) + occurrences);
        }

        const domains: DomainCount[] = Array.from(domainCounts.entries())
          .map(([domain, count]) => ({
            domain,
            count,
            percent: totalEmails > 0 ? (count / totalEmails) * 100 : 0,
          }))
          .sort((a, b) => b.count - a.count);

        return {
          tabName: result.resolvedTabName,
          sheetId: result.sheetId,
          totalEmails,
          uniqueEmails: result.emails.size,
          domains,
        };
      } catch (err) {
        return {
          tabName,
          sheetId: 0,
          totalEmails: 0,
          uniqueEmails: 0,
          domains: [],
          error: err instanceof Error ? err.message : String(err),
        };
      }
    })
  );

  return { spreadsheetId, tabs: results };
}

/**
 * Writes the analysis (table + chart) to the bottom of each analyzed tab.
 * Idempotent: detects prior runs (by marker text + chart title) and replaces them.
 */
export async function writeDomainAnalysesToSheets(
  refreshToken: string,
  spreadsheetId: string,
  analyses: TabAnalysis[]
): Promise<{ writtenTabs: string[]; errors: string[] }> {
  const writtenTabs: string[] = [];
  const errors: string[] = [];

  for (const analysis of analyses) {
    if (analysis.error || analysis.domains.length === 0) continue;
    try {
      await writeOneTabAnalysis(refreshToken, spreadsheetId, analysis);
      writtenTabs.push(analysis.tabName);
    } catch (err) {
      errors.push(
        `Failed to write analysis to "${analysis.tabName}": ${err instanceof Error ? err.message : String(err)}`
      );
    }
  }

  return { writtenTabs, errors };
}

async function writeOneTabAnalysis(
  refreshToken: string,
  spreadsheetId: string,
  analysis: TabAnalysis
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);
  const { tabName, sheetId } = analysis;
  const safeTab = `'${tabName.replace(/'/g, "''")}'`;

  // 1. Find existing marker in column A → clear from there down
  const colARes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${safeTab}!A:A`,
  });
  const colA = colARes.data.values ?? [];

  let markerRow = -1;
  for (let i = 0; i < colA.length; i++) {
    const cell = colA[i]?.[0];
    if (cell && String(cell).startsWith(MARKER)) {
      markerRow = i + 1; // 1-based
      break;
    }
  }

  let lastDataRow: number;
  if (markerRow > 0) {
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${safeTab}!A${markerRow}:Z${colA.length + 100}`,
    });
    lastDataRow = markerRow - 1;
    while (lastDataRow > 0 && !colA[lastDataRow - 1]?.[0]) {
      lastDataRow--;
    }
  } else {
    lastDataRow = colA.length;
    while (lastDataRow > 0 && !colA[lastDataRow - 1]?.[0]) {
      lastDataRow--;
    }
  }

  // 2. Delete prior charts with our marker title (for this tab)
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets(properties(sheetId),charts(chartId,spec(title)))",
  });
  const sheetMeta = meta.data.sheets?.find(
    (s) => s.properties?.sheetId === sheetId
  );
  const chartsToDelete = (sheetMeta?.charts ?? [])
    .filter((c) =>
      c.spec?.title?.startsWith(`${CHART_TITLE_PREFIX} - ${tabName}`)
    )
    .map((c) => c.chartId)
    .filter((id): id is number => typeof id === "number");

  // 3. Compute target row positions (1-based)
  const startRow = lastDataRow + 2; // one empty row separator
  const markerWriteRow = startRow;
  const tsRow = startRow + 1;
  const headerRow = startRow + 2;
  const dataStartRow = headerRow + 1;
  const dataEndRow = dataStartRow + analysis.domains.length - 1;

  // 4. Build the values to write
  const rows: (string | number)[][] = [
    [`${MARKER} (${tabName})`],
    [`Generated ${new Date().toISOString()}`],
    ["Domain", "Count", "Percent"],
    ...analysis.domains.map((d) => [d.domain, d.count, d.percent / 100]),
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${safeTab}!A${markerWriteRow}:C${dataEndRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: rows },
  });

  // 5. Format + chart via batchUpdate
  const requests: sheets_v4.Schema$Request[] = [];

  // delete prior charts first
  for (const id of chartsToDelete) {
    requests.push({ deleteEmbeddedObject: { objectId: id } });
  }

  // bold the marker + header rows
  requests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: markerWriteRow - 1,
        endRowIndex: markerWriteRow,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      cell: {
        userEnteredFormat: {
          textFormat: { bold: true, fontSize: 11 },
          backgroundColor: { red: 0.93, green: 0.93, blue: 0.96 },
        },
      },
      fields:
        "userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.fontSize,userEnteredFormat.backgroundColor",
    },
  });
  requests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: tsRow - 1,
        endRowIndex: tsRow,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      cell: {
        userEnteredFormat: {
          textFormat: { italic: true, foregroundColor: { red: 0.4, green: 0.4, blue: 0.4 } },
        },
      },
      fields: "userEnteredFormat.textFormat.italic,userEnteredFormat.textFormat.foregroundColor",
    },
  });
  requests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: headerRow - 1,
        endRowIndex: headerRow,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
      cell: {
        userEnteredFormat: { textFormat: { bold: true } },
      },
      fields: "userEnteredFormat.textFormat.bold",
    },
  });
  // format percent column as percentage
  requests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: dataStartRow - 1,
        endRowIndex: dataEndRow,
        startColumnIndex: 2,
        endColumnIndex: 3,
      },
      cell: {
        userEnteredFormat: {
          numberFormat: { type: "PERCENT", pattern: "0.00%" },
        },
      },
      fields: "userEnteredFormat.numberFormat",
    },
  });

  // add chart
  requests.push({
    addChart: {
      chart: {
        spec: {
          title: `${CHART_TITLE_PREFIX} - ${tabName}`,
          basicChart: {
            chartType: "BAR",
            legendPosition: "NO_LEGEND",
            axis: [
              { position: "BOTTOM_AXIS", title: "Count" },
              { position: "LEFT_AXIS", title: "Domain" },
            ],
            domains: [
              {
                domain: {
                  sourceRange: {
                    sources: [
                      {
                        sheetId,
                        startRowIndex: headerRow - 1,
                        endRowIndex: dataEndRow,
                        startColumnIndex: 0,
                        endColumnIndex: 1,
                      },
                    ],
                  },
                },
              },
            ],
            series: [
              {
                series: {
                  sourceRange: {
                    sources: [
                      {
                        sheetId,
                        startRowIndex: headerRow - 1,
                        endRowIndex: dataEndRow,
                        startColumnIndex: 1,
                        endColumnIndex: 2,
                      },
                    ],
                  },
                },
                targetAxis: "BOTTOM_AXIS",
              },
            ],
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: {
              sheetId,
              rowIndex: markerWriteRow - 1,
              columnIndex: 4, // column E
            },
            offsetXPixels: 0,
            offsetYPixels: 0,
            widthPixels: 600,
            heightPixels: Math.max(300, analysis.domains.length * 22 + 120),
          },
        },
      },
    },
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests },
  });
}
