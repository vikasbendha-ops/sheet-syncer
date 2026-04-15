import { readEmailsFromSheet } from "./sheets-reader";
import { getSheetsClient } from "./google-auth";
import { sheets_v4 } from "googleapis";

const TAB_PREFIX = "Domains - ";
const MAX_TAB_NAME_LENGTH = 100;

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

function buildAnalysisTabName(sourceTabName: string): string {
  const full = `${TAB_PREFIX}${sourceTabName}`;
  if (full.length <= MAX_TAB_NAME_LENGTH) return full;
  // Truncate the source name portion to fit within Sheets' 100-char limit
  return full.slice(0, MAX_TAB_NAME_LENGTH);
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
 * Writes one analysis tab per analyzed source tab.
 * Idempotent: if the analysis tab already exists, it is cleared and rewritten.
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
      const targetName = await writeOneAnalysisTab(
        refreshToken,
        spreadsheetId,
        analysis
      );
      writtenTabs.push(targetName);
    } catch (err) {
      errors.push(
        `Failed to write analysis for "${analysis.tabName}": ${err instanceof Error ? err.message : String(err)}`
      );
    }
  }

  return { writtenTabs, errors };
}

async function writeOneAnalysisTab(
  refreshToken: string,
  spreadsheetId: string,
  analysis: TabAnalysis
): Promise<string> {
  const sheets = getSheetsClient(refreshToken);
  const targetName = buildAnalysisTabName(analysis.tabName);

  // 1. Find or create the target tab
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets(properties(sheetId,title),charts(chartId))",
  });
  const allSheets = meta.data.sheets ?? [];
  const existing = allSheets.find((s) => s.properties?.title === targetName);

  let targetSheetId: number;
  if (existing && existing.properties?.sheetId !== undefined && existing.properties.sheetId !== null) {
    targetSheetId = existing.properties.sheetId;

    // Clear all values
    const safeName = `'${targetName.replace(/'/g, "''")}'`;
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${safeName}!A:ZZ`,
    });

    // Delete all charts in this tab
    const chartIds = (existing.charts ?? [])
      .map((c) => c.chartId)
      .filter((id): id is number => typeof id === "number");
    if (chartIds.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: chartIds.map((id) => ({
            deleteEmbeddedObject: { objectId: id },
          })),
        },
      });
    }
  } else {
    // Create the new tab at the end
    const addRes = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: targetName },
            },
          },
        ],
      },
    });
    const newSheetId = addRes.data.replies?.[0]?.addSheet?.properties?.sheetId;
    if (typeof newSheetId !== "number") {
      throw new Error("Failed to create analysis tab");
    }
    targetSheetId = newSheetId;
  }

  // 2. Compute row positions (1-based) for the new tab
  // Row 1: title
  // Row 2: timestamp
  // Row 3: stats
  // Row 4: blank
  // Row 5: header (Domain | Count | Percent)
  // Row 6+: data
  const titleRow = 1;
  const tsRow = 2;
  const statsRow = 3;
  const headerRow = 5;
  const dataStartRow = 6;
  const dataEndRow = dataStartRow + analysis.domains.length - 1;

  const safeName = `'${targetName.replace(/'/g, "''")}'`;

  // 3. Write the data
  const values: (string | number)[][] = [
    [`Domain Distribution — ${analysis.tabName}`],
    [`Generated ${new Date().toISOString()}`],
    [
      `${analysis.uniqueEmails} unique email(s) · ${analysis.totalEmails} total · ${analysis.domains.length} distinct domain(s)`,
    ],
    [],
    ["Domain", "Count", "Percent"],
    ...analysis.domains.map((d) => [d.domain, d.count, d.percent / 100]),
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${safeName}!A${titleRow}:C${dataEndRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values },
  });

  // 4. Format + chart
  const requests: sheets_v4.Schema$Request[] = [];

  // Title
  requests.push({
    repeatCell: {
      range: {
        sheetId: targetSheetId,
        startRowIndex: titleRow - 1,
        endRowIndex: titleRow,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      cell: {
        userEnteredFormat: {
          textFormat: { bold: true, fontSize: 14 },
        },
      },
      fields: "userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.fontSize",
    },
  });
  // Timestamp + stats: italic muted
  requests.push({
    repeatCell: {
      range: {
        sheetId: targetSheetId,
        startRowIndex: tsRow - 1,
        endRowIndex: statsRow,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      cell: {
        userEnteredFormat: {
          textFormat: {
            italic: true,
            foregroundColor: { red: 0.4, green: 0.4, blue: 0.4 },
          },
        },
      },
      fields:
        "userEnteredFormat.textFormat.italic,userEnteredFormat.textFormat.foregroundColor",
    },
  });
  // Bold the table header row
  requests.push({
    repeatCell: {
      range: {
        sheetId: targetSheetId,
        startRowIndex: headerRow - 1,
        endRowIndex: headerRow,
        startColumnIndex: 0,
        endColumnIndex: 3,
      },
      cell: {
        userEnteredFormat: {
          textFormat: { bold: true },
          backgroundColor: { red: 0.93, green: 0.93, blue: 0.96 },
        },
      },
      fields:
        "userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor",
    },
  });
  // Format percent column as percentage
  requests.push({
    repeatCell: {
      range: {
        sheetId: targetSheetId,
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

  // Add chart anchored to the right of the data table (column E)
  requests.push({
    addChart: {
      chart: {
        spec: {
          title: `Domain Distribution`,
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
                        sheetId: targetSheetId,
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
                        sheetId: targetSheetId,
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
              sheetId: targetSheetId,
              rowIndex: titleRow - 1,
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

  return targetName;
}
