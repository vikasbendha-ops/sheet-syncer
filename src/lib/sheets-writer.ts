import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { MasterSheetData } from "@/types";

const MASTER_TAB = "Master";

async function ensureMasterTab(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);

  const spreadsheet = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );

  const tabs = spreadsheet.data.sheets ?? [];
  const masterExists = tabs.some(
    (tab) => tab.properties?.title === MASTER_TAB
  );

  if (!masterExists) {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId: masterSheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title: MASTER_TAB } } }],
        },
      })
    );
  }
}

export async function writeMasterSheet(
  refreshToken: string,
  masterSheetId: string,
  data: MasterSheetData
): Promise<void> {
  await ensureMasterTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${MASTER_TAB}!A:ZZ`,
    })
  );

  const allRows = [data.headers, ...data.rows];
  if (allRows.length === 0) return;

  const lastCol = columnIndexToLetter(data.headers.length - 1);
  const range = `${MASTER_TAB}!A1:${lastCol}${allRows.length}`;

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range,
      valueInputOption: "RAW",
      requestBody: { values: allRows },
    })
  );
}

function columnIndexToLetter(index: number): string {
  let letter = "";
  let i = index;
  while (i >= 0) {
    letter = String.fromCharCode((i % 26) + 65) + letter;
    i = Math.floor(i / 26) - 1;
  }
  return letter;
}
