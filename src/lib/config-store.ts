import { getSheetsClient, getMasterSheetId } from "./google-auth";
import { LinkedSheet } from "@/types";

const CONFIG_TAB = "_config";

async function ensureConfigTab(refreshToken: string): Promise<void> {
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: masterSheetId,
  });

  const tabs = spreadsheet.data.sheets ?? [];
  const configExists = tabs.some(
    (tab) => tab.properties?.title === CONFIG_TAB
  );

  if (!configExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: masterSheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: CONFIG_TAB },
            },
          },
        ],
      },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A1:D1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [["url", "nickname", "emailColumn", "lastSynced"]],
      },
    });
  }
}

export async function getLinkedSheets(refreshToken: string): Promise<LinkedSheet[]> {
  await ensureConfigTab(refreshToken);
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A2:D1000`,
  });

  const rows = response.data.values ?? [];
  return rows
    .filter((row) => row[0])
    .map((row) => ({
      url: row[0] ?? "",
      nickname: row[1] ?? "",
      emailColumn: row[2] ?? "auto",
      lastSynced: row[3] ?? "",
    }));
}

export async function addLinkedSheet(
  refreshToken: string,
  url: string,
  nickname: string,
  emailColumn: string = "auto"
): Promise<void> {
  await ensureConfigTab(refreshToken);
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  await sheets.spreadsheets.values.append({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A:D`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[url, nickname, emailColumn, ""]],
    },
  });
}

export async function removeLinkedSheet(
  refreshToken: string,
  index: number
): Promise<void> {
  await ensureConfigTab(refreshToken);
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A2:D1000`,
  });

  const rows = response.data.values ?? [];
  if (index < 0 || index >= rows.length) {
    throw new Error(`Invalid sheet index: ${index}`);
  }

  rows.splice(index, 1);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A2:D1000`,
  });

  if (rows.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:D${rows.length + 1}`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    });
  }
}

export async function updateLastSynced(
  refreshToken: string,
  index: number,
  timestamp: string
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  const rowNumber = index + 2;
  await sheets.spreadsheets.values.update({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!D${rowNumber}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[timestamp]],
    },
  });
}
