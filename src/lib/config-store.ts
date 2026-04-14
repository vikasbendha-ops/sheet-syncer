import { getSheetsClient, getMasterSheetId } from "./google-auth";
import { LinkedSheet } from "@/types";

const CONFIG_TAB = "_config";
const NEW_HEADERS = ["url", "nickname", "tabName", "emailColumn", "lastSynced"];

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
      range: `${CONFIG_TAB}!A1:E1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [NEW_HEADERS],
      },
    });
    return;
  }

  // Check if existing config uses old 4-column format and migrate
  const headerResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A1:E1`,
  });

  const headers = headerResponse.data.values?.[0] ?? [];

  // Old format: [url, nickname, emailColumn, lastSynced] (no tabName)
  if (headers.length <= 4 && headers[2] !== "tabName") {
    // Read all existing data rows in old format
    const dataResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:D1000`,
    });

    const oldRows = dataResponse.data.values ?? [];

    // Convert old rows to new format by inserting tabName column
    // Old: [url, nickname, emailColumn, lastSynced]
    // New: [url, nickname, tabName, emailColumn, lastSynced]
    const newRows = oldRows
      .filter((row) => row[0])
      .map((row) => [
        row[0] ?? "",           // url
        row[1] ?? "",           // nickname
        "",                     // tabName (empty = will use first tab)
        row[2] ?? "auto",       // emailColumn
        row[3] ?? "",           // lastSynced
      ]);

    // Clear everything and rewrite with new format
    await sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A:E`,
    });

    const allRows = [NEW_HEADERS, ...newRows];
    await sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A1:E${allRows.length}`,
      valueInputOption: "RAW",
      requestBody: { values: allRows },
    });
  }
}

export async function getLinkedSheets(refreshToken: string): Promise<LinkedSheet[]> {
  await ensureConfigTab(refreshToken);
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A2:E1000`,
  });

  const rows = response.data.values ?? [];
  return rows
    .filter((row) => row[0])
    .map((row) => ({
      url: row[0] ?? "",
      nickname: row[1] ?? "",
      tabName: row[2] ?? "",
      emailColumn: row[3] ?? "auto",
      lastSynced: row[4] ?? "",
    }));
}

export async function addLinkedSheet(
  refreshToken: string,
  url: string,
  nickname: string,
  tabName: string,
  emailColumn: string = "auto"
): Promise<void> {
  await ensureConfigTab(refreshToken);
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  await sheets.spreadsheets.values.append({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A:E`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[url, nickname, tabName, emailColumn, ""]],
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
    range: `${CONFIG_TAB}!A2:E1000`,
  });

  const rows = response.data.values ?? [];
  if (index < 0 || index >= rows.length) {
    throw new Error(`Invalid sheet index: ${index}`);
  }

  rows.splice(index, 1);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A2:E1000`,
  });

  if (rows.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:E${rows.length + 1}`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    });
  }
}

export async function updateLinkedSheet(
  refreshToken: string,
  index: number,
  updates: { nickname?: string; tabName?: string; emailColumn?: string }
): Promise<void> {
  await ensureConfigTab(refreshToken);
  const sheets = getSheetsClient(refreshToken);
  const masterSheetId = getMasterSheetId();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A2:E1000`,
  });

  const rows = response.data.values ?? [];
  if (index < 0 || index >= rows.length) {
    throw new Error(`Invalid sheet index: ${index}`);
  }

  const row = rows[index];
  if (updates.nickname !== undefined) row[1] = updates.nickname;
  if (updates.tabName !== undefined) row[2] = updates.tabName;
  if (updates.emailColumn !== undefined) row[3] = updates.emailColumn;

  const rowNumber = index + 2;
  await sheets.spreadsheets.values.update({
    spreadsheetId: masterSheetId,
    range: `${CONFIG_TAB}!A${rowNumber}:E${rowNumber}`,
    valueInputOption: "RAW",
    requestBody: { values: [row] },
  });
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
    range: `${CONFIG_TAB}!E${rowNumber}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[timestamp]],
    },
  });
}
