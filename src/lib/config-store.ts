import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import { LinkedSheet } from "@/types";

const CONFIG_TAB = "_config";
const NEW_HEADERS = ["url", "nickname", "tabName", "emailColumn", "lastSynced"];

async function ensureConfigTab(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);

  const spreadsheet = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );

  const tabs = spreadsheet.data.sheets ?? [];
  const configExists = tabs.some((tab) => tab.properties?.title === CONFIG_TAB);

  if (!configExists) {
    await withRetry(() =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId: masterSheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title: CONFIG_TAB } } }],
        },
      })
    );

    await withRetry(() =>
      sheets.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${CONFIG_TAB}!A1:E1`,
        valueInputOption: "RAW",
        requestBody: { values: [NEW_HEADERS] },
      })
    );
    return;
  }

  const headerResponse = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A1:E1`,
    })
  );

  const headers = headerResponse.data.values?.[0] ?? [];

  if (headers.length <= 4 && headers[2] !== "tabName") {
    const dataResponse = await withRetry(() =>
      sheets.spreadsheets.values.get({
        spreadsheetId: masterSheetId,
        range: `${CONFIG_TAB}!A2:D1000`,
      })
    );

    const oldRows = dataResponse.data.values ?? [];

    const newRows = oldRows
      .filter((row) => row[0])
      .map((row) => [
        row[0] ?? "",
        row[1] ?? "",
        "",
        row[2] ?? "auto",
        row[3] ?? "",
      ]);

    await withRetry(() =>
      sheets.spreadsheets.values.clear({
        spreadsheetId: masterSheetId,
        range: `${CONFIG_TAB}!A:E`,
      })
    );

    const allRows = [NEW_HEADERS, ...newRows];
    await withRetry(() =>
      sheets.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${CONFIG_TAB}!A1:E${allRows.length}`,
        valueInputOption: "RAW",
        requestBody: { values: allRows },
      })
    );
  }
}

export async function getLinkedSheets(
  refreshToken: string,
  masterSheetId: string
): Promise<LinkedSheet[]> {
  await ensureConfigTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const response = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:E1000`,
    })
  );

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
  masterSheetId: string,
  url: string,
  nickname: string,
  tabName: string,
  emailColumn: string = "auto"
): Promise<void> {
  await ensureConfigTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.append({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A:E`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[url, nickname, tabName, emailColumn, ""]],
      },
    })
  );
}

export async function removeLinkedSheet(
  refreshToken: string,
  masterSheetId: string,
  index: number
): Promise<void> {
  await ensureConfigTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const response = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:E1000`,
    })
  );

  const rows = response.data.values ?? [];
  if (index < 0 || index >= rows.length) {
    throw new Error(`Invalid sheet index: ${index}`);
  }

  rows.splice(index, 1);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:E1000`,
    })
  );

  if (rows.length > 0) {
    await withRetry(() =>
      sheets.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${CONFIG_TAB}!A2:E${rows.length + 1}`,
        valueInputOption: "RAW",
        requestBody: { values: rows },
      })
    );
  }
}

export async function updateLinkedSheet(
  refreshToken: string,
  masterSheetId: string,
  index: number,
  updates: { nickname?: string; tabName?: string; emailColumn?: string }
): Promise<void> {
  await ensureConfigTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const response = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A2:E1000`,
    })
  );

  const rows = response.data.values ?? [];
  if (index < 0 || index >= rows.length) {
    throw new Error(`Invalid sheet index: ${index}`);
  }

  const row = rows[index];
  if (updates.nickname !== undefined) row[1] = updates.nickname;
  if (updates.tabName !== undefined) row[2] = updates.tabName;
  if (updates.emailColumn !== undefined) row[3] = updates.emailColumn;

  const rowNumber = index + 2;
  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${CONFIG_TAB}!A${rowNumber}:E${rowNumber}`,
      valueInputOption: "RAW",
      requestBody: { values: [row] },
    })
  );
}

/**
 * Batches many lastSynced updates into a single API call.
 * indices are 0-based positions in the _config data rows.
 */
export async function updateLastSyncedBatch(
  refreshToken: string,
  masterSheetId: string,
  indices: number[],
  timestamp: string
): Promise<void> {
  if (indices.length === 0) return;
  const sheets = getSheetsClient(refreshToken);

  const data = indices.map((index) => {
    const rowNumber = index + 2;
    return {
      range: `${CONFIG_TAB}!E${rowNumber}`,
      values: [[timestamp]],
    };
  });

  await withRetry(() =>
    sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: masterSheetId,
      requestBody: { valueInputOption: "RAW", data },
    })
  );
}
