/**
 * Extracts the spreadsheet ID from a Google Sheets URL.
 * Supports URLs like:
 *   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
 *   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit#gid=0
 *   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID
 */
export function extractSpreadsheetId(url: string): string {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) {
    throw new Error(
      `Invalid Google Sheets URL: "${url}". Expected a URL like https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`
    );
  }
  return match[1];
}
