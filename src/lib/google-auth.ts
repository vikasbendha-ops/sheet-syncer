import { google, sheets_v4 } from "googleapis";
import { GoogleAuth } from "google-auth-library";

let cachedClient: sheets_v4.Sheets | null = null;

export function getSheetsClient(): sheets_v4.Sheets {
  if (cachedClient) return cachedClient;

  const keyJson = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!keyJson) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set");
  }

  const credentials = JSON.parse(keyJson);

  const auth = new GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  cachedClient = google.sheets({ version: "v4", auth });
  return cachedClient;
}

export function getMasterSheetId(): string {
  const id = process.env.MASTER_SHEET_ID;
  if (!id) {
    throw new Error("MASTER_SHEET_ID environment variable is not set");
  }
  return id;
}
