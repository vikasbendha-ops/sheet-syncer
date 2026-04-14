import { google, sheets_v4 } from "googleapis";
import { OAuth2Client } from "google-auth-library";

let cachedOAuthClient: OAuth2Client | null = null;

function getOAuthClient(): OAuth2Client {
  if (cachedOAuthClient) return cachedOAuthClient;

  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const redirectUri = process.env.GOOGLE_REDIRECT_URI;

  if (!clientId || !clientSecret || !redirectUri) {
    throw new Error(
      "GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and GOOGLE_REDIRECT_URI must be set"
    );
  }

  cachedOAuthClient = new OAuth2Client(clientId, clientSecret, redirectUri);
  return cachedOAuthClient;
}

export function getAuthUrl(): string {
  const client = getOAuthClient();
  return client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/userinfo.email",
    ],
  });
}

export async function exchangeCode(code: string) {
  const client = getOAuthClient();
  const { tokens } = await client.getToken(code);
  return tokens;
}

export function getSheetsClient(refreshToken: string): sheets_v4.Sheets {
  const client = getOAuthClient();
  client.setCredentials({ refresh_token: refreshToken });
  return google.sheets({ version: "v4", auth: client });
}

