import { NextRequest, NextResponse } from "next/server";
import { exchangeCode } from "@/lib/google-auth";
import { saveSession } from "@/lib/session";
import { google } from "googleapis";
import { OAuth2Client } from "google-auth-library";

export async function GET(request: NextRequest) {
  const code = request.nextUrl.searchParams.get("code");
  const error = request.nextUrl.searchParams.get("error");

  if (error) {
    return NextResponse.redirect(
      new URL(`/?error=${encodeURIComponent(error)}`, request.url)
    );
  }

  if (!code) {
    return NextResponse.redirect(
      new URL("/?error=no_code", request.url)
    );
  }

  try {
    const tokens = await exchangeCode(code);

    // Get user email
    let email = "";
    if (tokens.access_token) {
      const oauth2Client = new OAuth2Client();
      oauth2Client.setCredentials({ access_token: tokens.access_token });
      const oauth2 = google.oauth2({ version: "v2", auth: oauth2Client });
      const userInfo = await oauth2.userinfo.get();
      email = userInfo.data.email ?? "";
    }

    await saveSession({
      refreshToken: tokens.refresh_token ?? undefined,
      accessToken: tokens.access_token ?? undefined,
      expiresAt: tokens.expiry_date ?? undefined,
      email,
    });

    return NextResponse.redirect(new URL("/", request.url));
  } catch (err) {
    const message = err instanceof Error ? err.message : "Auth failed";
    return NextResponse.redirect(
      new URL(`/?error=${encodeURIComponent(message)}`, request.url)
    );
  }
}
