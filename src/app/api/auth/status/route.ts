import { NextResponse } from "next/server";
import { getSession } from "@/lib/session";

export async function GET() {
  try {
    const session = await getSession();
    if (session.refreshToken) {
      return NextResponse.json({
        authenticated: true,
        email: session.email ?? null,
      });
    }
    return NextResponse.json({ authenticated: false });
  } catch {
    return NextResponse.json({ authenticated: false });
  }
}
