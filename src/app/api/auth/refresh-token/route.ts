import { NextResponse } from "next/server";
import { getSession } from "@/lib/session";

export async function GET() {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      );
    }

    return NextResponse.json({ refreshToken: session.refreshToken });
  } catch {
    return NextResponse.json(
      { error: "Failed to get refresh token" },
      { status: 500 }
    );
  }
}
