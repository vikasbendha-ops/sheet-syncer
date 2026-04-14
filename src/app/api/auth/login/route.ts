import { NextResponse } from "next/server";
import { getAuthUrl } from "@/lib/google-auth";

export async function GET() {
  try {
    const url = getAuthUrl();
    return NextResponse.redirect(url);
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to start auth" },
      { status: 500 }
    );
  }
}
