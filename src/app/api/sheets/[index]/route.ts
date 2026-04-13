import { NextResponse } from "next/server";
import { removeLinkedSheet } from "@/lib/config-store";

export async function DELETE(
  _request: Request,
  { params }: { params: Promise<{ index: string }> }
) {
  try {
    const { index: indexStr } = await params;
    const index = parseInt(indexStr, 10);
    if (isNaN(index) || index < 0) {
      return NextResponse.json(
        { error: "Invalid index" },
        { status: 400 }
      );
    }

    await removeLinkedSheet(index);
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to remove sheet" },
      { status: 500 }
    );
  }
}
