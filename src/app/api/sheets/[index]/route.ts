import { NextResponse } from "next/server";
import { z } from "zod";
import { removeLinkedSheet, updateLinkedSheet } from "@/lib/config-store";
import { getSession } from "@/lib/session";

const UpdateSheetSchema = z.object({
  nickname: z.string().min(1).max(100).optional(),
  tabName: z.string().min(1).optional(),
  emailColumn: z.string().min(1).optional(),
});

async function getRefreshTokenOrError(): Promise<
  | { ok: true; refreshToken: string }
  | { ok: false; response: NextResponse }
> {
  const session = await getSession();
  if (!session.refreshToken) {
    return {
      ok: false,
      response: NextResponse.json({ error: "Not authenticated" }, { status: 401 }),
    };
  }
  return { ok: true, refreshToken: session.refreshToken };
}

function parseIndex(indexStr: string): number | null {
  const index = parseInt(indexStr, 10);
  if (isNaN(index) || index < 0) return null;
  return index;
}

export async function DELETE(
  _request: Request,
  { params }: { params: Promise<{ index: string }> }
) {
  try {
    const auth = await getRefreshTokenOrError();
    if (!auth.ok) return auth.response;

    const { index: indexStr } = await params;
    const index = parseIndex(indexStr);
    if (index === null) {
      return NextResponse.json({ error: "Invalid index" }, { status: 400 });
    }

    await removeLinkedSheet(auth.refreshToken, index);
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to remove sheet" },
      { status: 500 }
    );
  }
}

export async function PATCH(
  request: Request,
  { params }: { params: Promise<{ index: string }> }
) {
  try {
    const auth = await getRefreshTokenOrError();
    if (!auth.ok) return auth.response;

    const { index: indexStr } = await params;
    const index = parseIndex(indexStr);
    if (index === null) {
      return NextResponse.json({ error: "Invalid index" }, { status: 400 });
    }

    const body = await request.json();
    const parsed = UpdateSheetSchema.parse(body);

    await updateLinkedSheet(auth.refreshToken, index, parsed);
    return NextResponse.json({ success: true });
  } catch (err) {
    if (err instanceof z.ZodError) {
      return NextResponse.json(
        {
          error: err.issues
            .map((e: { message: string }) => e.message)
            .join(", "),
        },
        { status: 400 }
      );
    }
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to update sheet" },
      { status: 500 }
    );
  }
}
