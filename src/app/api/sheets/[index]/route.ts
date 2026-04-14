import { NextResponse } from "next/server";
import { z } from "zod";
import { removeLinkedSheet, updateLinkedSheet } from "@/lib/config-store";
import { getSession } from "@/lib/session";

const UpdateSheetSchema = z.object({
  nickname: z.string().min(1).max(100).optional(),
  tabName: z.string().min(1).optional(),
  emailColumn: z.string().min(1).optional(),
});

interface AuthContext {
  refreshToken: string;
  masterSheetId: string;
}

async function getAuthContext(): Promise<
  { ok: true; ctx: AuthContext } | { ok: false; response: NextResponse }
> {
  const session = await getSession();
  if (!session.refreshToken) {
    return {
      ok: false,
      response: NextResponse.json({ error: "Not authenticated" }, { status: 401 }),
    };
  }
  if (!session.masterSheetId) {
    return {
      ok: false,
      response: NextResponse.json(
        { error: "No master sheet configured" },
        { status: 400 }
      ),
    };
  }
  return {
    ok: true,
    ctx: {
      refreshToken: session.refreshToken,
      masterSheetId: session.masterSheetId,
    },
  };
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
    const auth = await getAuthContext();
    if (!auth.ok) return auth.response;

    const { index: indexStr } = await params;
    const index = parseIndex(indexStr);
    if (index === null) {
      return NextResponse.json({ error: "Invalid index" }, { status: 400 });
    }

    await removeLinkedSheet(auth.ctx.refreshToken, auth.ctx.masterSheetId, index);
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
    const auth = await getAuthContext();
    if (!auth.ok) return auth.response;

    const { index: indexStr } = await params;
    const index = parseIndex(indexStr);
    if (index === null) {
      return NextResponse.json({ error: "Invalid index" }, { status: 400 });
    }

    const body = await request.json();
    const parsed = UpdateSheetSchema.parse(body);

    await updateLinkedSheet(
      auth.ctx.refreshToken,
      auth.ctx.masterSheetId,
      index,
      parsed
    );
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
