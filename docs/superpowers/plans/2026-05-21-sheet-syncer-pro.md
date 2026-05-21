# Sheet Syncer Pro Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new feature module "Sync Pro" — a more flexible sibling of the existing basic Present In main sync. Adds (1) column propagation between linked source sheets (fill blanks across sheets for selected columns, never overwrite, surface conflicts) and (2) multiple independent sections per page.

**Architecture:** Separate feature module mirroring the existing `consolidator` layout. Basic `sync-engine.ts` stays 100% untouched. Prefix-folding name-merge logic is duplicated into the Pro engine. Each Pro section is independent; per-section config persisted in a hidden `_sync_pro_config` tab in the user's master sheet.

**Tech Stack:** Next.js 16 App Router, React 19, TypeScript strict, `googleapis` / `google-auth-library`, Tailwind v4. Shared helpers (`getSheetsClient`, `withRetry`, `extractSpreadsheetId`, `normalizeEmail`, `findEmailColumnIndex`, `findNameColumns`, `columnIndexToLetter`, `getTabNames`, `writePresentInColumn`) reused as-is.

**Verification convention:** This repo has no test runner. Each task ends with `npx tsc --noEmit && npm run lint` as the verification gate. Final task includes manual smoke testing via `npm run dev`.

**Spec:** [`docs/superpowers/specs/2026-05-21-sheet-syncer-pro-design.md`](../specs/2026-05-21-sheet-syncer-pro-design.md)

---

## File Structure

**New files:**

- `src/lib/sync-pro-types.ts` — shared types (config + result shapes). Imported by engine, config-store, API routes, page.
- `src/lib/sync-pro-config-store.ts` — `_sync_pro_config` tab schema, `ensureTab` + `get/save/clear`.
- `src/lib/sync-pro-engine.ts` — `runSyncProSection`, `runSyncProBatch`, internal helpers (per-sheet reader, propagation pass, name-merge — duplicated from `sync-engine.ts`).
- `src/app/api/sync-pro/route.ts` — `POST` runs one or many sections.
- `src/app/api/sync-pro/config/route.ts` — `GET` / `PUT` / `DELETE` config.
- `src/app/sync-pro/page.tsx` — multi-section UI.

**Modified files:**

- `src/components/nav.tsx` — add `{ href: "/sync-pro", label: "Sync Pro" }` entry.
- `CLAUDE.md` — add feature note + new hidden tab in storage model.

**Reused without changes:** `src/lib/google-auth.ts`, `src/lib/retry.ts`, `src/lib/url-parser.ts`, `src/lib/email-utils.ts`, `src/lib/present-in-writer.ts`, `src/lib/sheets-reader.ts` (only `getTabNames`).

---

## Task 1: Shared types module

Defines all interfaces shared between engine, config-store, routes, and UI. Doing this first locks the contract for every later task.

**Files:**

- Create: `src/lib/sync-pro-types.ts`

- [ ] **Step 1: Create the types file**

```typescript
// src/lib/sync-pro-types.ts

/**
 * Sheet Syncer Pro — shared types.
 *
 * Pro = multi-section, column-propagation sibling of the basic main sync.
 * Match key is always lowercased email. Each section is independent.
 */

/** A single propagate column defined on a section. */
export interface ProPropagateColumn {
  /** User-defined logical name, e.g. "Phone" or "Course". */
  name: string;
}

/**
 * One source sheet linked into a section. Each linked sheet contributes
 * rows that are joined by email with every other linked sheet in the same
 * section. The `columnMapping` says which actual header in this sheet
 * corresponds to each of the section's logical propagate columns; a value
 * of null means "skip this sheet for this column".
 */
export interface ProLinkedSheet {
  url: string;
  nickname: string;
  /** Single tab per linked sheet. */
  tabName: string;
  /** "auto" or column letter (A-Z, AA, etc.). */
  emailColumn: string;
  /**
   * Map from logical column name (from section.propagateColumns) → actual
   * header text in this sheet. Null = skip this sheet for that column.
   */
  columnMapping: Record<string, string | null>;
}

/** One independent Pro sync config. */
export interface ProSection {
  /** Stable id (sec_<timestamp>_<rand>). Lets us track sections across reloads. */
  id: string;
  /** User-typed display name. */
  name: string;
  /** Tab written into the master sheet. Default `Pro: <name>`. */
  masterTabName: string;
  linkedSheets: ProLinkedSheet[];
  propagateColumns: ProPropagateColumn[];
  /** When false, the Present In column write-back step is skipped. */
  writePresentIn: boolean;
}

/** Per-sheet outcome of one section run. */
export interface ProTabResult {
  nickname: string;
  url: string;
  tabName: string;
  /** Non-empty rows read from this sheet (rows past the header with any value). */
  rowsRead: number;
  /** Rows whose email column parsed to a valid normalized email. */
  emailsFound: number;
  /** Cells in this sheet that got filled in by propagation. */
  cellsFilled: number;
  error?: string;
}

/** Stats for one logical propagate column within a section run. */
export interface ProColumnStats {
  /** Logical column name. */
  name: string;
  /** Cells filled across all sheets for this column. */
  cellsFilled: number;
  /** Distinct emails where 2+ sheets had different non-blank values. */
  conflicts: number;
  /** Linked sheets that mapped this column to "skip". */
  skippedSheets: number;
}

/** A single column-value conflict where blank-fill couldn't decide a winner. */
export interface ProConflict {
  email: string;
  /** Logical column name. */
  column: string;
  /** Values per sheet (only sheets that had a non-blank value). */
  values: Array<{ nickname: string; value: string }>;
}

/** Full result of one section run. */
export interface ProSectionResult {
  sectionId: string;
  sectionName: string;
  /** Deep link to the master tab written for this section. Empty on early failure. */
  masterSpreadsheetUrl: string;
  masterTabName: string;
  /** Distinct normalized emails seen across all linked sheets. */
  totalUniqueEmails: number;
  /** Sum of cells filled across all sheets / columns. */
  totalCellsFilled: number;
  totalConflicts: number;
  linkedSheets: ProTabResult[];
  columnStats: ProColumnStats[];
  /** Detailed conflict log. UI usually shows first ~10. */
  conflicts: ProConflict[];
  presentInWritten: boolean;
  /** Fatal: whole section bailed before producing useful output. */
  error?: string;
}

export interface ProBatchResult {
  sections: ProSectionResult[];
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/lib/sync-pro-types.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): shared types module

Defines ProSection, ProLinkedSheet, ProPropagateColumn, ProTabResult,
ProColumnStats, ProConflict, ProSectionResult, ProBatchResult — the
contracts shared between the Sync Pro engine, config-store, API
routes, and page UI.

Spec: docs/superpowers/specs/2026-05-21-sheet-syncer-pro-design.md

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: Config store (`_sync_pro_config` tab)

Mirrors the consolidator config-store. JSON blobs for the two array columns. Hidden tab inside the master sheet.

**Files:**

- Create: `src/lib/sync-pro-config-store.ts`

- [ ] **Step 1: Create the config-store file**

```typescript
// src/lib/sync-pro-config-store.ts
import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import type {
  ProLinkedSheet,
  ProPropagateColumn,
  ProSection,
} from "./sync-pro-types";

/**
 * Hidden config tab in the user's master spreadsheet. One row per Pro
 * section. The two array columns (`propagateColumns`, `linkedSheets`)
 * are stored as JSON strings to keep the schema flat — same pattern
 * the consolidator uses for its `sources` column.
 */

const TAB = "_sync_pro_config";
const HEADERS = [
  "sectionId",
  "sectionName",
  "masterTabName",
  "writePresentIn",
  "propagateColumns",
  "linkedSheets",
];

export interface SyncProConfig {
  sections: ProSection[];
}

const EMPTY: SyncProConfig = { sections: [] };

function genId(): string {
  return `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

async function ensureTab(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId: masterSheetId })
  );
  const exists = (meta.data.sheets ?? []).some(
    (t) => t.properties?.title === TAB
  );
  if (exists) return;

  await withRetry(() =>
    sheets.spreadsheets.batchUpdate({
      spreadsheetId: masterSheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: TAB } } }],
      },
    })
  );
  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A1:F1`,
      valueInputOption: "RAW",
      requestBody: { values: [HEADERS] },
    })
  );
}

function parsePropagateColumns(raw: unknown): ProPropagateColumn[] {
  if (typeof raw !== "string" || !raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .filter(
        (c): c is { name?: unknown } =>
          typeof c === "object" && c !== null
      )
      .map((c) => ({ name: typeof c.name === "string" ? c.name : "" }))
      .filter((c) => c.name);
  } catch {
    return [];
  }
}

function parseLinkedSheets(raw: unknown): ProLinkedSheet[] {
  if (typeof raw !== "string" || !raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .filter(
        (s): s is Record<string, unknown> =>
          typeof s === "object" && s !== null
      )
      .map((s) => ({
        url: typeof s.url === "string" ? s.url : "",
        nickname: typeof s.nickname === "string" ? s.nickname : "",
        tabName: typeof s.tabName === "string" ? s.tabName : "",
        emailColumn:
          typeof s.emailColumn === "string" ? s.emailColumn : "auto",
        columnMapping:
          typeof s.columnMapping === "object" && s.columnMapping !== null
            ? Object.fromEntries(
                Object.entries(s.columnMapping as Record<string, unknown>).map(
                  ([k, v]) => [
                    k,
                    typeof v === "string" ? v : v === null ? null : null,
                  ]
                )
              )
            : {},
      }));
  } catch {
    return [];
  }
}

export async function getSyncProConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<SyncProConfig> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  const res = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:F1000`,
    })
  );
  const rows = res.data.values ?? [];
  if (!rows.length) return { ...EMPTY };

  const sections: ProSection[] = [];
  for (const row of rows) {
    const id = typeof row[0] === "string" && row[0] ? row[0] : genId();
    const name = typeof row[1] === "string" ? row[1] : "";
    const masterTabName = typeof row[2] === "string" ? row[2] : "";
    const writePresentIn =
      typeof row[3] === "string"
        ? row[3].toLowerCase() !== "false"
        : true;
    const propagateColumns = parsePropagateColumns(row[4]);
    const linkedSheets = parseLinkedSheets(row[5]);

    if (
      !name &&
      !masterTabName &&
      linkedSheets.length === 0 &&
      propagateColumns.length === 0
    ) {
      continue;
    }

    sections.push({
      id,
      name,
      masterTabName,
      writePresentIn,
      propagateColumns,
      linkedSheets,
    });
  }
  return { sections };
}

export async function saveSyncProConfig(
  refreshToken: string,
  masterSheetId: string,
  config: SyncProConfig
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A:F`,
    })
  );

  const rows: string[][] = [HEADERS];
  for (const s of config.sections) {
    rows.push([
      s.id || genId(),
      s.name ?? "",
      s.masterTabName ?? "",
      s.writePresentIn ? "true" : "false",
      JSON.stringify(s.propagateColumns ?? []),
      JSON.stringify(s.linkedSheets ?? []),
    ]);
  }

  await withRetry(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A1:F${rows.length}`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    })
  );
}

export async function clearSyncProConfig(
  refreshToken: string,
  masterSheetId: string
): Promise<void> {
  await ensureTab(refreshToken, masterSheetId);
  const sheets = getSheetsClient(refreshToken);

  await withRetry(() =>
    sheets.spreadsheets.values.clear({
      spreadsheetId: masterSheetId,
      range: `${TAB}!A2:F1000`,
    })
  );
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/lib/sync-pro-config-store.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): config store backed by _sync_pro_config hidden tab

One row per section. Two JSON columns (propagateColumns,
linkedSheets) keep the schema flat — same pattern as the
consolidator config store. Provides ensureTab + get/save/clear.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: Config API route (`/api/sync-pro/config`)

`GET` returns current config, `PUT` saves, `DELETE` wipes. 401 if no session, 400 with `code: "no_master_sheet"` if no master sheet (so the UI can show its banner — same shape as the consolidator config route).

**Files:**

- Create: `src/app/api/sync-pro/config/route.ts`

- [ ] **Step 1: Create the route file**

```typescript
// src/app/api/sync-pro/config/route.ts
import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import {
  getSyncProConfig,
  saveSyncProConfig,
  clearSyncProConfig,
} from "@/lib/sync-pro-config-store";

const LinkedSheetSchema = z.object({
  url: z.string(),
  nickname: z.string(),
  tabName: z.string(),
  emailColumn: z.string(),
  columnMapping: z.record(z.string(), z.union([z.string(), z.null()])),
});

const SectionSchema = z.object({
  id: z.string(),
  name: z.string(),
  masterTabName: z.string(),
  writePresentIn: z.boolean(),
  propagateColumns: z.array(z.object({ name: z.string() })),
  linkedSheets: z.array(LinkedSheetSchema),
});

const Schema = z.object({
  sections: z.array(SectionSchema),
});

async function requireSession() {
  const session = await getSession();
  if (!session.refreshToken) {
    return {
      error: NextResponse.json(
        { error: "Not authenticated" },
        { status: 401 }
      ),
    };
  }
  if (!session.masterSheetId) {
    return {
      error: NextResponse.json(
        {
          error:
            "No master sheet configured. Set one on the Sheets page first.",
          code: "no_master_sheet",
        },
        { status: 400 }
      ),
    };
  }
  return {
    refreshToken: session.refreshToken,
    masterSheetId: session.masterSheetId,
  };
}

export async function GET() {
  const session = await requireSession();
  if ("error" in session) return session.error;
  try {
    const config = await getSyncProConfig(
      session.refreshToken,
      session.masterSheetId
    );
    return NextResponse.json(config);
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to load config" },
      { status: 500 }
    );
  }
}

export async function PUT(request: Request) {
  const session = await requireSession();
  if ("error" in session) return session.error;
  try {
    const body = await request.json();
    const parsed = Schema.parse(body);
    await saveSyncProConfig(
      session.refreshToken,
      session.masterSheetId,
      parsed
    );
    return NextResponse.json({ ok: true });
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
      { error: err instanceof Error ? err.message : "Failed to save config" },
      { status: 500 }
    );
  }
}

export async function DELETE() {
  const session = await requireSession();
  if ("error" in session) return session.error;
  try {
    await clearSyncProConfig(session.refreshToken, session.masterSheetId);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Failed to clear config" },
      { status: 500 }
    );
  }
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/app/api/sync-pro/config/route.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): config API route (GET / PUT / DELETE)

Mirrors consolidator/biz-tutor-sync config-route shape:
- GET returns SyncProConfig.
- PUT accepts and saves (zod-validated).
- DELETE clears all sections.
- 401 / 400 (with code:"no_master_sheet") behavior matches existing
  feature config routes so the page can show its master-sheet banner.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: Engine skeleton + helpers + name-merge (duplicated)

Lays down the engine file with module-level helpers and the duplicated prefix-folding name-merge function from `sync-engine.ts`. No section-run logic yet — that comes in Tasks 5–7.

**Files:**

- Create: `src/lib/sync-pro-engine.ts`

- [ ] **Step 1: Create the engine file with helpers**

```typescript
// src/lib/sync-pro-engine.ts
import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import {
  normalizeEmail,
  findEmailColumnIndex,
  findNameColumns,
  columnIndexToLetter,
} from "./email-utils";
import { extractSpreadsheetId } from "./url-parser";
import { writePresentInColumn } from "./present-in-writer";
import { sheets_v4 } from "googleapis";
import type {
  ProBatchResult,
  ProColumnStats,
  ProConflict,
  ProLinkedSheet,
  ProSection,
  ProSectionResult,
  ProTabResult,
} from "./sync-pro-types";

/**
 * Sheet Syncer Pro engine.
 *
 * Each section is independent. Within a section:
 *   1. Read every linked sheet's email + name + mapped propagate columns.
 *   2. Cross-reference emails to build {email: which sheets contain it}.
 *   3. Propagation pass: for each logical propagate column, fill blanks
 *      from non-blank values. Never overwrite. Surface conflicts when
 *      multiple sheets have different non-blank values.
 *   4. Apply per-sheet writes (one batchUpdate per sheet).
 *   5. Write Present In column into each source sheet (optional).
 *   6. Write Master ✅/❌ tab into the master sheet.
 */

// Lowered from 5 to stay under Google Sheets' 60 read/min per-user quota,
// matching the main sync engine.
const MAX_CONCURRENCY = 3;

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

async function withConcurrencyLimit<T>(
  tasks: (() => Promise<T>)[],
  limit: number
): Promise<T[]> {
  const results: T[] = [];
  const executing = new Set<Promise<void>>();
  for (const task of tasks) {
    const p = task().then((result) => {
      results.push(result);
      executing.delete(p);
    });
    executing.add(p);
    if (executing.size >= limit) {
      await Promise.race(executing);
    }
  }
  await Promise.all(executing);
  return results;
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    return value.toLocaleString("en-US", {
      useGrouping: false,
      maximumFractionDigits: 20,
    });
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim();
}

/**
 * Resolve a 1-based column letter (or "auto") to a 0-based column index.
 * `auto` returns -1 and the caller should fall back to header detection.
 */
function letterToIndex(letter: string): number {
  if (!letter || letter.toLowerCase() === "auto") return -1;
  const upper = letter.toUpperCase();
  let idx = 0;
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    idx = idx * 26 + (code - 64);
  }
  return idx - 1;
}

/**
 * Merge names across sheets for the Master tab.
 *
 * DUPLICATED FROM src/lib/sync-engine.ts so the basic sync stays 100%
 * untouched (per spec / user direction). If a bug is found in either
 * copy, fix both.
 *
 * Strategy:
 *  1. Count every name occurrence (case-insensitive).
 *  2. If name A is a word-prefix of name B (e.g. "Laura" ⊂ "Laura
 *     Pegoraro"), fold A's count into B so the more complete name wins.
 *  3. Pick winner by mergedCount → words.length → firstSeenIdx.
 */
function mergeNamesForEmail(namesByEmail: Map<string, string[]>): Map<string, string> {
  const out = new Map<string, string>();
  for (const [email, names] of namesByEmail) {
    interface Candidate {
      name: string;
      words: string[];
      count: number;
      mergedCount: number;
      firstSeenIdx: number;
    }
    const candidates = new Map<string, Candidate>();
    let nextIdx = 0;
    for (const name of names) {
      const key = name.toLowerCase();
      const existing = candidates.get(key);
      if (existing) {
        existing.count++;
        existing.mergedCount++;
      } else {
        candidates.set(key, {
          name,
          words: key.split(/\s+/).filter(Boolean),
          count: 1,
          mergedCount: 1,
          firstSeenIdx: nextIdx++,
        });
      }
    }
    if (candidates.size === 0) continue;

    const list = Array.from(candidates.values());
    for (const a of list) {
      for (const b of list) {
        if (a === b) continue;
        if (b.words.length <= a.words.length) continue;
        const isPrefix = a.words.every((w, i) => b.words[i] === w);
        if (isPrefix) b.mergedCount += a.count;
      }
    }
    list.sort((a, b) => {
      if (b.mergedCount !== a.mergedCount) return b.mergedCount - a.mergedCount;
      if (b.words.length !== a.words.length) return b.words.length - a.words.length;
      return a.firstSeenIdx - b.firstSeenIdx;
    });
    out.set(email, list[0].name);
  }
  return out;
}

// ===========================================================================
// Per-sheet read + section run will be implemented in subsequent tasks.
// ===========================================================================

export async function runSyncProSection(
  refreshToken: string,
  masterSheetId: string,
  section: ProSection
): Promise<ProSectionResult> {
  // Placeholder — implemented incrementally over Tasks 5-7.
  const result: ProSectionResult = {
    sectionId: section.id,
    sectionName: section.name,
    masterSpreadsheetUrl: "",
    masterTabName: section.masterTabName || `Pro: ${section.name}`,
    totalUniqueEmails: 0,
    totalCellsFilled: 0,
    totalConflicts: 0,
    linkedSheets: [],
    columnStats: [],
    conflicts: [],
    presentInWritten: false,
    error: "runSyncProSection: not yet implemented",
  };
  // Silence unused-import warnings while skeleton compiles.
  void refreshToken;
  void masterSheetId;
  void getSheetsClient;
  void withRetry;
  void withConcurrencyLimit;
  void clean;
  void letterToIndex;
  void normalizeEmail;
  void findEmailColumnIndex;
  void findNameColumns;
  void columnIndexToLetter;
  void extractSpreadsheetId;
  void writePresentInColumn;
  void mergeNamesForEmail;
  void HEADER_BG;
  const _unused: sheets_v4.Sheets | null = null;
  void _unused;
  void ({} as ProLinkedSheet);
  void ({} as ProTabResult);
  void ({} as ProColumnStats);
  void ({} as ProConflict);
  return result;
}

export async function runSyncProBatch(
  refreshToken: string,
  masterSheetId: string,
  sections: ProSection[]
): Promise<ProBatchResult> {
  const results: ProSectionResult[] = [];
  for (const section of sections) {
    try {
      results.push(await runSyncProSection(refreshToken, masterSheetId, section));
    } catch (err) {
      results.push({
        sectionId: section.id,
        sectionName: section.name,
        masterSpreadsheetUrl: "",
        masterTabName: section.masterTabName || `Pro: ${section.name}`,
        totalUniqueEmails: 0,
        totalCellsFilled: 0,
        totalConflicts: 0,
        linkedSheets: [],
        columnStats: [],
        conflicts: [],
        presentInWritten: false,
        error: err instanceof Error ? err.message : "Section failed",
      });
    }
  }
  return { sections: results };
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors (the `void` calls and `_unused` keep unused-import / unused-variable lints quiet while the skeleton is incomplete).

- [ ] **Step 3: Commit**

```bash
git add src/lib/sync-pro-engine.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): engine skeleton + helpers + duplicated name-merge

Scaffolds runSyncProSection / runSyncProBatch (returning placeholder
results), plus module-level helpers:
- withConcurrencyLimit (copied; MAX_CONCURRENCY=3 to match basic sync)
- clean (unknown → safe string, handles JS numbers without sci. notation)
- letterToIndex (A → 0, "auto" → -1)
- mergeNamesForEmail (prefix-folding name merge, duplicated verbatim
  from src/lib/sync-engine.ts per spec direction to keep basic sync
  100% untouched)

Body of runSyncProSection is implemented incrementally over Tasks 5-7.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: Engine — read linked sheets + cross-reference

Implements the read phase: read each linked sheet's email + name + mapped propagate columns, build per-sheet email→row records, and build the cross-reference map. Replaces the placeholder section runner from Task 4.

**Files:**

- Modify: `src/lib/sync-pro-engine.ts`

- [ ] **Step 1: Replace the `runSyncProSection` body and add the read helper**

Find the block `// ===========================================================================` through the end of `runSyncProSection` (the placeholder version from Task 4). Replace it entirely with:

```typescript
// ===========================================================================
// Read phase
// ===========================================================================

/** A single read row from one linked sheet, keyed by normalized email. */
interface SheetRowRecord {
  /** 1-based row number in the source tab. */
  rowNumber: number;
  /** Display name extracted from the sheet's name column(s), if any. */
  name: string;
  /**
   * Map from logical column name → raw cell value (cleaned). Only contains
   * columns this linked sheet mapped to a real header.
   */
  perColumn: Map<string, string>;
}

interface SheetReadOutcome {
  result: ProTabResult;
  /** Resolved Google Sheets numeric sheetId for the picked tab. */
  sheetId: number;
  /** Resolved spreadsheetId. */
  spreadsheetId: string;
  /** Map<normalizedEmail, SheetRowRecord>. First-seen row wins per email. */
  byEmail: Map<string, SheetRowRecord>;
  /**
   * The actual column index in this sheet for each logical column the
   * section requested. -1 = not mapped / not found in this sheet.
   */
  logicalColIndex: Map<string, number>;
  /** 0-based index of the email column in this sheet. */
  emailColIdx: number;
}

/**
 * Reads one linked sheet. Resolves the email column (auto-detect or letter),
 * resolves each mapped logical column to a 0-based index, then bulk-reads
 * the values via batchGet.
 */
async function readLinkedSheet(
  sheets: sheets_v4.Sheets,
  linked: ProLinkedSheet,
  propagateColumns: string[]
): Promise<SheetReadOutcome> {
  const tabResult: ProTabResult = {
    nickname: linked.nickname,
    url: linked.url,
    tabName: linked.tabName,
    rowsRead: 0,
    emailsFound: 0,
    cellsFilled: 0,
  };
  const byEmail = new Map<string, SheetRowRecord>();
  const logicalColIndex = new Map<string, number>();

  const spreadsheetId = extractSpreadsheetId(linked.url);

  // Resolve tab + sheetId
  const meta = await withRetry(() =>
    sheets.spreadsheets.get({ spreadsheetId })
  );
  const allTabs = meta.data.sheets ?? [];
  const matched = allTabs.find(
    (t) => t.properties?.title === linked.tabName
  );
  const sheetId = matched?.properties?.sheetId ?? -1;
  const resolvedTab = matched?.properties?.title ?? linked.tabName;
  if (sheetId < 0) {
    tabResult.error = `Tab "${linked.tabName}" not found`;
    return {
      result: tabResult,
      sheetId: -1,
      spreadsheetId,
      byEmail,
      logicalColIndex,
      emailColIdx: -1,
    };
  }

  const safeTab = `'${resolvedTab.replace(/'/g, "''")}'`;

  // Read header row to resolve column indices
  const headerRes = await withRetry(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeTab}!1:1`,
    })
  );
  const headers = (headerRes.data.values?.[0] ?? []).map((h) =>
    clean(h)
  );

  // Email column
  let emailColIdx = letterToIndex(linked.emailColumn);
  if (emailColIdx < 0) {
    emailColIdx = findEmailColumnIndex(headers);
  }
  if (emailColIdx < 0) {
    tabResult.error = `No email column found in "${resolvedTab}". Headers: ${headers.filter(Boolean).join(", ")}`;
    return {
      result: tabResult,
      sheetId,
      spreadsheetId,
      byEmail,
      logicalColIndex,
      emailColIdx: -1,
    };
  }

  // Name columns (for the Master tab's Name field)
  const nameCols = findNameColumns(headers);

  // Logical propagate columns → 0-based column indices (case-insensitive)
  const normalizedHeaders = headers.map((h) =>
    h.toLowerCase().trim().replace(/\s+/g, " ")
  );
  for (const logical of propagateColumns) {
    const mapped = linked.columnMapping[logical];
    if (!mapped) {
      logicalColIndex.set(logical, -1);
      continue;
    }
    const want = mapped.toLowerCase().trim().replace(/\s+/g, " ");
    const idx = normalizedHeaders.findIndex((h) => h === want);
    logicalColIndex.set(logical, idx);
  }

  // Build batch ranges: email + name columns + every mapped (idx >= 0) logical column
  const ranges: string[] = [];
  const rangeForCol = new Map<number, number>(); // col idx → ranges[] index
  function addColumnRange(colIdx: number) {
    if (colIdx < 0 || rangeForCol.has(colIdx)) return;
    rangeForCol.set(colIdx, ranges.length);
    const letter = columnIndexToLetter(colIdx);
    ranges.push(`${safeTab}!${letter}2:${letter}`);
  }
  addColumnRange(emailColIdx);
  if (nameCols.fullName !== undefined) addColumnRange(nameCols.fullName);
  if (nameCols.firstName !== undefined) addColumnRange(nameCols.firstName);
  if (nameCols.lastName !== undefined) addColumnRange(nameCols.lastName);
  for (const [, idx] of logicalColIndex) addColumnRange(idx);

  const batchRes = await withRetry(() =>
    sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges,
      valueRenderOption: "UNFORMATTED_VALUE",
    })
  );
  const valueRanges = batchRes.data.valueRanges ?? [];

  function valuesFor(colIdx: number): unknown[][] {
    if (colIdx < 0) return [];
    const ri = rangeForCol.get(colIdx);
    if (ri === undefined) return [];
    return (valueRanges[ri]?.values ?? []) as unknown[][];
  }

  const emailRows = valuesFor(emailColIdx);
  const fullNameRows =
    nameCols.fullName !== undefined ? valuesFor(nameCols.fullName) : [];
  const firstNameRows =
    nameCols.firstName !== undefined ? valuesFor(nameCols.firstName) : [];
  const lastNameRows =
    nameCols.lastName !== undefined ? valuesFor(nameCols.lastName) : [];

  for (let i = 0; i < emailRows.length; i++) {
    const rawEmail = emailRows[i]?.[0];
    if (rawEmail === null || rawEmail === undefined || rawEmail === "")
      continue;
    tabResult.rowsRead++;
    const normalized = normalizeEmail(String(rawEmail));
    if (!normalized) continue;
    tabResult.emailsFound++;

    // Build display name (same logic as sync-engine reader)
    const full = clean(fullNameRows[i]?.[0]);
    const first = clean(firstNameRows[i]?.[0]);
    const last = clean(lastNameRows[i]?.[0]);
    let name = "";
    if (full) {
      name = full;
    } else if (first || last) {
      if (
        first &&
        last &&
        first.toLowerCase() === last.toLowerCase()
      ) {
        name = first;
      } else {
        name = [first, last].filter(Boolean).join(" ");
      }
    }

    // Collect mapped column values
    const perColumn = new Map<string, string>();
    for (const [logical, idx] of logicalColIndex) {
      if (idx < 0) continue;
      const colRows = valuesFor(idx);
      const v = clean(colRows[i]?.[0]);
      if (v) perColumn.set(logical, v);
    }

    // First-seen row wins per email within a single sheet
    if (!byEmail.has(normalized)) {
      byEmail.set(normalized, {
        rowNumber: i + 2,
        name,
        perColumn,
      });
    }
  }

  return {
    result: tabResult,
    sheetId,
    spreadsheetId,
    byEmail,
    logicalColIndex,
    emailColIdx,
  };
}

export async function runSyncProSection(
  refreshToken: string,
  masterSheetId: string,
  section: ProSection
): Promise<ProSectionResult> {
  const masterTabName = section.masterTabName || `Pro: ${section.name}`;
  const result: ProSectionResult = {
    sectionId: section.id,
    sectionName: section.name,
    masterSpreadsheetUrl: "",
    masterTabName,
    totalUniqueEmails: 0,
    totalCellsFilled: 0,
    totalConflicts: 0,
    linkedSheets: [],
    columnStats: section.propagateColumns.map((c) => ({
      name: c.name,
      cellsFilled: 0,
      conflicts: 0,
      skippedSheets: 0,
    })),
    conflicts: [],
    presentInWritten: false,
  };

  if (!section.linkedSheets.length) {
    result.error = "Section needs at least one linked sheet.";
    return result;
  }

  const sheetsClient = getSheetsClient(refreshToken);
  const propagateColumnNames = section.propagateColumns.map((c) => c.name);

  // Read every linked sheet in parallel under MAX_CONCURRENCY
  const readTasks = section.linkedSheets.map(
    (linked) => async () => {
      try {
        return await readLinkedSheet(sheetsClient, linked, propagateColumnNames);
      } catch (err) {
        const tabResult: ProTabResult = {
          nickname: linked.nickname,
          url: linked.url,
          tabName: linked.tabName,
          rowsRead: 0,
          emailsFound: 0,
          cellsFilled: 0,
          error: err instanceof Error ? err.message : "Failed to read sheet",
        };
        return {
          result: tabResult,
          sheetId: -1,
          spreadsheetId: "",
          byEmail: new Map<string, SheetRowRecord>(),
          logicalColIndex: new Map<string, number>(),
          emailColIdx: -1,
        } as SheetReadOutcome;
      }
    }
  );
  const outcomes = await withConcurrencyLimit(readTasks, MAX_CONCURRENCY);
  // Preserve config order in the result list
  outcomes.sort(
    (a, b) =>
      section.linkedSheets.findIndex((l) => l.nickname === a.result.nickname) -
      section.linkedSheets.findIndex((l) => l.nickname === b.result.nickname)
  );

  result.linkedSheets = outcomes.map((o) => o.result);

  // Skipped-sheets count per logical column (sheets that mapped to null or
  // whose header lookup failed end up with idx=-1)
  for (const col of result.columnStats) {
    col.skippedSheets = outcomes.filter(
      (o) => (o.logicalColIndex.get(col.name) ?? -1) < 0
    ).length;
  }

  // Build cross-reference: email → array of {linkedSheetIdx, sheetId, rowNumber}
  const crossRef = new Map<
    string,
    Array<{
      linkedSheetIdx: number;
      sheetId: number;
      rowNumber: number;
      spreadsheetId: string;
    }>
  >();
  for (let i = 0; i < outcomes.length; i++) {
    const o = outcomes[i];
    for (const [email, record] of o.byEmail) {
      const list = crossRef.get(email) ?? [];
      list.push({
        linkedSheetIdx: i,
        sheetId: o.sheetId,
        rowNumber: record.rowNumber,
        spreadsheetId: o.spreadsheetId,
      });
      crossRef.set(email, list);
    }
  }
  result.totalUniqueEmails = crossRef.size;

  // Stash outcomes + crossRef on a local object that subsequent tasks
  // (propagation + present-in + master tab writes) will consume. For now
  // we early-return with read stats only; Tasks 6-7 fill in the rest.
  void crossRef;
  return result;
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/lib/sync-pro-engine.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): engine read phase + cross-reference

Implements readLinkedSheet: resolves the email column (auto-detect
or letter override), resolves each mapped logical propagate column
to a 0-based column index, then bulk-reads email + name + mapped
columns via a single batchGet per sheet.

runSyncProSection now:
- Fans out reads across linked sheets at MAX_CONCURRENCY=3 (matches
  the basic sync engine).
- Collects per-sheet results (rowsRead, emailsFound) into
  result.linkedSheets.
- Tallies skippedSheets per logical column.
- Builds the cross-reference map (email → which sheets contain it).
- Early-returns with read stats only. Propagation, Present In, and
  Master tab writes land in Tasks 6-7.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 6: Engine — propagation pass + write-back

Implements the column-propagation logic and writes filled cells back into each source sheet via `values.batchUpdate`. Per-section result fields (`totalCellsFilled`, `totalConflicts`, `conflicts`, `columnStats[*].cellsFilled`, per-tab `cellsFilled`) get populated.

**Files:**

- Modify: `src/lib/sync-pro-engine.ts`

- [ ] **Step 1: Replace the section runner's tail (the `void crossRef; return result;` block)**

Open `src/lib/sync-pro-engine.ts`, find the lines:

```typescript
  // Stash outcomes + crossRef on a local object that subsequent tasks
  // (propagation + present-in + master tab writes) will consume. For now
  // we early-return with read stats only; Tasks 6-7 fill in the rest.
  void crossRef;
  return result;
}
```

Replace those four lines (plus the trailing `}` that closes `runSyncProSection`) with:

```typescript
  // ----- Propagation pass -----
  //
  // For each logical propagate column, look at every email present in 2+
  // linked sheets. Bucket the per-sheet values into blanks vs non-blanks:
  //
  //   0 non-blank          → nothing to do
  //   1 non-blank          → propagate that value into every blank cell
  //   2+ non-blank, equal  → same (treat as one value)
  //   2+ non-blank, differ → CONFLICT. Don't write anything. Log it.
  //
  // "Equal" is case-insensitive after trimming + collapsing whitespace
  // (same notion of equality we use everywhere else in this codebase).

  function normForCompare(s: string): string {
    return s.toLowerCase().trim().replace(/\s+/g, " ");
  }

  // Per-sheet collected writes: linkedSheetIdx → array of cell writes
  interface PendingCellWrite {
    rowNumber: number; // 1-based
    colIdx: number; // 0-based
    value: string;
  }
  const pendingPerSheet = new Map<number, PendingCellWrite[]>();

  for (const col of section.propagateColumns) {
    const stats = result.columnStats.find((c) => c.name === col.name);
    if (!stats) continue;

    for (const [email, locations] of crossRef) {
      if (locations.length < 2) continue;

      // Collect this column's value from each sheet that actually mapped it
      const perSheetValue: Array<{
        linkedSheetIdx: number;
        sheetIdx: number; // index in outcomes[]
        value: string;
      }> = [];
      for (const loc of locations) {
        const outcome = outcomes[loc.linkedSheetIdx];
        const idx = outcome.logicalColIndex.get(col.name) ?? -1;
        if (idx < 0) continue; // sheet skipped this column
        const rec = outcome.byEmail.get(email);
        const v = rec?.perColumn.get(col.name) ?? "";
        perSheetValue.push({
          linkedSheetIdx: loc.linkedSheetIdx,
          sheetIdx: loc.linkedSheetIdx,
          value: v,
        });
      }

      if (perSheetValue.length < 2) continue;

      const blanks = perSheetValue.filter((p) => !p.value);
      const nonBlanks = perSheetValue.filter((p) => p.value);
      if (nonBlanks.length === 0) continue;
      if (blanks.length === 0) {
        // All sheets that mapped this column already have a value.
        // Conflict-check (logged but never overwritten).
        const distinct = new Set(nonBlanks.map((n) => normForCompare(n.value)));
        if (distinct.size > 1) {
          stats.conflicts++;
          result.totalConflicts++;
          result.conflicts.push({
            email,
            column: col.name,
            values: nonBlanks.map((n) => ({
              nickname:
                section.linkedSheets[n.linkedSheetIdx]?.nickname ?? "?",
              value: n.value,
            })),
          });
        }
        continue;
      }

      const distinctNonBlank = new Set(
        nonBlanks.map((n) => normForCompare(n.value))
      );
      if (distinctNonBlank.size > 1) {
        stats.conflicts++;
        result.totalConflicts++;
        result.conflicts.push({
          email,
          column: col.name,
          values: nonBlanks.map((n) => ({
            nickname: section.linkedSheets[n.linkedSheetIdx]?.nickname ?? "?",
            value: n.value,
          })),
        });
        continue;
      }

      // Single value (or all equal) → propagate into every blank cell.
      const valueToWrite = nonBlanks[0].value;
      for (const b of blanks) {
        const outcome = outcomes[b.linkedSheetIdx];
        const rec = outcome.byEmail.get(email);
        const colIdx = outcome.logicalColIndex.get(col.name) ?? -1;
        if (!rec || colIdx < 0) continue;
        const writes = pendingPerSheet.get(b.linkedSheetIdx) ?? [];
        writes.push({
          rowNumber: rec.rowNumber,
          colIdx,
          value: valueToWrite,
        });
        pendingPerSheet.set(b.linkedSheetIdx, writes);
        stats.cellsFilled++;
        result.totalCellsFilled++;
      }
    }
  }

  // Apply per-sheet writes (one values.batchUpdate per sheet, parallelized
  // under MAX_CONCURRENCY).
  const writeTasks: Array<() => Promise<void>> = [];
  for (const [sheetIdx, writes] of pendingPerSheet) {
    if (!writes.length) continue;
    const outcome = outcomes[sheetIdx];
    const linked = section.linkedSheets[sheetIdx];
    const tabResult = result.linkedSheets[sheetIdx];
    const safeTab = `'${linked.tabName.replace(/'/g, "''")}'`;
    const data = writes.map((w) => ({
      range: `${safeTab}!${columnIndexToLetter(w.colIdx)}${w.rowNumber}`,
      values: [[w.value]],
    }));
    writeTasks.push(async () => {
      try {
        await withRetry(() =>
          sheetsClient.spreadsheets.values.batchUpdate({
            spreadsheetId: outcome.spreadsheetId,
            requestBody: { valueInputOption: "USER_ENTERED", data },
          })
        );
        tabResult.cellsFilled = writes.length;
      } catch (err) {
        tabResult.error = `${tabResult.error ? tabResult.error + "; " : ""}Failed to write propagated cells: ${err instanceof Error ? err.message : String(err)}`;
      }
    });
  }
  await withConcurrencyLimit(writeTasks, MAX_CONCURRENCY);

  // Present In + Master tab writes land in Task 7.
  return result;
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/lib/sync-pro-engine.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): column propagation + write-back

Implements the propagation pass per the spec:
- For each logical propagate column, scan emails present in 2+
  linked sheets, bucket values into blanks vs non-blanks.
- 0 non-blank → no-op.
- 1 non-blank (or 2+ non-blank that are all equal after normalize)
  → propagate to every blank cell.
- 2+ non-blank with distinct values → CONFLICT. Don't write.
  Logged to result.conflicts with the per-sheet value list.

Pending writes are batched per source sheet and applied via
values.batchUpdate (one round trip per sheet) with valueInputOption
USER_ENTERED. Parallelized under MAX_CONCURRENCY=3. Per-sheet write
failures are captured on the tab result, not fatal to the section.

Stats updated: totalCellsFilled, totalConflicts, per-column
cellsFilled / conflicts, per-tab cellsFilled.

Present In + Master tab writes still pending (Task 7).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 7: Engine — Present In write-back + Master tab

Final engine task. Calls the existing `writePresentInColumn` helper into each source sheet (skipped if `section.writePresentIn === false`), then writes the Master ✅/❌ tab into the user's master sheet at `section.masterTabName`.

**Files:**

- Modify: `src/lib/sync-pro-engine.ts`

- [ ] **Step 1: Replace the closing block of `runSyncProSection`**

Find the lines:

```typescript
  await withConcurrencyLimit(writeTasks, MAX_CONCURRENCY);

  // Present In + Master tab writes land in Task 7.
  return result;
}
```

Replace those four lines with:

```typescript
  await withConcurrencyLimit(writeTasks, MAX_CONCURRENCY);

  // ----- Present In write-back (same shape as basic sync) -----
  if (section.writePresentIn) {
    const presentInTasks: Array<() => Promise<void>> = [];
    for (let i = 0; i < outcomes.length; i++) {
      const outcome = outcomes[i];
      if (outcome.sheetId < 0) continue;
      const linked = section.linkedSheets[i];
      const tabResult = result.linkedSheets[i];

      // Re-read this sheet's headers + presentIn column index (we need the
      // last filled column index and any existing "Present In" header).
      const safeTab = `'${linked.tabName.replace(/'/g, "''")}'`;
      let presentInColumnIndex: number | null = null;
      let lastColumnIndex = -1;
      try {
        const headerRes = await withRetry(() =>
          sheetsClient.spreadsheets.values.get({
            spreadsheetId: outcome.spreadsheetId,
            range: `${safeTab}!1:1`,
          })
        );
        const headers = (headerRes.data.values?.[0] ?? []).map((h) =>
          clean(h)
        );
        for (let c = 0; c < headers.length; c++) {
          if (headers[c]) lastColumnIndex = c;
          if (headers[c]?.toLowerCase().trim() === "present in") {
            presentInColumnIndex = c;
          }
        }
      } catch (err) {
        tabResult.error = `${tabResult.error ? tabResult.error + "; " : ""}Failed to read headers for Present In: ${err instanceof Error ? err.message : String(err)}`;
        continue;
      }

      // Build cell data: every row in this sheet that has cross-refs
      const cellData: Array<{
        rowIndex: number; // 0-based
        links: Array<{ text: string; url: string }>;
      }> = [];
      for (const [email, rec] of outcome.byEmail) {
        const others = (crossRef.get(email) ?? []).filter(
          (x) => x.linkedSheetIdx !== i
        );
        if (others.length === 0) continue;
        cellData.push({
          rowIndex: rec.rowNumber - 1,
          links: others.map((o) => ({
            text: section.linkedSheets[o.linkedSheetIdx]?.nickname ?? "?",
            url: `https://docs.google.com/spreadsheets/d/${o.spreadsheetId}/edit#gid=${o.sheetId}&range=A${o.rowNumber}`,
          })),
        });
      }

      const colIdx =
        presentInColumnIndex !== null
          ? presentInColumnIndex
          : lastColumnIndex + 1;
      const headerNeeded = presentInColumnIndex === null;

      presentInTasks.push(async () => {
        try {
          await writePresentInColumn(
            refreshToken,
            outcome.spreadsheetId,
            outcome.sheetId,
            colIdx,
            headerNeeded,
            cellData
          );
        } catch (err) {
          tabResult.error = `${tabResult.error ? tabResult.error + "; " : ""}Failed to write Present In: ${err instanceof Error ? err.message : String(err)}`;
        }
      });
    }
    await withConcurrencyLimit(presentInTasks, MAX_CONCURRENCY);
    result.presentInWritten = true;
  }

  // ----- Master tab write -----
  try {
    const masterMeta = await withRetry(() =>
      sheetsClient.spreadsheets.get({ spreadsheetId: masterSheetId })
    );
    const allMasterTabs = masterMeta.data.sheets ?? [];
    const existing = allMasterTabs.find(
      (t) => t.properties?.title === masterTabName
    );
    let masterSheetIdNumber: number;
    if (
      existing?.properties?.sheetId !== undefined &&
      existing.properties.sheetId !== null
    ) {
      masterSheetIdNumber = existing.properties.sheetId;
    } else {
      const addRes = await withRetry(() =>
        sheetsClient.spreadsheets.batchUpdate({
          spreadsheetId: masterSheetId,
          requestBody: {
            requests: [
              {
                addSheet: { properties: { title: masterTabName } },
              },
            ],
          },
        })
      );
      masterSheetIdNumber =
        addRes.data.replies?.[0]?.addSheet?.properties?.sheetId ?? -1;
      if (masterSheetIdNumber < 0) {
        throw new Error(`Could not create master tab "${masterTabName}"`);
      }
    }

    const safeMaster = `'${masterTabName.replace(/'/g, "''")}'`;
    await withRetry(() =>
      sheetsClient.spreadsheets.values.clear({
        spreadsheetId: masterSheetId,
        range: `${safeMaster}!A:ZZ`,
      })
    );

    // Aggregate names per email across all sheets for prefix-folding merge
    const namesByEmail = new Map<string, string[]>();
    for (const o of outcomes) {
      for (const [email, rec] of o.byEmail) {
        if (!rec.name) continue;
        const list = namesByEmail.get(email) ?? [];
        list.push(rec.name);
        namesByEmail.set(email, list);
      }
    }
    const mergedNames = mergeNamesForEmail(namesByEmail);

    // Build Master tab values: header + one row per unique email
    const sortedEmails = Array.from(crossRef.keys()).sort();
    const headerRow = [
      "Name",
      "Email",
      ...section.linkedSheets.map((l) => l.nickname || l.tabName),
    ];
    const dataRows = sortedEmails.map((email) => {
      const presence = section.linkedSheets.map((_, idx) =>
        outcomes[idx]?.byEmail.has(email) ? "✅" : "❌"
      );
      return [mergedNames.get(email) ?? "", email, ...presence];
    });
    const valueRows = [headerRow, ...dataRows];
    const lastColLetter = columnIndexToLetter(headerRow.length - 1);

    await withRetry(() =>
      sheetsClient.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${safeMaster}!A1:${lastColLetter}${valueRows.length}`,
        valueInputOption: "RAW",
        requestBody: { values: valueRows },
      })
    );

    // Best-effort header styling (don't fail the section if styling errors).
    try {
      await withRetry(() =>
        sheetsClient.spreadsheets.batchUpdate({
          spreadsheetId: masterSheetId,
          requestBody: {
            requests: [
              {
                repeatCell: {
                  range: {
                    sheetId: masterSheetIdNumber,
                    startRowIndex: 0,
                    endRowIndex: 1,
                    startColumnIndex: 0,
                    endColumnIndex: headerRow.length,
                  },
                  cell: {
                    userEnteredFormat: {
                      backgroundColor: HEADER_BG,
                      textFormat: { bold: true },
                    },
                  },
                  fields:
                    "userEnteredFormat(backgroundColor,textFormat)",
                },
              },
              {
                updateSheetProperties: {
                  properties: {
                    sheetId: masterSheetIdNumber,
                    gridProperties: { frozenRowCount: 1 },
                  },
                  fields: "gridProperties.frozenRowCount",
                },
              },
            ],
          },
        })
      );
    } catch {
      // ignore styling failures
    }

    result.masterSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${masterSheetId}/edit#gid=${masterSheetIdNumber}`;
  } catch (err) {
    result.error = `${result.error ? result.error + "; " : ""}Failed to write Master tab: ${err instanceof Error ? err.message : String(err)}`;
  }

  return result;
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/lib/sync-pro-engine.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): Present In write-back + Master tab

Completes runSyncProSection.

Present In:
- Skipped entirely when section.writePresentIn === false.
- Per linked sheet: re-reads headers to detect existing "Present In"
  column or compute the append index, then builds per-row hyperlink
  cells via writePresentInColumn (existing helper). Deep links use
  spreadsheet/sheet/range fragments — same convention as basic sync.

Master tab:
- ensureTab (lazy create) at section.masterTabName under the user's
  master spreadsheet.
- Clears + rewrites in full each run.
- Header [Name, Email, <nickname1>, ...]. Presence ✅/❌ per linked
  sheet. Name comes from the duplicated prefix-folding merge.
- Best-effort header bold + bg + frozen first row (styling failures
  are swallowed; data write still counts as success).

Result returns masterSpreadsheetUrl deep-linked to the section's
master tab. Per-tab + per-section errors are captured non-fatally so
one bad sheet doesn't take down the section.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 8: POST API route (`/api/sync-pro`)

Handles both single-section runs (UI per-section button) and multi-section batches (UI Run all). Same shape as `/api/consolidator` — it always sends `sections: []`.

**Files:**

- Create: `src/app/api/sync-pro/route.ts`

- [ ] **Step 1: Create the route file**

```typescript
// src/app/api/sync-pro/route.ts
import { NextResponse } from "next/server";
import { z } from "zod";
import { getSession } from "@/lib/session";
import { runSyncProBatch } from "@/lib/sync-pro-engine";

const LinkedSheetSchema = z.object({
  url: z
    .string()
    .url()
    .refine((url) => url.includes("docs.google.com/spreadsheets"), {
      message: "Linked sheet URL must be a Google Sheets URL",
    }),
  nickname: z.string().min(1, "Linked sheet nickname is required"),
  tabName: z.string().min(1, "Linked sheet tab name is required"),
  emailColumn: z.string(),
  columnMapping: z.record(z.string(), z.union([z.string(), z.null()])),
});

const SectionSchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1, "Section name is required"),
  masterTabName: z.string().min(1, "Master tab name is required"),
  writePresentIn: z.boolean(),
  propagateColumns: z.array(z.object({ name: z.string().min(1) })),
  linkedSheets: z
    .array(LinkedSheetSchema)
    .min(1, "Section needs at least one linked sheet"),
});

const Schema = z.object({
  sections: z.array(SectionSchema).min(1, "Provide at least one section"),
});

export async function POST(request: Request) {
  try {
    const session = await getSession();
    if (!session.refreshToken) {
      return NextResponse.json(
        { error: "Not authenticated. Please sign in with Google first." },
        { status: 401 }
      );
    }
    if (!session.masterSheetId) {
      return NextResponse.json(
        {
          error:
            "No master sheet configured. Set one on the Sheets page first.",
        },
        { status: 400 }
      );
    }

    const body = await request.json();
    const parsed = Schema.parse(body);

    const result = await runSyncProBatch(
      session.refreshToken,
      session.masterSheetId,
      parsed.sections
    );
    return NextResponse.json(result);
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
      {
        error: err instanceof Error ? err.message : "Failed to run Sync Pro",
      },
      { status: 500 }
    );
  }
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/app/api/sync-pro/route.ts
git commit -m "$(cat <<'EOF'
feat(sync-pro): POST API route

Accepts the same { sections: [...] } shape used by /api/consolidator
so the UI's per-section run sends sections:[oneSection] and Run all
sends the full list. Zod-validated. Calls runSyncProBatch and
returns the ProBatchResult.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 9: Nav link

Adds the "Sync Pro" entry to the global nav so the page is reachable as soon as it lands.

**Files:**

- Modify: `src/components/nav.tsx`

- [ ] **Step 1: Update the navLinks array**

Find this block in `src/components/nav.tsx`:

```typescript
const navLinks = [
  { href: "/", label: "Dashboard" },
  { href: "/sheets", label: "Sheets" },
  { href: "/consolidator", label: "Consolidator" },
  { href: "/duplicate-finder", label: "Duplicate Finder" },
  { href: "/domain-analyzer", label: "Domain Analyzer" },
  { href: "/email-finder", label: "Email Finder" },
  { href: "/renewal-sync", label: "Renewal Sync" },
  { href: "/biz-tutor-sync", label: "BIZ Tutor" },
  { href: "/report-sync", label: "Report Sync" },
];
```

Replace it with:

```typescript
const navLinks = [
  { href: "/", label: "Dashboard" },
  { href: "/sheets", label: "Sheets" },
  { href: "/sync-pro", label: "Sync Pro" },
  { href: "/consolidator", label: "Consolidator" },
  { href: "/duplicate-finder", label: "Duplicate Finder" },
  { href: "/domain-analyzer", label: "Domain Analyzer" },
  { href: "/email-finder", label: "Email Finder" },
  { href: "/renewal-sync", label: "Renewal Sync" },
  { href: "/biz-tutor-sync", label: "BIZ Tutor" },
  { href: "/report-sync", label: "Report Sync" },
];
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

(The page route is added in the next task — clicking the link before Task 10 lands will 404, which is fine.)

- [ ] **Step 3: Commit**

```bash
git add src/components/nav.tsx
git commit -m "$(cat <<'EOF'
feat(sync-pro): add Sync Pro nav link

Placed immediately after Sheets so it sits next to the basic sync
config screen in the menu order.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 10: Page UI scaffold + state model + hydrate / persist

Bootstraps the new page with the per-section state model, hydration from `_sync_pro_config` on mount, and debounced auto-save. No section card yet — placeholder render. Section card UI lands in Task 11.

**Files:**

- Create: `src/app/sync-pro/page.tsx`

- [ ] **Step 1: Create the page file**

```typescript
// src/app/sync-pro/page.tsx
"use client";

import { useCallback, useEffect, useState } from "react";
import type {
  ProConflict,
  ProColumnStats,
  ProLinkedSheet,
  ProPropagateColumn,
  ProSection,
  ProSectionResult,
  ProTabResult,
} from "@/lib/sync-pro-types";

// ---------- Local UI state shapes ----------

interface SourceTabState {
  loadedTabs: string[]; // tabs fetched from Google for the URL
  loadedHeaders: string[]; // headers for the picked tab (after Fetch tabs)
  loading: boolean;
}

interface LinkedSheetState extends ProLinkedSheet {
  ui: SourceTabState;
}

interface SectionState
  extends Omit<ProSection, "linkedSheets"> {
  linkedSheets: LinkedSheetState[];
}

function newId(): string {
  return `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function emptyLinkedSheet(): LinkedSheetState {
  return {
    url: "",
    nickname: "",
    tabName: "",
    emailColumn: "auto",
    columnMapping: {},
    ui: { loadedTabs: [], loadedHeaders: [], loading: false },
  };
}

function makeSection(index: number): SectionState {
  const name = `Section ${index + 1}`;
  return {
    id: newId(),
    name,
    masterTabName: `Pro: ${name}`,
    writePresentIn: true,
    propagateColumns: [],
    linkedSheets: [emptyLinkedSheet()],
  };
}

function toConfigSection(s: SectionState): ProSection {
  return {
    id: s.id,
    name: s.name,
    masterTabName: s.masterTabName,
    writePresentIn: s.writePresentIn,
    propagateColumns: s.propagateColumns,
    linkedSheets: s.linkedSheets.map(({ ui: _ui, ...rest }) => {
      void _ui;
      return rest;
    }),
  };
}

async function fetchTabsForUrl(url: string): Promise<string[]> {
  const res = await fetch(`/api/sheets/tabs?url=${encodeURIComponent(url)}`);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to fetch tabs");
  return data.tabs ?? [];
}

async function fetchHeadersForUrlTab(
  url: string,
  tabName: string
): Promise<string[]> {
  // Reuses /api/sheets/tabs? — not quite, headers aren't exposed by any
  // existing route. We add a small client-side fallback: fetch headers via
  // /api/sheets/headers when added; until then keep loadedHeaders empty
  // and let the UI fall back to a free-text mapping input.
  void url;
  void tabName;
  return [];
}

export default function SyncProPage() {
  const [sections, setSections] = useState<SectionState[]>([makeSection(0)]);
  const [results, setResults] = useState<Map<string, ProSectionResult>>(
    new Map()
  );
  const [errors, setErrors] = useState<Map<string, string>>(new Map());
  const [runningSections, setRunningSections] = useState<Set<string>>(
    new Set()
  );

  const [hydrated, setHydrated] = useState(false);
  const [hasSaved, setHasSaved] = useState(false);
  const [noMasterSheet, setNoMasterSheet] = useState(false);

  // ---------- Per-section state helpers ----------

  function setSectionRunning(id: string, on: boolean) {
    setRunningSections((prev) => {
      const next = new Set(prev);
      if (on) next.add(id);
      else next.delete(id);
      return next;
    });
  }

  function setSectionError(id: string, msg: string | null) {
    setErrors((prev) => {
      const next = new Map(prev);
      if (msg) next.set(id, msg);
      else next.delete(id);
      return next;
    });
  }

  function setSectionResult(id: string, result: ProSectionResult | null) {
    setResults((prev) => {
      const next = new Map(prev);
      if (result) next.set(id, result);
      else next.delete(id);
      return next;
    });
  }

  // ---------- Hydrate from server on mount ----------

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/sync-pro/config");
        if (res.status === 400) {
          const data = await res.json();
          if (data.code === "no_master_sheet") setNoMasterSheet(true);
        } else if (res.ok) {
          const data = (await res.json()) as { sections: ProSection[] };
          if (data.sections && data.sections.length > 0) {
            setHasSaved(true);
            const rehydrated: SectionState[] = data.sections.map(
              (s, idx): SectionState => ({
                id: s.id || newId(),
                name: s.name || `Section ${idx + 1}`,
                masterTabName:
                  s.masterTabName || `Pro: ${s.name || `Section ${idx + 1}`}`,
                writePresentIn:
                  typeof s.writePresentIn === "boolean"
                    ? s.writePresentIn
                    : true,
                propagateColumns: s.propagateColumns ?? [],
                linkedSheets:
                  s.linkedSheets && s.linkedSheets.length > 0
                    ? s.linkedSheets.map((l) => ({
                        ...l,
                        ui: {
                          loadedTabs: [],
                          loadedHeaders: [],
                          loading: false,
                        },
                      }))
                    : [emptyLinkedSheet()],
              })
            );
            setSections(rehydrated);

            // Eagerly load tab lists for any linked sheet that already has a URL
            await Promise.all(
              rehydrated.map(async (sec, sIdx) => {
                await Promise.all(
                  sec.linkedSheets.map(async (linked, lIdx) => {
                    if (!linked.url) return;
                    try {
                      const tabs = await fetchTabsForUrl(linked.url);
                      setSections((prev) => {
                        const next = prev.slice();
                        const sec2 = { ...next[sIdx] };
                        const ls = sec2.linkedSheets.slice();
                        ls[lIdx] = {
                          ...ls[lIdx],
                          ui: {
                            ...ls[lIdx].ui,
                            loadedTabs: tabs,
                            loading: false,
                          },
                        };
                        sec2.linkedSheets = ls;
                        next[sIdx] = sec2;
                        return next;
                      });
                    } catch {
                      // ignore — user can refetch
                    }
                  })
                );
              })
            );
          }
        }
      } catch {
        // empty form
      }
      setHydrated(true);
    })();
  }, []);

  // ---------- Persist to server on changes (debounced) ----------

  useEffect(() => {
    if (!hydrated || noMasterSheet) return;
    const hasAnything = sections.some(
      (s) =>
        s.name ||
        s.linkedSheets.some((l) => l.url || l.nickname) ||
        s.propagateColumns.length > 0
    );
    if (!hasAnything) return;

    const handle = setTimeout(() => {
      const payload = { sections: sections.map(toConfigSection) };
      fetch("/api/sync-pro/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      })
        .then((r) => {
          if (r.ok) setHasSaved(true);
        })
        .catch(() => {
          // ignore
        });
    }, 800);
    return () => clearTimeout(handle);
  }, [sections, hydrated, noMasterSheet]);

  // ---------- Mutators ----------

  const updateSection = useCallback(
    (idx: number, patch: Partial<SectionState>) => {
      setSections((prev) => {
        const next = prev.slice();
        next[idx] = { ...next[idx], ...patch };
        return next;
      });
    },
    []
  );

  function addSection() {
    setSections((prev) => [...prev, makeSection(prev.length)]);
  }

  function removeSection(idx: number) {
    setSections((prev) => {
      if (prev.length === 1) return [makeSection(0)];
      const removed = prev[idx];
      if (removed) {
        setSectionResult(removed.id, null);
        setSectionError(removed.id, null);
        setSectionRunning(removed.id, false);
      }
      return prev.filter((_, i) => i !== idx);
    });
  }

  async function handleClearSaved() {
    setErrors(new Map());
    setResults(new Map());
    try {
      await fetch("/api/sync-pro/config", { method: "DELETE" });
    } catch {
      // best-effort
    }
    setSections([makeSection(0)]);
    setHasSaved(false);
  }

  // Stubs filled in by Task 12 / 13.
  function handleRunSection(_idx: number) {
    void _idx;
  }
  async function handleRunAll(e: React.FormEvent) {
    e.preventDefault();
  }

  // Silence "unused" warnings on imports referenced only inside future tasks
  void updateSection;
  void addSection;
  void removeSection;
  void handleRunSection;
  void handleClearSaved;
  void fetchHeadersForUrlTab;
  void ({} as ProConflict);
  void ({} as ProColumnStats);
  void ({} as ProPropagateColumn);
  void ({} as ProTabResult);

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Sync Pro
        </h1>
        <p className="text-muted text-sm mt-1">
          Per-section, multi-sheet Present In with column propagation.
          Sections are independent — each runs on its own button. (UI under
          construction.)
        </p>
      </div>

      {noMasterSheet && (
        <div className="bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          Set a master sheet on the Sheets page to use Sync Pro.
        </div>
      )}

      {hasSaved && !noMasterSheet && (
        <div className="flex items-center justify-between gap-3 bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          <span>
            Configuration saved to your master sheet — it&apos;ll auto-load
            next time.
          </span>
          <button
            type="button"
            onClick={handleClearSaved}
            className="text-danger hover:underline font-medium cursor-pointer whitespace-nowrap"
          >
            Clear saved
          </button>
        </div>
      )}

      <form onSubmit={handleRunAll} className="space-y-5">
        {sections.map((s, idx) => (
          <div
            key={s.id}
            className="bg-card border border-border rounded-lg p-5 text-sm text-muted"
          >
            <span className="font-medium text-foreground">{s.name}</span>{" "}
            (section {idx + 1} placeholder — full UI lands in Task 11)
          </div>
        ))}

        <div className="flex items-center justify-between gap-3 flex-wrap">
          <button
            type="button"
            onClick={addSection}
            className="text-sm font-medium px-3 py-1.5 rounded-md border border-border bg-background hover:bg-card transition-colors cursor-pointer"
          >
            + Add section
          </button>
          {sections.length > 1 && (
            <button
              type="submit"
              disabled={runningSections.size > 0}
              className="bg-primary text-white px-5 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
            >
              {runningSections.size > 0
                ? "Running..."
                : `Run all ${sections.length} sections`}
            </button>
          )}
        </div>
      </form>
    </div>
  );
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/app/sync-pro/page.tsx
git commit -m "$(cat <<'EOF'
feat(sync-pro): page scaffold + state model + hydrate/persist

Bootstraps the Sync Pro page with the per-section state model
(SectionState wraps ProSection with per-linked-sheet UI fields like
loadedTabs / loading), debounced auto-save to _sync_pro_config, and
hydration on mount (also eagerly fetches the tab list for any linked
sheet that already has a URL so the UI is interactive immediately).

Per-section runtime state (results / errors / runningSections) is
keyed by section.id — sections are independent. Section add/remove
auto-cleans the removed section's runtime state. handleRunSection /
handleRunAll are stubbed; SectionCard, mapping UI, and the Result
panel land in Tasks 11-13.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 11: Page UI — SectionCard + LinkedSheetRow

Replaces the placeholder render with a real `SectionCard` component containing per-section settings (name, master tab name, Present In toggle) plus a list of `LinkedSheetRow` subcomponents (URL, fetch tabs, tab dropdown, email column override, remove sheet).

**Files:**

- Modify: `src/app/sync-pro/page.tsx`

- [ ] **Step 1: Replace the placeholder map block and append the two components**

In `src/app/sync-pro/page.tsx`, find the block:

```typescript
        {sections.map((s, idx) => (
          <div
            key={s.id}
            className="bg-card border border-border rounded-lg p-5 text-sm text-muted"
          >
            <span className="font-medium text-foreground">{s.name}</span>{" "}
            (section {idx + 1} placeholder — full UI lands in Task 11)
          </div>
        ))}
```

Replace it with:

```typescript
        {sections.map((section, sIdx) => (
          <SectionCard
            key={section.id}
            section={section}
            index={sIdx}
            total={sections.length}
            running={runningSections.has(section.id)}
            error={errors.get(section.id) ?? null}
            result={results.get(section.id) ?? null}
            onChangeSection={(patch) => updateSection(sIdx, patch)}
            onRemove={() => removeSection(sIdx)}
            onAddLinkedSheet={() =>
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                sec.linkedSheets = [
                  ...sec.linkedSheets,
                  emptyLinkedSheet(),
                ];
                next[sIdx] = sec;
                return next;
              })
            }
            onRemoveLinkedSheet={(lIdx) =>
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                sec.linkedSheets =
                  sec.linkedSheets.length === 1
                    ? [emptyLinkedSheet()]
                    : sec.linkedSheets.filter((_, i) => i !== lIdx);
                next[sIdx] = sec;
                return next;
              })
            }
            onChangeLinkedSheet={(lIdx, patch) =>
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                const ls = sec.linkedSheets.slice();
                ls[lIdx] = { ...ls[lIdx], ...patch };
                sec.linkedSheets = ls;
                next[sIdx] = sec;
                return next;
              })
            }
            onFetchTabs={async (lIdx) => {
              const linked = sections[sIdx]?.linkedSheets[lIdx];
              if (!linked?.url) return;
              setSectionError(section.id, null);
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                const ls = sec.linkedSheets.slice();
                ls[lIdx] = {
                  ...ls[lIdx],
                  ui: { ...ls[lIdx].ui, loading: true },
                };
                sec.linkedSheets = ls;
                next[sIdx] = sec;
                return next;
              });
              try {
                const tabs = await fetchTabsForUrl(linked.url);
                setSections((prev) => {
                  const next = prev.slice();
                  const sec = { ...next[sIdx] };
                  const ls = sec.linkedSheets.slice();
                  ls[lIdx] = {
                    ...ls[lIdx],
                    ui: {
                      ...ls[lIdx].ui,
                      loadedTabs: tabs,
                      loading: false,
                    },
                  };
                  sec.linkedSheets = ls;
                  next[sIdx] = sec;
                  return next;
                });
              } catch (err) {
                setSectionError(
                  section.id,
                  err instanceof Error
                    ? err.message
                    : "Failed to fetch tabs"
                );
                setSections((prev) => {
                  const next = prev.slice();
                  const sec = { ...next[sIdx] };
                  const ls = sec.linkedSheets.slice();
                  ls[lIdx] = {
                    ...ls[lIdx],
                    ui: { ...ls[lIdx].ui, loading: false },
                  };
                  sec.linkedSheets = ls;
                  next[sIdx] = sec;
                  return next;
                });
              }
            }}
            onRunSection={() => handleRunSection(sIdx)}
          />
        ))}
```

Then append these two components at the bottom of the file (after the default export's closing brace):

```typescript
// ---------- SectionCard ----------

function SectionCard({
  section,
  index,
  total,
  running,
  error,
  result,
  onChangeSection,
  onRemove,
  onAddLinkedSheet,
  onRemoveLinkedSheet,
  onChangeLinkedSheet,
  onFetchTabs,
  onRunSection,
}: {
  section: SectionState;
  index: number;
  total: number;
  running: boolean;
  error: string | null;
  result: ProSectionResult | null;
  onChangeSection: (patch: Partial<SectionState>) => void;
  onRemove: () => void;
  onAddLinkedSheet: () => void;
  onRemoveLinkedSheet: (linkedIdx: number) => void;
  onChangeLinkedSheet: (
    linkedIdx: number,
    patch: Partial<LinkedSheetState>
  ) => void;
  onFetchTabs: (linkedIdx: number) => void;
  onRunSection: () => void;
}) {
  // When the user edits the section name, auto-update the master tab name
  // unless the user has already manually edited it away from the default.
  const defaultMasterTabName = `Pro: ${section.name}`;
  return (
    <div className="bg-card border border-border rounded-lg p-5 space-y-5">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div className="flex items-center gap-2 flex-1 min-w-0">
          <span className="text-xs text-muted uppercase tracking-wide">
            Section {index + 1}
          </span>
          <input
            type="text"
            value={section.name}
            onChange={(e) => {
              const newName = e.target.value;
              const masterTabName =
                section.masterTabName === defaultMasterTabName
                  ? `Pro: ${newName}`
                  : section.masterTabName;
              onChangeSection({ name: newName, masterTabName });
            }}
            placeholder="Section name"
            className="flex-1 min-w-0 border border-border rounded-md px-2 py-1 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </div>
        {total > 1 && (
          <button
            type="button"
            onClick={onRemove}
            className="text-xs text-danger hover:underline cursor-pointer whitespace-nowrap"
          >
            Remove section
          </button>
        )}
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <label className="flex flex-col gap-1 text-xs">
          <span className="text-muted uppercase tracking-wide">
            Master tab name
          </span>
          <input
            type="text"
            value={section.masterTabName}
            onChange={(e) =>
              onChangeSection({ masterTabName: e.target.value })
            }
            placeholder="Pro: Section 1"
            className="border border-border rounded-md px-2 py-1 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </label>
        <label className="flex items-center gap-2 text-sm self-end">
          <input
            type="checkbox"
            checked={section.writePresentIn}
            onChange={(e) =>
              onChangeSection({ writePresentIn: e.target.checked })
            }
          />
          Also write Present In column into each source
        </label>
      </div>

      {/* Linked sheets list */}
      <div className="space-y-3">
        <p className="text-xs font-medium text-muted uppercase tracking-wide">
          Linked sheets ({section.linkedSheets.length})
        </p>
        {section.linkedSheets.map((linked, lIdx) => (
          <LinkedSheetRow
            key={lIdx}
            linked={linked}
            index={lIdx}
            total={section.linkedSheets.length}
            propagateColumns={section.propagateColumns}
            onChange={(patch) => onChangeLinkedSheet(lIdx, patch)}
            onRemove={() => onRemoveLinkedSheet(lIdx)}
            onFetch={() => onFetchTabs(lIdx)}
          />
        ))}
        <button
          type="button"
          onClick={onAddLinkedSheet}
          className="text-xs font-medium px-3 py-1.5 rounded-md border border-dashed border-border bg-background hover:bg-card transition-colors cursor-pointer"
        >
          + Add another linked sheet
        </button>
      </div>

      {/* Per-section Run + inline error */}
      <div className="border-t border-border pt-4 flex items-center justify-between gap-3 flex-wrap">
        {error ? (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2 flex-1 min-w-0">
            {error}
          </p>
        ) : (
          <span className="text-xs text-muted">
            Runs this section only. Other sections are untouched.
          </span>
        )}
        <button
          type="button"
          onClick={onRunSection}
          disabled={running}
          className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap"
        >
          {running ? "Running..." : "Run section"}
        </button>
      </div>

      {/* Result panel — filled in Task 13. */}
      {result && (
        <div className="border-t border-border pt-4 text-xs text-muted">
          Last run produced a result (rendering UI lands in Task 13).
          Cells filled: {result.totalCellsFilled} · Conflicts:{" "}
          {result.totalConflicts}
        </div>
      )}
    </div>
  );
}

// ---------- LinkedSheetRow ----------

function LinkedSheetRow({
  linked,
  index,
  total,
  propagateColumns: _propagateColumns,
  onChange,
  onRemove,
  onFetch,
}: {
  linked: LinkedSheetState;
  index: number;
  total: number;
  propagateColumns: ProPropagateColumn[];
  onChange: (patch: Partial<LinkedSheetState>) => void;
  onRemove: () => void;
  onFetch: () => void;
}) {
  void _propagateColumns; // mapping UI lands in Task 12
  return (
    <div className="border border-border rounded-md p-3 space-y-3 bg-background/50">
      <div className="flex items-center justify-between gap-2">
        <span className="text-xs text-muted">Sheet #{index + 1}</span>
        {total > 1 && (
          <button
            type="button"
            onClick={onRemove}
            className="text-xs text-danger hover:underline cursor-pointer"
          >
            Remove sheet
          </button>
        )}
      </div>

      <input
        type="text"
        value={linked.nickname}
        onChange={(e) => onChange({ nickname: e.target.value })}
        placeholder="Nickname (e.g. Newsletter)"
        className="w-full border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
      />

      <div className="flex flex-col sm:flex-row gap-2">
        <input
          type="url"
          value={linked.url}
          onChange={(e) =>
            onChange({
              url: e.target.value,
              tabName: "",
              columnMapping: {},
              ui: { loadedTabs: [], loadedHeaders: [], loading: false },
            })
          }
          placeholder="https://docs.google.com/spreadsheets/d/..."
          className="flex-1 min-w-0 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
        <button
          type="button"
          onClick={onFetch}
          disabled={!linked.url || linked.ui.loading}
          className="bg-foreground text-background px-3 py-2 rounded-md text-sm font-medium hover:opacity-80 disabled:opacity-50 transition-opacity cursor-pointer whitespace-nowrap"
        >
          {linked.ui.loading ? "Loading..." : "Fetch tabs"}
        </button>
      </div>

      {linked.ui.loadedTabs.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <label className="flex flex-col gap-1 text-xs">
            <span className="text-muted uppercase tracking-wide">Tab</span>
            <select
              value={linked.tabName}
              onChange={(e) =>
                onChange({ tabName: e.target.value, columnMapping: {} })
              }
              className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            >
              <option value="">— select —</option>
              {linked.ui.loadedTabs.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </label>
          <label className="flex flex-col gap-1 text-xs">
            <span className="text-muted uppercase tracking-wide">
              Email column
            </span>
            <input
              type="text"
              value={linked.emailColumn}
              onChange={(e) =>
                onChange({ emailColumn: e.target.value || "auto" })
              }
              placeholder="auto"
              className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            />
          </label>
        </div>
      )}
    </div>
  );
}
```

- [ ] **Step 2: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/app/sync-pro/page.tsx
git commit -m "$(cat <<'EOF'
feat(sync-pro): SectionCard + LinkedSheetRow components

SectionCard has the section name + master tab name + Present-In
checkbox + list of linked sheets + per-section Run button + inline
error slot. Editing the section name auto-updates the master tab
name when the user hasn't manually edited it away from the default
("Pro: <name>").

LinkedSheetRow has nickname + URL + Fetch tabs + Tab dropdown
(populated after Fetch tabs lands) + email column override (default
"auto"). The per-sheet column mapping UI for propagate columns is
deliberately stubbed and lands in Task 12 once the propagate-columns
chip row exists at the section level.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 12: Page UI — Propagate columns row + per-sheet column mapping

Adds a propagate-columns chip row at the section level (add/remove logical columns) and renders per-sheet mapping inputs inside each `LinkedSheetRow`. Mapping uses a free-text input per logical column for the actual header (since there's no API endpoint that returns a tab's headers and we don't want to add one for this).

**Files:**

- Modify: `src/app/sync-pro/page.tsx`

- [ ] **Step 1: Add the propagate-columns row inside `SectionCard`**

In `src/app/sync-pro/page.tsx`, find the `SectionCard` function. After the `<div>` block that contains the master-tab-name + write-PresentIn checkbox (the one with class `"grid grid-cols-1 sm:grid-cols-2 gap-3"`), insert this new block:

```typescript
      {/* Propagate columns */}
      <div className="space-y-2">
        <p className="text-xs font-medium text-muted uppercase tracking-wide">
          Propagate columns ({section.propagateColumns.length})
        </p>
        <div className="flex flex-wrap gap-2 items-center">
          {section.propagateColumns.map((col, cIdx) => (
            <span
              key={`${col.name}-${cIdx}`}
              className="inline-flex items-center gap-1 bg-background border border-border rounded-md px-2 py-1 text-sm"
            >
              <code>{col.name}</code>
              <button
                type="button"
                onClick={() => {
                  const removedName = col.name;
                  const nextCols = section.propagateColumns.filter(
                    (_, i) => i !== cIdx
                  );
                  const nextLinked = section.linkedSheets.map((l) => {
                    const { [removedName]: _removed, ...rest } =
                      l.columnMapping;
                    void _removed;
                    return { ...l, columnMapping: rest };
                  });
                  onChangeSection({
                    propagateColumns: nextCols,
                    linkedSheets: nextLinked,
                  });
                }}
                className="text-danger hover:underline cursor-pointer"
                aria-label={`Remove ${col.name}`}
              >
                ×
              </button>
            </span>
          ))}
          <AddPropagateColumn
            onAdd={(name) => {
              const trimmed = name.trim();
              if (!trimmed) return;
              if (
                section.propagateColumns.some(
                  (c) => c.name.toLowerCase() === trimmed.toLowerCase()
                )
              ) {
                return;
              }
              onChangeSection({
                propagateColumns: [
                  ...section.propagateColumns,
                  { name: trimmed },
                ],
              });
            }}
          />
        </div>
        {section.propagateColumns.length === 0 && (
          <p className="text-xs text-muted">
            No propagate columns yet. Add one — e.g.{" "}
            <code className="bg-background px-1 rounded">Phone</code> — and
            map it to each linked sheet&apos;s actual header below. Pro will
            fill blanks across sheets for these columns.
          </p>
        )}
      </div>
```

- [ ] **Step 2: Add the `AddPropagateColumn` component at the bottom of the file**

Append at the end of the file (after `LinkedSheetRow`):

```typescript
// ---------- AddPropagateColumn ----------

function AddPropagateColumn({ onAdd }: { onAdd: (name: string) => void }) {
  const [editing, setEditing] = useState(false);
  const [value, setValue] = useState("");
  if (!editing) {
    return (
      <button
        type="button"
        onClick={() => setEditing(true)}
        className="text-xs font-medium px-2 py-1 rounded-md border border-dashed border-border bg-background hover:bg-card transition-colors cursor-pointer"
      >
        + Add column
      </button>
    );
  }
  return (
    <span className="inline-flex items-center gap-1">
      <input
        autoFocus
        type="text"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            onAdd(value);
            setValue("");
            setEditing(false);
          } else if (e.key === "Escape") {
            setValue("");
            setEditing(false);
          }
        }}
        onBlur={() => {
          if (value.trim()) onAdd(value);
          setValue("");
          setEditing(false);
        }}
        placeholder="Phone"
        className="border border-border rounded-md px-2 py-1 text-xs bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary w-24"
      />
    </span>
  );
}
```

- [ ] **Step 3: Render the per-sheet mapping inside `LinkedSheetRow`**

In the `LinkedSheetRow` function, remove the `void _propagateColumns; // mapping UI lands in Task 12` line. Then, at the end of the function's JSX (right before the closing `</div>` of the outermost div), insert this block:

```typescript
      {_propagateColumns.length > 0 && linked.ui.loadedTabs.length > 0 && (
        <div className="space-y-2 border-t border-border pt-3">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Column mapping
          </p>
          <p className="text-[11px] text-muted">
            Type the header text from <code>{linked.tabName || "this tab"}</code>{" "}
            that should feed each section column. Leave blank to skip this
            sheet for that column.
          </p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            {_propagateColumns.map((col) => (
              <label
                key={col.name}
                className="flex items-center gap-2 text-xs"
              >
                <span className="font-medium w-24 truncate">{col.name}</span>
                <input
                  type="text"
                  value={linked.columnMapping[col.name] ?? ""}
                  onChange={(e) => {
                    const v = e.target.value;
                    onChange({
                      columnMapping: {
                        ...linked.columnMapping,
                        [col.name]: v ? v : null,
                      },
                    });
                  }}
                  placeholder="(skip)"
                  className="flex-1 min-w-0 border border-border rounded-md px-2 py-1 text-xs bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
                />
              </label>
            ))}
          </div>
        </div>
      )}
```

Rename the `_propagateColumns` parameter to `propagateColumns` (and update the two references inside the new JSX) since it's now used:

Find:

```typescript
  propagateColumns: _propagateColumns,
```

Replace with:

```typescript
  propagateColumns,
```

And replace both `_propagateColumns.length` / `_propagateColumns.map` occurrences in the JSX you just inserted with `propagateColumns.length` / `propagateColumns.map`.

- [ ] **Step 4: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 5: Commit**

```bash
git add src/app/sync-pro/page.tsx
git commit -m "$(cat <<'EOF'
feat(sync-pro): propagate columns chip row + per-sheet mapping UI

Adds the section-level "Propagate columns" chip row (add / remove
logical columns; in-place edit via AddPropagateColumn component).
Removing a column also wipes that key from every linked sheet's
columnMapping so config stays clean.

Inside LinkedSheetRow, when both propagate columns exist on the
section AND the linked sheet has loaded tabs, a "Column mapping"
section appears with one free-text input per logical column. The
user types the actual header text from this tab to map; an empty
value persists as null (= skip this sheet for that column).

Free-text was chosen over a header dropdown to avoid adding a new
/api/sheets/headers route just for this UI. If the user later wants
dropdowns we can add the endpoint and swap the input for a select.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 13: Page UI — per-section Run + Run all + Result panel

Wires the run buttons to the real POST endpoint and renders an inline result panel (`InlineResult`) below each section card with stats, per-column stats, and the conflict log.

**Files:**

- Modify: `src/app/sync-pro/page.tsx`

- [ ] **Step 1: Add validation + serialize helpers near the top of `SyncProPage`**

In `src/app/sync-pro/page.tsx`, find the line:

```typescript
  // Stubs filled in by Task 12 / 13.
  function handleRunSection(_idx: number) {
    void _idx;
  }
  async function handleRunAll(e: React.FormEvent) {
    e.preventDefault();
  }
```

Replace those eight lines with:

```typescript
  function validateSection(s: SectionState): string | null {
    if (!s.name.trim()) return "Section needs a name.";
    if (!s.masterTabName.trim()) return "Master tab name is required.";
    if (!s.linkedSheets.length) return "Add at least one linked sheet.";
    for (let i = 0; i < s.linkedSheets.length; i++) {
      const l = s.linkedSheets[i];
      if (!l.url) return `Linked sheet #${i + 1}: URL is required.`;
      if (!l.nickname.trim())
        return `Linked sheet #${i + 1}: nickname is required.`;
      if (!l.tabName)
        return `Linked sheet #${i + 1}: pick a tab (use Fetch tabs first).`;
    }
    return null;
  }

  async function handleRunSection(idx: number) {
    const section = sections[idx];
    if (!section) return;
    const err = validateSection(section);
    if (err) {
      setSectionError(section.id, err);
      return;
    }
    setSectionError(section.id, null);
    setSectionRunning(section.id, true);
    try {
      const payload = { sections: [toConfigSection(section)] };
      const res = await fetch("/api/sync-pro", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        setSectionError(section.id, data.error || "Failed to run");
        return;
      }
      const sec = data.sections?.[0];
      if (sec) setSectionResult(section.id, sec as ProSectionResult);
    } catch {
      setSectionError(section.id, "Network error");
    } finally {
      setSectionRunning(section.id, false);
    }
  }

  async function handleRunAll(e: React.FormEvent) {
    e.preventDefault();
    let anyFailed = false;
    for (const s of sections) {
      const err = validateSection(s);
      if (err) {
        setSectionError(s.id, err);
        anyFailed = true;
      } else {
        setSectionError(s.id, null);
      }
    }
    if (anyFailed) return;
    setRunningSections(new Set(sections.map((s) => s.id)));
    try {
      const payload = { sections: sections.map(toConfigSection) };
      const res = await fetch("/api/sync-pro", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        const msg = data.error || "Failed to run";
        for (const s of sections) setSectionError(s.id, msg);
        return;
      }
      const newResults = new Map<string, ProSectionResult>();
      for (const r of (data.sections ?? []) as ProSectionResult[]) {
        newResults.set(r.sectionId, r);
      }
      setResults((prev) => {
        const merged = new Map(prev);
        for (const [k, v] of newResults) merged.set(k, v);
        return merged;
      });
    } catch {
      for (const s of sections) setSectionError(s.id, "Network error");
    } finally {
      setRunningSections(new Set());
    }
  }
```

- [ ] **Step 2: Replace the placeholder result block inside `SectionCard`**

Find this block inside `SectionCard`:

```typescript
      {/* Result panel — filled in Task 13. */}
      {result && (
        <div className="border-t border-border pt-4 text-xs text-muted">
          Last run produced a result (rendering UI lands in Task 13).
          Cells filled: {result.totalCellsFilled} · Conflicts:{" "}
          {result.totalConflicts}
        </div>
      )}
```

Replace it with:

```typescript
      {result && <InlineResult result={result} />}
```

- [ ] **Step 3: Append the `InlineResult` + `Stat` components at the bottom of the file**

Append at the very end of `src/app/sync-pro/page.tsx`:

```typescript
// ---------- InlineResult ----------

function InlineResult({ result }: { result: ProSectionResult }) {
  const failed = !!result.error;
  const visibleConflicts = result.conflicts.slice(0, 10);
  const hiddenConflicts = Math.max(
    0,
    result.conflicts.length - visibleConflicts.length
  );
  return (
    <div
      className={`border-t pt-4 ${
        failed ? "border-danger/30" : "border-border"
      }`}
    >
      <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2 mb-3">
        <h3 className="font-medium text-sm">Last run</h3>
        {result.masterSpreadsheetUrl && (
          <a
            href={result.masterSpreadsheetUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="text-sm text-primary hover:underline font-medium"
          >
            Open {result.masterTabName} →
          </a>
        )}
      </div>

      {failed && (
        <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2 mb-3">
          {result.error}
        </p>
      )}

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-sm">
        <Stat label="Unique emails" value={result.totalUniqueEmails} />
        <Stat
          label="Cells filled"
          value={result.totalCellsFilled}
          accent="success"
        />
        <Stat
          label="Conflicts"
          value={result.totalConflicts}
          accent={result.totalConflicts > 0 ? "danger" : "muted"}
        />
        <Stat
          label="Present In"
          value={result.presentInWritten ? 1 : 0}
          accent="muted"
          hint={result.presentInWritten ? "written" : "skipped"}
        />
      </div>

      {result.columnStats.length > 0 && (
        <div className="mt-3 space-y-1">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Per column
          </p>
          {result.columnStats.map((c) => (
            <div
              key={c.name}
              className="flex items-center justify-between text-xs py-1 border-b border-border last:border-0"
            >
              <code className="font-medium">{c.name}</code>
              <span className="text-muted">
                {c.cellsFilled} filled · {c.conflicts} conflict
                {c.conflicts === 1 ? "" : "s"} · {c.skippedSheets} skipped
                sheet{c.skippedSheets === 1 ? "" : "s"}
              </span>
            </div>
          ))}
        </div>
      )}

      {result.linkedSheets.length > 0 && (
        <div className="mt-3 space-y-1">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Per sheet
          </p>
          {result.linkedSheets.map((t, i) => (
            <div
              key={i}
              className="flex flex-col gap-0.5 text-xs py-1 border-b border-border last:border-0"
            >
              <div className="flex items-center justify-between gap-2">
                <span className="font-medium">{t.nickname}</span>
                {t.error ? (
                  <span className="text-danger">{t.error}</span>
                ) : (
                  <span className="text-muted">
                    {t.rowsRead} rows · {t.emailsFound} emails ·{" "}
                    {t.cellsFilled} cells filled
                  </span>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {visibleConflicts.length > 0 && (
        <div className="mt-3 space-y-1">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Conflicts (first {visibleConflicts.length})
          </p>
          <div className="space-y-1">
            {visibleConflicts.map((c, i) => (
              <div
                key={i}
                className="text-xs text-muted py-1 border-b border-border last:border-0"
              >
                <code className="font-medium">{c.email}</code> ·{" "}
                <code>{c.column}</code>:{" "}
                {c.values
                  .map((v) => `${v.nickname}="${v.value}"`)
                  .join(" vs ")}
              </div>
            ))}
          </div>
          {hiddenConflicts > 0 && (
            <p className="text-[11px] text-muted">
              … and {hiddenConflicts} more (open the master tab to see all).
            </p>
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Stat ----------

function Stat({
  label,
  value,
  accent,
  hint,
}: {
  label: string;
  value: number;
  accent?: "success" | "danger" | "muted";
  hint?: string;
}) {
  const valueClass =
    accent === "success"
      ? "text-success"
      : accent === "danger"
        ? "text-danger"
        : accent === "muted"
          ? "text-muted"
          : "text-foreground";
  return (
    <div className="bg-background border border-border rounded-md px-3 py-2">
      <p className="text-xs text-muted">{label}</p>
      <p className={`text-lg font-semibold ${valueClass}`}>{value}</p>
      {hint && <p className="text-[10px] text-muted mt-0.5">{hint}</p>}
    </div>
  );
}
```

- [ ] **Step 4: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors.

- [ ] **Step 5: Commit**

```bash
git add src/app/sync-pro/page.tsx
git commit -m "$(cat <<'EOF'
feat(sync-pro): per-section Run + Run all + InlineResult panel

handleRunSection / handleRunAll wired to POST /api/sync-pro with
validation that catches missing names / URLs / tabs before sending.
Per-section validation errors surface inline on that section's
card; Run all aborts the batch if any section fails validation
(other sections' previous results stay visible).

InlineResult renders a stats grid (unique emails, cells filled,
conflicts, Present In written), per-column stats (filled / conflicts
/ skipped sheets), per-sheet stats (rows / emails / cells filled +
any error), and a conflict log showing the first 10 with a "… and N
more" footer. Each section's "Open <masterTabName>" link deep-jumps
to its master tab.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 14: Update CLAUDE.md

Adds the Pro feature to the storage-model + feature-module section so future agents can find it.

**Files:**

- Modify: `CLAUDE.md`

- [ ] **Step 1: Add the new hidden tab to the storage-model section**

In `CLAUDE.md`, find the block:

```markdown
- `_config` — linked source sheets for the main sync (`src/lib/config-store.ts`). Columns: `url, nickname, tabName, emailColumn, lastSynced`. `ensureConfigTab` lazily creates this tab and migrates old 4-column schemas to the current 5-column layout on read.
- `_report_sync_config` — `src/lib/report-config-store.ts`
- `_biz_tutor_config` — `src/lib/biz-tutor-config-store.ts`
```

Replace it with:

```markdown
- `_config` — linked source sheets for the main sync (`src/lib/config-store.ts`). Columns: `url, nickname, tabName, emailColumn, lastSynced`. `ensureConfigTab` lazily creates this tab and migrates old 4-column schemas to the current 5-column layout on read.
- `_report_sync_config` — `src/lib/report-config-store.ts`
- `_biz_tutor_config` — `src/lib/biz-tutor-config-store.ts`
- `_consolidator_config` — `src/lib/consolidator-config-store.ts`
- `_sync_pro_config` — `src/lib/sync-pro-config-store.ts`
```

- [ ] **Step 2: Add the Pro feature note to the existing-features paragraph**

Find this block:

```markdown
Existing features: main sync (`sync-engine`), `email-finder`, `biz-tutor-sync`, `renewal-sync`, `report-sync`, `domain-analyzer`, `consolidator`, `duplicate-finder`. The main sync is the only one that uses `_config` — the others use their own `_<feature>_config` tab and don't share state.
```

Replace it with:

```markdown
Existing features: main sync (`sync-engine`), `email-finder`, `biz-tutor-sync`, `renewal-sync`, `report-sync`, `domain-analyzer`, `consolidator`, `duplicate-finder`, `sync-pro`. The main sync is the only one that uses `_config` — the others use their own `_<feature>_config` tab and don't share state.

`sync-pro` is the multi-section, column-propagation sibling of the basic main sync. Each section has its own linked-sheets list, propagate-columns list, per-sheet column mapping, and master-tab name. Within a section: read every linked sheet, build a cross-reference by lowercased email, then for each logical propagate column fill any blank cell using the non-blank value from another linked sheet (never overwrite an existing non-blank value). Multiple non-blank values that disagree are surfaced in `result.conflicts` rather than touched. After propagation, the existing `writePresentInColumn` helper adds a "Present In" deep-link column to each source sheet (toggleable per section), and the section writes a `[Name, Email, ✅/❌ per sheet]` Master tab into the user's master spreadsheet at `section.masterTabName` (defaults to `Pro: <section name>`). The prefix-folding name-merge logic is duplicated verbatim from `sync-engine.ts` so the basic sync stays 100% untouched. Sections are run independently via per-section `Run section` buttons (or "Run all N sections"); per-section failures don't abort the batch.
```

- [ ] **Step 3: Verify**

Run: `npx tsc --noEmit && npm run lint`
Expected: no errors (docs change only — should be a no-op).

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md
git commit -m "$(cat <<'EOF'
docs(claude.md): add sync-pro to storage model + feature index

Documents the new _sync_pro_config hidden tab and the sync-pro
feature module. Notes the duplicated prefix-folding name-merge
logic and the per-sheet column mapping pattern.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 15: End-to-end manual verification

There's no test runner. This task is the manual smoke test that proves the feature works end-to-end against real Google Sheets.

**Files:** none (verification only)

- [ ] **Step 1: Build verification**

Run:

```bash
npx tsc --noEmit && npm run lint && npm run build
```

Expected: clean typecheck, clean lint, successful Next.js build with no warnings about the new routes / page.

- [ ] **Step 2: Start dev server**

```bash
npm run dev
```

Expected: server listening on `http://localhost:3000`.

- [ ] **Step 3: Manual smoke walk-through**

In a browser:

1. Log in with Google. Set a master sheet on the **Sheets** page (skip if already set).
2. Open **Sync Pro** from the nav. Confirm the page renders with one empty section and no console errors.
3. Name the section, add 2 linked sheets pointing at two test Google Sheets you own (use throw-away copies). Fetch tabs on each, pick a tab, set email column to `auto`.
4. Add a propagate column called `Phone`. In sheet #1, leave `Phone` mapping blank for one row. In sheet #2, put a value for the same email's `Phone`. Make sure each linked sheet has an `Email` column.
5. In Sync Pro, map `Phone` → `Telefono` (or whatever your column is named) for each linked sheet. Leave one sheet's mapping blank to confirm "skip" works.
6. Click **Run section**. Verify:
   - The result panel shows `unique emails`, `cells filled ≥ 1`, `conflicts = 0` if no disagreements.
   - The source sheet that had a blank `Phone` for the shared email now has the value from the other sheet.
   - The other source sheet was NOT overwritten.
   - A new tab named `Pro: <section name>` exists in your master sheet with a `[Name, Email, <nick1>, <nick2>]` matrix.
   - Each source sheet got a `Present In` column with hyperlinks to the other sheet for shared emails.
7. Create a conflict: put a different `Phone` value for the same email in sheet #1's row. Rerun the section. Verify:
   - The result panel shows `conflicts ≥ 1`.
   - The conflict log lists `email · Phone: <nick1>="..." vs <nick2>="..."`.
   - Neither cell got overwritten.
8. Toggle `Also write Present In column into each source` OFF. Run again. Verify the new "Present In" column ISN'T re-written on this run (existing one stays as-is from the previous run).
9. Click **+ Add section**, add a second section pointing at a different sheet pair. Run only section 1 — confirm section 2's "Last run" panel stays empty. Then run section 2 independently.
10. Click **Run all N sections** with both sections valid. Verify both run and both result panels update.
11. Refresh the page. Verify config rehydrates from `_sync_pro_config` (URLs, tab names, mappings, propagate columns all come back).
12. Verify no PII / secrets leaked to the browser console.

- [ ] **Step 4: Commit nothing — record findings**

No code changes here. If any of the steps above failed, file the bug as a separate task with the failing step number + actual behavior, then fix it before shipping. If everything passed, push the branch:

```bash
git log --oneline -20
git push origin main
```

Expected: 14 commits in the feature pushed (Tasks 1–14 each produced one commit), `main` updated cleanly with no force-push.

---

## Spec coverage check

- ✅ Match key (email only) — Task 5 (`findEmailColumnIndex` + `normalizeEmail`).
- ✅ Column propagation between source sheets — Task 6.
- ✅ Master ✅/❌ presence matrix — Task 7.
- ✅ Fill-blanks-only conflict rule with conflict log — Task 6.
- ✅ Per-sheet column mapping — Tasks 5, 11, 12.
- ✅ Master tab in master sheet, section-named — Task 7.
- ✅ Present In write-back with per-section toggle — Task 7 (toggle in Task 11 UI).
- ✅ Multiple independent sections, per-section Run, Run all — Tasks 10, 13.
- ✅ Config persistence in `_sync_pro_config` — Tasks 2, 3, 10.
- ✅ Basic sync untouched — confirmed (no edits to `sync-engine.ts`).
- ✅ Name prefix-folding duplicated, not extracted — Task 4.
- ✅ File layout — created in Tasks 1, 2, 3, 4, 8, 10; nav edited in Task 9.
- ✅ Edge cases (missing email column, mapped header gone, zero propagate columns, zero linked sheets, duplicate linked entries) — handled across Tasks 5–7 (return per-sheet errors, section error for zero sheets, blank-mapping = skip, zero propagate = no-op).
- ✅ End-to-end manual verification — Task 15.

## Placeholder scan

- No "TBD" / "TODO" / "fill in details" / "similar to Task N" placeholders.
- Every code step includes the actual code; every command step includes the exact command + expected output where applicable.
- The `void _unused` patterns in Task 4 are intentional skeleton-compiles-clean tricks; they get replaced with real usage in Tasks 5–7 (verified by the spec-coverage list above).

## Type consistency check

- `ProSection` / `ProLinkedSheet` / `ProPropagateColumn` defined once in Task 1; used unchanged in every later task.
- `runSyncProSection(refreshToken, masterSheetId, section)` signature defined in Task 4, used in Task 8, used in Tasks 5–7 (implementation grows in place, signature stable).
- `writePresentInColumn` parameter order matches the existing helper in `src/lib/present-in-writer.ts` (Task 7 mirrors basic sync's call site).
- `valueInputOption` strings are spec-correct (`"USER_ENTERED"` for propagation writes in Task 6; `"RAW"` for the Master tab in Task 7).
