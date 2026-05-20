# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev      # next dev (http://localhost:3000)
npm run build    # next build
npm run start    # serve production build
npm run lint     # eslint (flat config: next/core-web-vitals + next/typescript)
```

No test runner is configured. The top-level `test` file is scratch, not a script.

## Required env vars

Auth and crypto fail closed if any of these are missing:

- `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `GOOGLE_REDIRECT_URI` — read lazily by `src/lib/google-auth.ts` on first OAuth use. Scopes requested: `spreadsheets`, `userinfo.email`. `access_type=offline` + `prompt=consent` means we always get a refresh token.
- `SESSION_SECRET` — ≥ 32 chars. Used by `src/lib/session.ts` to AES-256-GCM encrypt the session cookie (which holds the refresh token).
- `APP_URL` — optional. Set explicitly when running behind a reverse proxy (Hostinger, Nginx, etc.) where `request.url` would be the internal address. Otherwise `getAppUrl()` falls back to `GOOGLE_REDIRECT_URI`'s origin, then the request origin.

## Architecture

Next.js 16 App Router + React 19 + TypeScript strict. Tailwind v4 via `@tailwindcss/postcss`. Path alias `@/*` → `./src/*`. No database — see "Storage model" below.

### Storage model (important)

**The user's own master Google Sheet IS the database.** There is no Postgres, KV, or filesystem state. All per-user config is persisted in hidden tabs inside the master spreadsheet they pick during onboarding:

- `_config` — linked source sheets for the main sync (`src/lib/config-store.ts`). Columns: `url, nickname, tabName, emailColumn, lastSynced`. `ensureConfigTab` lazily creates this tab and migrates old 4-column schemas to the current 5-column layout on read.
- `_report_sync_config` — `src/lib/report-config-store.ts`
- `_biz_tutor_config` — `src/lib/biz-tutor-config-store.ts`
- `Master` — output of the main sync (`src/lib/sheets-writer.ts`). Cleared and rewritten in full on each sync.
- `Domains - <tab>` — output of the domain analyzer.

The session cookie only holds: `refreshToken, accessToken, expiresAt, email, masterSheetId`. Losing the cookie = re-auth, but never data loss.

### Request flow

1. User hits `/api/auth/login` → redirect to Google → `/api/auth/callback` exchanges code, calls `oauth2.userinfo.get()` for the email, encrypts everything into the session cookie.
2. Every API route does `getSession()` and rejects with 401 if `refreshToken` is missing, 400 if `masterSheetId` is missing. See `src/app/api/sheets/route.ts` `getAuthContext()` for the canonical pattern.
3. `getSheetsClient(refreshToken)` returns a `sheets_v4.Sheets` bound to a singleton `OAuth2Client` with `setCredentials({ refresh_token })`. The library auto-refreshes access tokens — we never store them long-term.

### Google API discipline

Every Google API call goes through `withRetry()` (`src/lib/retry.ts`): exponential backoff on 429 / 5xx / `rateLimitExceeded`. Don't bypass this — Sheets has a 60 reads/min/user quota that's easy to trip.

`runFullSync` in `src/lib/sync-engine.ts` is the perf-critical path. Two things matter:

- **`MAX_CONCURRENCY = 3`** for both read and write fan-out (was 5, lowered to stay under the per-user read quota). Use `withConcurrencyLimit()` from this file for any new fan-out.
- **Tab metadata is pre-fetched once per unique spreadsheet** into `tabMetadataCache` before reading begins. This avoids N extra `spreadsheets.get` calls and matters for Vercel's serverless timeout. `readEmailsFromSheet` accepts cached tabs as its 5th arg — pass them when looping.

### Sync engine semantics

`src/lib/sync-engine.ts` does more than aggregate:

1. Reads each linked tab (`sheets-reader.ts`) — auto-detects email column from headers unless overridden by `emailColumn: "A"|"B"|...`. Also detects name columns (full / first+last, English + Italian aliases — see `email-utils.ts findNameColumns`).
2. Builds a cross-reference map `email → [{nickname, spreadsheetId, sheetId, rowNumber}, ...]`.
3. Writes a **"Present In" column** back into each source tab (`present-in-writer.ts`) as deep-link hyperlinks (`#gid=<sheetId>&range=A<row>`). The column is reused if a `Present In` header already exists, otherwise appended at `lastColumnIndex + 1`.
4. Merges names across sheets with **prefix folding**: if "Laura" appears in one sheet and "Laura Pegoraro" in another, "Laura"'s count is folded into the longer name so the more complete form wins. Winner is picked by `mergedCount` → word count → first-seen order.
5. Wipes and rewrites the `Master` tab in full. Header is `[Name, Email, <nickname1>, <nickname2>, ...]`, presence marked with `✅`/`❌`.
6. Batches all `lastSynced` updates into a single `values.batchUpdate` call.

Errors per-sheet are collected into `result.errors` rather than aborting the whole sync. `success` is `errors.length === 0`.

### Feature module pattern

Each feature follows the same shape — when adding one, mirror it:

```
src/lib/<feature>.ts                        # engine: reads + writes via google-auth + retry
src/lib/<feature>-config-store.ts           # ensureTab() + load/save in a `_<feature>_config` tab
src/app/api/<feature>/route.ts              # POST runs the engine
src/app/api/<feature>/config/route.ts       # GET/PUT config
src/app/<feature>/page.tsx                  # client UI
```

Existing features: main sync (`sync-engine`), `email-finder`, `biz-tutor-sync`, `renewal-sync`, `report-sync`, `domain-analyzer`, `consolidator`, `duplicate-finder`. The main sync is the only one that uses `_config` — the others use their own `_<feature>_config` tab and don't share state.

`duplicate-finder` is read+format-only (no row writes). Treats all picked tabs of one source spreadsheet as a single dataset, finds duplicates in the `Email` and `Telefono Cellulare` columns, then paints cell backgrounds: first occurrence light green, subsequent occurrences light red. Singletons untouched. Phone matching strips non-digits before comparing (raw value preserved). Re-running clears formatting in just the Email + Phone columns of the picked tabs before repainting. No `_config` tab — scans are one-off, not persisted.

`consolidator` is multi-section + multi-source. Each **section** has one or more source spreadsheets (each with its own picked tabs), a user-chosen **output spreadsheet URL**, and a user-typed **output tab name**. Engine dedupes rows by lowercased Email across all sources within a section (phone-wins merge), then writes header `[Name, Surname, Email, Phone]` to the section's output tab (clear + rewrite each run). Multiple sections are configured in one form and run sequentially by `runConsolidatorBatch` — per-section failures are returned in `result.sections[i].error` rather than aborting the batch. Config schema in `_consolidator_config` is `[sectionId, name, outputUrl, outputTabName, sources]` with `sources` as a JSON array; the original single-section legacy schema (`[sourceUrl, sourceTabs]`) is detected by header sniff and migrated to one section on read.

### Notes when editing

- `extractSpreadsheetId` in `url-parser.ts` is the only sanctioned way to parse a Sheets URL. Don't regex it inline.
- Tab names containing apostrophes are escaped with `'` → `''` and wrapped in single quotes when building A1 ranges (see `sheets-reader.ts`). Use the same convention.
- The `_config` schema migration in `ensureConfigTab` runs on every load — keep it idempotent if you touch it.
- `cookies()` is `await`ed (Next 16). All session helpers are async.
