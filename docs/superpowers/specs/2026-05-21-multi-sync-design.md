# Multi Sync — Design Spec

**Status:** Approved (brainstorm round, 2026-05-21)

## Goal

Add a **Multi Sync** feature: an exact behavioral duplicate of the existing basic Present In main sync (`src/lib/sync-engine.ts`), but wrapped in the multi-section pattern. Each section is an independent basic-sync config (own linked sheets, own Master tab, own Present In column).

No column propagation (that's Sync Pro's job). No per-sheet column mapping. Just: basic Present In behavior, multiplied by N independent sections.

## Locked decisions

| Decision | Choice |
|---|---|
| Menu name | "Multi Sync" |
| Per-section Present In column | Unique per section. Default `"Present In - <slot>"` where slot is a monotonic counter. User can rename. Uniqueness validated. |
| Slot counter | Stored in `_multi_sync_meta` hidden tab (single row, single cell). Increments on every section create. Never decrements, even when sections are deleted. |
| Per-section / per-linked-sheet lastSynced | Yes — each linked sheet inside each section tracks its own timestamp. |
| Master tab destination | In the user's master spreadsheet. Default `"Multi Master - <slot>"`, user-editable. |
| Basic sync (`/sheets`, `/api/sync`) | 100% untouched. |
| Sync Pro | 100% untouched. |
| Name-merge | Prefix-folding logic duplicated again into the Multi Sync engine. |

## Data model

```ts
interface MultiLinkedSheet {
  url: string;
  nickname: string;
  tabName: string;
  emailColumn: string;        // "auto" or column letter
  lastSynced: string;         // ISO timestamp, "" if never synced
}

interface MultiSyncSection {
  id: string;                  // sec_<ts>_<rand>
  name: string;                // user-typed
  /**
   * Permanent monotonic slot id assigned at section creation. Drives the
   * default column / master tab names. Never reused after section delete.
   */
  slot: number;
  masterTabName: string;        // default `Multi Master - ${slot}`
  presentInColumnName: string;  // default `Present In - ${slot}`
  linkedSheets: MultiLinkedSheet[];
}
```

## Engine flow (per section)

`runMultiSyncSection(refreshToken, masterSheetId, section): Promise<MultiSyncSectionResult>`

1. **Read every linked sheet** in parallel under `MAX_CONCURRENCY=3`. Per sheet:
   - One `spreadsheets.get` per unique spreadsheetId (cached).
   - Resolve email column (auto-detect or letter).
   - `readEmailsFromSheet`-equivalent batchGet (email + name columns).
   - Build `Map<normalizedEmail, RowRecord>`.
2. **Build cross-reference**: `Map<email, Array<{linkedSheetIdx, sheetId, rowNumber}>>`.
3. **Write Present In column** into each source sheet using `section.presentInColumnName` as the header text.
   - Find existing column whose header **exactly equals** the section's `presentInColumnName`. Reuse if found, otherwise append at `lastFilledColumnIndex + 1`.
   - Never touches the unsuffixed `"Present In"` (basic sync's) or any other section's column.
   - Cell content: hyperlinks to the OTHER sheets in this section where that email is found.
4. **Write Master tab** at `section.masterTabName` in the user's master spreadsheet.
   - Clear + rewrite full.
   - Header: `[Name, Email, <nickname1>, <nickname2>, ...]`.
   - Rows: alphabetically sorted by email, ✅/❌ per linked sheet. Name from prefix-folding merge.
5. **Update lastSynced**: batched `values.batchUpdate` writes the current ISO timestamp back into each linked sheet's row in `_multi_sync_config`.

Every Google API call wrapped in `withRetry`. Per-sheet failures don't abort the section.

## Result shape

```ts
interface MultiTabResult {
  nickname: string;
  url: string;
  tabName: string;
  rowsRead: number;
  emailsFound: number;
  error?: string;
}

interface MultiSyncSectionResult {
  sectionId: string;
  sectionName: string;
  slot: number;
  masterSpreadsheetUrl: string;
  masterTabName: string;
  presentInColumnName: string;
  totalUniqueEmails: number;
  linkedSheets: MultiTabResult[];
  presentInWritten: boolean;
  timestamp: string;     // ISO of this run
  error?: string;
}

interface MultiSyncBatchResult {
  sections: MultiSyncSectionResult[];
}
```

## Persistence

Two hidden tabs in the master sheet:

**`_multi_sync_meta`** — single-cell counter
- Header: `[nextSlot]`. Row 2 cell A2 holds the integer.
- Read on every section create. Increment, write back, return the claimed slot.

**`_multi_sync_config`** — one row per section
- Headers: `[sectionId, sectionName, slot, masterTabName, presentInColumnName, linkedSheets]`
- `linkedSheets` is a JSON array of `MultiLinkedSheet` objects.

## File layout

```
src/lib/
  multi-sync-types.ts          (NEW — shared types)
  multi-sync-config-store.ts   (NEW — ensureTab + get/save/clear + claimSlot)
  multi-sync-engine.ts         (NEW — runMultiSyncSection + runMultiSyncBatch)
src/app/api/multi-sync/
  route.ts                     (NEW — POST runs sections)
  config/route.ts              (NEW — GET / PUT / DELETE / claimSlot)
src/app/multi-sync/
  page.tsx                     (NEW — multi-section UI)
src/components/nav.tsx         (EDIT — add "Multi Sync" link)
src/lib/present-in-writer.ts   (EDIT — add optional `headerText` parameter, default "Present In")
CLAUDE.md                      (EDIT — add feature note + new hidden tabs)
```

## Helper modification

`writePresentInColumn` currently hardcodes the header text as `"Present In"`. Add a new optional parameter `headerText?: string` with default `"Present In"`. All existing callers (basic sync, Sync Pro) keep their behavior unchanged. Multi Sync passes `section.presentInColumnName`.

## UI

New nav item: **Multi Sync**. Page mirrors Sync Pro's layout (per-section card + Run section + Run all + auto-saved config).

Per-section card:
- Header: `Section N — [name input]    [Remove section]`
- `Slot: N` (read-only display)
- `Master tab name: [input]` (default `Multi Master - <slot>`)
- `Present In column: [input]` (default `Present In - <slot>`)
- Linked sheets list (each: nickname + URL + Fetch tabs + Tab dropdown + Email col input + per-sheet `lastSynced` display + Remove sheet)
- `+ Add another linked sheet`
- Per-section `[Run section]` button + inline error
- Inline result panel after a run (stats + per-sheet breakdown + open-master link)

`Run all N sections` button at the bottom (only when 2+ sections).

## Validation rules

- Section needs at least one linked sheet.
- Each linked sheet needs URL + nickname + tabName.
- `masterTabName` non-empty.
- `presentInColumnName` non-empty, must be unique across all current sections (case-insensitive trim).
- `slot` is assigned by the server's `claimSlot` endpoint when a section is created; the client never picks the slot.

## Non-goals (YAGNI cuts)

- No column propagation (Sync Pro covers this).
- No per-sheet column mapping (basic sync style — auto-detect everything).
- No cross-section reconciliation.
- No "free / reclaim slot" feature — deleted slots stay burned forever.
- No undo / history.
- No conditional formatting / chip dropdowns on Master.

## Edge cases

- **Deleting a section**: the section row vanishes from `_multi_sync_config`. The counter in `_multi_sync_meta` is NOT decremented. The Present In columns this section wrote into source sheets are NOT removed (those stay until user manually deletes).
- **Renaming `presentInColumnName` after a section has already run once**: section's next run will write to the NEW column name and the OLD column will be left untouched (stale). Documented behavior — user can manually delete the stale column.
- **Two sections accidentally configured with the same `presentInColumnName`**: caught by uniqueness validation in the API route (and on save in the config store). Save returns 400; UI surfaces it.
- **The slot counter is corrupted / missing**: `claimSlot` initializes it to 1 if `_multi_sync_meta` is empty.
- **A linked sheet's URL changes**: the next read fails for that sheet (no migration logic). `lastSynced` keeps its old value, error captured in the result panel.
