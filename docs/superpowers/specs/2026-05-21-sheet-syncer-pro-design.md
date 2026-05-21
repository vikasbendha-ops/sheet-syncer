# Sheet Syncer Pro — Design Spec

**Status:** Approved (brainstorm round, 2026-05-21)
**Author:** brainstorm pair (vikas + claude)
**Implementation owner:** TBD

## Goal

Add a new feature module — **Sync Pro** — that is a more flexible sibling of the existing basic "Present In" main sync. The basic sync stays exactly as it is. Pro adds two capabilities the basic version lacks:

1. **Column propagation between linked source sheets.** For user-selected columns, if one linked sheet has a value for a given email but another doesn't, Pro fills in the blank. Never overwrites a non-blank cell. Surfaces value conflicts in the result.
2. **Multiple independent sections in one page.** Each section is its own self-contained Pro sync config (own linked sheets, own propagate columns, own Master tab name) — no need to wipe and reconfigure between use cases.

## Decisions locked during brainstorm

| Decision | Choice |
|---|---|
| Match key | Lowercased email, same as basic |
| What "pick columns" means | (a) which columns get propagated BETWEEN source sheets + (b) keep basic ✅/❌ presence matrix in Master |
| Conflict rule | Fill blanks only — never overwrite a non-blank cell. Surface conflicts in result. |
| Column matching across sheets | Per-sheet explicit mapping (user picks each linked sheet's actual column for each section-defined logical column) |
| Master tab destination | User's master sheet, one tab per section (`Pro: <section name>` default, user-editable) |
| Present In write-back | Yes (same as basic), with per-section toggle |
| Engine architecture | Separate feature module mirroring `consolidator` layout. Basic `sync-engine.ts` stays untouched. |
| Shared name-merge logic | Duplicate the prefix-folding code into the Pro engine — do not refactor basic sync. |

## Data model

```ts
interface ProSection {
  id: string;              // stable id (sec_<ts>_<rand>)
  name: string;            // user-typed (e.g. "Italian newsletter ↔ CRM")
  masterTabName: string;   // tab written into master sheet. default "Pro: <name>"
  linkedSheets: ProLinkedSheet[];
  propagateColumns: ProPropagateColumn[];
  writePresentIn: boolean; // default true
}

interface ProLinkedSheet {
  url: string;
  nickname: string;
  tabName: string;         // single tab per linked sheet
  emailColumn: string;     // "auto" or column letter
  columnMapping: Record<string, string | null>;
    // logical column name (from section.propagateColumns) → actual header text in this sheet
    // null = skip this sheet for this column
}

interface ProPropagateColumn {
  name: string;            // user-defined logical name (e.g. "Phone", "Course")
}
```

A section like:

- Sheets: `[Newsletter (Subscribers tab), CRM (Contacts tab), Old Imports (Archive tab)]`
- Propagate columns: `[Phone, Course]`
- Mappings:
  - Newsletter: `Phone → Telefono`, `Course → Corso`
  - CRM: `Phone → Phone`, `Course → Course Name`
  - Old Imports: `Phone → skip`, `Course → Subscription`

## Engine flow (per section)

`runSyncProSection(refreshToken, masterSheetId, section): Promise<ProSectionResult>`

1. **Read every linked sheet** in parallel with `MAX_CONCURRENCY = 3`. One `spreadsheets.get` cached per unique spreadsheetId. Per linked sheet:
   - Header row resolves email column (auto-detect or letter) + each mapped logical column → column index.
   - Bulk `batchGet` for the email column, the name column(s) (for the Master tab's Name field), and every mapped propagate column.
   - Build `Map<normalizedEmail, RowRecord>` where `RowRecord = { rowNumber, sheetId, perColumn: Map<logicalCol, rawValue> }`.

2. **Cross-reference**: `Map<email, Array<{ linkedSheetIdx, sheetId, rowNumber }>>`. Same shape as basic — feeds the Present In writer.

3. **Propagation pass** (the core Pro behavior). For each logical propagate column:
   - For each email present in 2+ linked sheets:
     - Collect `{ linkedSheetIdx → value }` (only for sheets that mapped this column to a real header).
     - Bucket into `blanks` (cells where the value is empty/whitespace) and `nonBlanks`.
     - Decide:
       - 0 non-blank → no-op.
       - 1 non-blank → propagate that value to every blank cell. Record N writes.
       - 2+ non-blank, all equal after trim + case-insensitive compare → propagate to blank cells. Record N writes.
       - 2+ non-blank, distinct values → **conflict**. No writes. Record `{ email, column, values: Array<{ linkedSheetIdx, value }> }` in `result.conflicts`.

4. **Apply propagation writes** back into source sheets. One `values.batchUpdate` per source sheet collecting all `{ row, col, value }` cells we filled. Within the section, parallelize across sheets at `MAX_CONCURRENCY = 3`.

5. **Write Present In column** into each source sheet via the existing `present-in-writer.ts` (skip if `section.writePresentIn === false`). Same deep-link behavior as basic. The Present In cell for sheet A at row R lists hyperlinks to every other linked sheet where the row's email is also found.

6. **Write Master tab** to the user's master sheet at `section.masterTabName`:
   - Clear + rewrite.
   - Header: `[Name, Email, <nickname1>, <nickname2>, ...]`.
   - Rows: one per unique email, sorted alphabetically, with ✅/❌ per linked sheet.
   - Name comes from the prefix-folding merge — the same algorithm basic sync uses (logic duplicated into Pro engine; basic sync untouched).

7. **Return result** with per-sheet read counts, propagation stats per column (cells filled, conflicts), writebacks per sheet, Master row count, success flag.

Every Google API call wrapped in `withRetry`. Per-sheet failures don't abort the section — captured in `result.linkedSheets[i].error`. Per-section failures are caught by `runSyncProBatch` and surfaced as `result.sections[i].error`.

## Result shape

```ts
interface ProTabResult {
  nickname: string;
  url: string;
  tabName: string;
  rowsRead: number;
  emailsFound: number;
  cellsFilled: number;   // propagation writes to this sheet
  error?: string;
}

interface ProColumnStats {
  name: string;          // logical column name
  cellsFilled: number;
  conflicts: number;
  skippedSheets: number; // sheets that mapped this column to "skip"
}

interface ProConflict {
  email: string;
  column: string;
  values: Array<{ nickname: string; value: string }>;
}

interface ProSectionResult {
  sectionId: string;
  sectionName: string;
  masterSpreadsheetUrl: string;   // deep link to the section's master tab
  masterTabName: string;
  totalUniqueEmails: number;
  totalCellsFilled: number;
  totalConflicts: number;
  linkedSheets: ProTabResult[];
  columnStats: ProColumnStats[];
  conflicts: ProConflict[];        // detailed; UI shows first ~10 + count of rest
  presentInWritten: boolean;
  error?: string;
}

interface ProBatchResult {
  sections: ProSectionResult[];
}
```

## UI

New nav item: **Sync Pro**. Page mirrors Consolidator's per-section pattern (per-section card, per-section Run button, "Run all" at bottom, auto-saved config).

Per-section card contents:

- Header: `Section N — [section name input]    [Remove section]`
- Master tab name input (default `Pro: <name>`, regenerated when name changes if user hasn't manually overridden)
- `☑ Also write Present In column into each source` (default checked)
- **Propagate columns** row: chip-style list of logical column names with `+ Add column` and `×` to remove.
- **Linked sheets** list: one card per linked sheet.
  - Inside each linked sheet card:
    - Nickname input
    - URL input + `[Fetch tabs]` button
    - Tab selector (dropdown — single select)
    - Email column override (`auto` or letter A–Z)
    - **Column mapping**: for each propagate column defined on the section, render `LogicalName → [dropdown of this tab's headers | skip]`. Only renders once `Fetch tabs` has succeeded and the tab is selected.
    - `[Remove sheet]`
  - `+ Add another linked sheet`
- Per-section `[Run section]` button.
- Inline **Last run** panel after a run:
  - Top-line stats: `N sheets · M unique emails · K cells filled · C conflicts`
  - Per-column stats line
  - Conflict log: first ~10 conflicts, each formatted `email@x.it · Phone: Newsletter="333…" vs CRM="334…"`. Show `… and N more` if longer. Each conflict row has a small `Open in master` deep link.
  - `Open <masterTabName> →` link to the master tab.

`Run all N sections` button at the bottom (visible only when 2+ sections).

## Config persistence

Hidden tab `_sync_pro_config` in the master sheet.

Headers: `[sectionId, sectionName, masterTabName, writePresentIn, propagateColumns, linkedSheets]`

One row per section. The two JSON columns:

- `propagateColumns` → `[{"name": "Phone"}, ...]`
- `linkedSheets` → `[{"url", "nickname", "tabName", "emailColumn", "columnMapping": {"Phone": "Telefono", ...}}, ...]`

Save on debounced change (800ms). Auto-load on page mount. Clear button wipes the tab. Same pattern as `_consolidator_config`.

No legacy schema migration — Sync Pro is a brand-new feature, no prior schema exists.

## File layout

```
src/lib/
  sync-pro-engine.ts             # runSyncProSection, runSyncProBatch, name-merge (duplicated)
  sync-pro-config-store.ts       # ensureTab + get/save/clear

src/app/api/sync-pro/
  route.ts                       # POST runs one or many sections
  config/route.ts                # GET / PUT / DELETE config

src/app/sync-pro/
  page.tsx                       # multi-section UI

src/components/nav.tsx           # add { href: "/sync-pro", label: "Sync Pro" }
CLAUDE.md                        # add feature note + storage tab entry
```

**Shared modules reused as-is (no edits to basic sync):**

- `getSheetsClient`, `withRetry`, `extractSpreadsheetId`, `normalizeEmail`, `findEmailColumnIndex`, `findNameColumns`, `columnIndexToLetter` (`email-utils.ts`)
- `writePresentInColumn` from `present-in-writer.ts` (works as-is)
- `getTabNames` from `sheets-reader.ts`

**Duplicated into Pro engine (per user direction):**

- The prefix-folding name-merge block from `sync-engine.ts` lines ~230-284. Same algorithm, copy-pasted. Basic sync stays 100% untouched.

## Edge cases

- **Email column missing in a linked sheet.** Per-sheet error in `result.linkedSheets[i].error`, sheet skipped, other sheets continue. Same as basic.
- **Propagate column mapped to a header that no longer exists in the sheet** (user renamed it after saving config). Per-sheet warning. That column treated as `skip` for this sheet. Section run still succeeds.
- **All values for a column are blank.** No-op, recorded as 0 cellsFilled, 0 conflicts.
- **Section has zero propagate columns.** Pro becomes a multi-section variant of basic Present In — just writes Master ✅/❌ + Present In columns. Allowed.
- **Section has zero linked sheets.** Per-section error: "Section needs at least one linked sheet."
- **Two linked sheets are the same spreadsheet/tab.** Allowed; rows in that tab can match themselves trivially. Probably user error — emit a soft warning in the result if any two linked entries resolve to the same `(spreadsheetId, tabName)`.
- **User selects the master sheet itself as a linked sheet.** Allowed but warned. The Master tab is rewritten in step 6 so any cross-references from the basic sync's Master could get clobbered. Mitigation: explicit warning in the result.
- **`Run all` failures.** Each section runs independently. One section's failure doesn't abort the rest. UI shows the error on that section's card.

## Non-goals (explicit YAGNI cuts)

- Two-way merge with conflict resolution beyond fill-blanks. Conflicts are reported, not auto-resolved.
- Cross-section dedupe or shared registry. Sections are 100% independent.
- Writing more than the existing `Present In` column shape into source sheets (e.g. richer back-link cells).
- Schema validation for propagate columns (no enforcement of type/format — values copied verbatim).
- Undo / history. Re-run is full overwrite of Master tab and full additive blank-fill on sources.
- Conditional formatting / chip dropdowns on Master. Master stays a simple ✅/❌ matrix.

## Open items for the implementation plan

- Final wording / placement of the `Pro: ` prefix toggle (always-on default? user-removable?).
- Whether to render the column-mapping dropdowns inline under each linked sheet or as a separate matrix (linked-sheets × propagate-columns).
- Visual treatment of the conflict log — collapsible vs. always-shown.

These are not blockers — the implementation plan can pick reasonable defaults and the user can adjust during build.
