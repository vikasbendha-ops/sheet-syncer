// src/lib/multi-sync-types.ts

/**
 * Multi Sync — exact behavioral duplicate of the basic Present In main sync,
 * wrapped in the multi-section pattern. Each section is independent and
 * writes its own uniquely-named Present In column into source sheets.
 *
 * Slot is a permanent monotonic counter stored in `_multi_sync_meta`. The
 * default Present In column name and Master tab name are derived from the
 * slot so values are unique across all current and previously-deleted
 * sections.
 */

export interface MultiLinkedSheet {
  url: string;
  nickname: string;
  tabName: string;
  /** "auto" or column letter. */
  emailColumn: string;
  /** ISO timestamp of the last successful read of this sheet (or ""). */
  lastSynced: string;
}

export interface MultiSyncSection {
  /** Stable id (sec_<ts>_<rand>). */
  id: string;
  /** User-typed display name. */
  name: string;
  /**
   * Permanent monotonic slot id assigned at section creation. Drives the
   * default Master tab name and Present In column name. Never reused after
   * section delete.
   */
  slot: number;
  /** Default `Multi Master - <slot>`. */
  masterTabName: string;
  /**
   * Default `Present In - <slot>`. Must be unique across all current
   * sections (case-insensitive trim). The section only ever touches the
   * column whose header exactly matches this value.
   */
  presentInColumnName: string;
  linkedSheets: MultiLinkedSheet[];
}

export interface MultiTabResult {
  nickname: string;
  url: string;
  tabName: string;
  rowsRead: number;
  emailsFound: number;
  error?: string;
}

export interface MultiSyncSectionResult {
  sectionId: string;
  sectionName: string;
  slot: number;
  /** Deep link to this section's Master tab. */
  masterSpreadsheetUrl: string;
  masterTabName: string;
  presentInColumnName: string;
  totalUniqueEmails: number;
  linkedSheets: MultiTabResult[];
  presentInWritten: boolean;
  /** ISO of this run. */
  timestamp: string;
  /** Fatal: whole section bailed before producing useful output. */
  error?: string;
}

export interface MultiSyncBatchResult {
  sections: MultiSyncSectionResult[];
}
