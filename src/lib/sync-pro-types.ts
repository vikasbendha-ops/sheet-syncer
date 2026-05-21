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
