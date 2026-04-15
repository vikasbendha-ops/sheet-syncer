export interface LinkedSheet {
  url: string;
  nickname: string;
  tabName: string; // the specific tab/sheet within the spreadsheet
  emailColumn: string; // "auto" or a column letter like "A", "B"
  lastSynced: string; // ISO timestamp or empty
}

export interface EmailReadResult {
  emails: Map<string, number[]>; // email → 1-based row numbers
  names: Map<string, string>; // email → name (first non-empty occurrence)
  columnLetter: string;
  presentInColumnIndex: number | null; // 0-based index of existing "Present In" column, or null
  lastColumnIndex: number; // 0-based index of the last filled header column
  sheetId: number; // numeric sheetId of the resolved tab
  resolvedTabName: string;
}

export interface SyncResult {
  success: boolean;
  sheetsProcessed: number;
  totalEmails: number;
  errors: string[];
  timestamp: string;
}

export interface MasterSheetData {
  headers: string[]; // ["Email", "Sheet A", "Sheet B", ...]
  rows: string[][]; // [["email@example.com", "YES", "NO", ...], ...]
}
