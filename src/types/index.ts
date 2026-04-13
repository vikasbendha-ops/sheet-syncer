export interface LinkedSheet {
  url: string;
  nickname: string;
  tabName: string; // the specific tab/sheet within the spreadsheet
  emailColumn: string; // "auto" or a column letter like "A", "B"
  lastSynced: string; // ISO timestamp or empty
}

export interface EmailReadResult {
  emails: Set<string>;
  columnLetter: string;
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
