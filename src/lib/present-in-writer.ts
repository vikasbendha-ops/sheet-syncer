import { getSheetsClient } from "./google-auth";
import { sheets_v4 } from "googleapis";

const PRESENT_IN_HEADER = "Present In";
const SEPARATOR = "\n";

export interface PresentInCell {
  rowIndex: number; // 0-based row index (header is row 0)
  links: Array<{ text: string; url: string }>;
}

/**
 * Writes the "Present In" column to a specific tab in a spreadsheet.
 * Each cell contains comma-separated nicknames; each nickname is a hyperlink.
 */
export async function writePresentInColumn(
  refreshToken: string,
  spreadsheetId: string,
  sheetId: number,
  columnIndex: number, // 0-based
  headerNeeded: boolean,
  cellData: PresentInCell[]
): Promise<void> {
  const sheets = getSheetsClient(refreshToken);

  const requests: sheets_v4.Schema$Request[] = [];

  if (headerNeeded) {
    requests.push({
      updateCells: {
        rows: [
          {
            values: [
              {
                userEnteredValue: { stringValue: PRESENT_IN_HEADER },
                userEnteredFormat: { textFormat: { bold: true } },
              },
            ],
          },
        ],
        fields: "userEnteredValue,userEnteredFormat.textFormat.bold",
        start: {
          sheetId,
          rowIndex: 0,
          columnIndex,
        },
      },
    });
  }

  // Emit one updateCells per cell. Each cell uses WRAP so newlines render as multi-line.
  for (const cell of cellData) {
    const { text, runs } = buildCellContent(cell.links);
    requests.push({
      updateCells: {
        rows: [
          {
            values: [
              {
                userEnteredValue: { stringValue: text },
                textFormatRuns: runs,
                userEnteredFormat: {
                  wrapStrategy: "WRAP",
                  verticalAlignment: "TOP",
                },
              },
            ],
          },
        ],
        fields:
          "userEnteredValue,textFormatRuns,userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment",
        start: {
          sheetId,
          rowIndex: cell.rowIndex,
          columnIndex,
        },
      },
    });
  }

  if (requests.length === 0) return;

  // Split into chunks to avoid hitting request size limits (~100k per batch)
  const CHUNK_SIZE = 500;
  for (let i = 0; i < requests.length; i += CHUNK_SIZE) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: requests.slice(i, i + CHUNK_SIZE) },
    });
  }
}

function buildCellContent(
  links: Array<{ text: string; url: string }>
): { text: string; runs: sheets_v4.Schema$TextFormatRun[] } {
  const runs: sheets_v4.Schema$TextFormatRun[] = [];
  let text = "";

  links.forEach((link, i) => {
    const startIndex = text.length;
    runs.push({
      startIndex,
      format: { link: { uri: link.url } },
    });
    text += link.text;

    if (i < links.length - 1) {
      const sepStart = text.length;
      runs.push({
        startIndex: sepStart,
        format: {}, // no link for separator
      });
      text += SEPARATOR;
    }
  });

  return { text, runs };
}
