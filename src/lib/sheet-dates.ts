/**
 * Shared date helpers for Sheets-aware features (renewal-sync, consolidator).
 *
 * `parseFlexibleDate` accepts the messy real-world date strings users tend
 * to type — Italian DD/MM, ISO YYYY-MM-DD, US MM/DD, two-digit years — and
 * returns a JS Date or null. It's lenient on purpose; the caller decides
 * what to do with unparseable input.
 *
 * `toSheetsDateSerial` converts a JS Date to the integer "serial number"
 * Google Sheets uses internally (days since 1899-12-30, anchored to absorb
 * the historical Lotus 1-2-3 leap-year-1900 bug). UTC math is used so the
 * server's local timezone never shifts the result by ±1 day.
 */

const SHEETS_DATE_EPOCH_UTC_MS = Date.UTC(1899, 11, 30);

export function toSheetsDateSerial(d: Date): number {
  const y = d.getFullYear();
  const m = d.getMonth();
  const day = d.getDate();
  return Math.floor(
    (Date.UTC(y, m, day) - SHEETS_DATE_EPOCH_UTC_MS) / 86400000
  );
}

export function parseFlexibleDate(raw: string): Date | null {
  const str = raw.trim();
  if (!str) return null;

  // ISO YYYY-MM-DD (or YYYY/MM/DD or YYYY.MM.DD)
  const iso = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (iso) {
    const d = new Date(
      parseInt(iso[1]),
      parseInt(iso[2]) - 1,
      parseInt(iso[3])
    );
    if (!isNaN(d.getTime())) return d;
  }

  // DD/MM/YYYY (European default — data is Italian)
  // Falls back to MM/DD/YYYY if first part > 12 unambiguously.
  const parts = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (parts) {
    let a = parseInt(parts[1]);
    let b = parseInt(parts[2]);
    let year = parseInt(parts[3]);
    if (year < 100) year += 2000;

    if (a <= 12 && b > 12) {
      [a, b] = [b, a];
    }
    const d = new Date(year, b - 1, a);
    if (!isNaN(d.getTime())) return d;
  }

  const fallback = new Date(str);
  if (!isNaN(fallback.getTime())) return fallback;

  return null;
}
