const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function normalizeEmail(raw: string): string | null {
  const trimmed = raw.trim().toLowerCase();
  if (!trimmed || !EMAIL_REGEX.test(trimmed)) {
    return null;
  }
  return trimmed;
}

export function columnIndexToLetter(index: number): string {
  let letter = "";
  let i = index;
  while (i >= 0) {
    letter = String.fromCharCode((i % 26) + 65) + letter;
    i = Math.floor(i / 26) - 1;
  }
  return letter;
}

export function findEmailColumnIndex(headers: string[]): number {
  const normalized = headers.map((h) => h?.toString().toLowerCase().trim() ?? "");
  const exactMatch = normalized.findIndex(
    (h) => h === "email" || h === "e-mail" || h === "email address" || h === "emailaddress"
  );
  if (exactMatch !== -1) return exactMatch;

  const partialMatch = normalized.findIndex((h) => h.includes("email"));
  if (partialMatch !== -1) return partialMatch;

  return -1;
}
