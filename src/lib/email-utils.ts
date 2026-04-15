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

export interface NameColumns {
  fullName?: number;
  firstName?: number;
  lastName?: number;
}

export function findNameColumns(headers: string[]): NameColumns {
  const normalized = headers.map((h) => h?.toString().toLowerCase().trim() ?? "");

  const fullName = normalized.findIndex(
    (h) =>
      h === "name" ||
      h === "full name" ||
      h === "fullname" ||
      h === "full_name"
  );
  const firstName = normalized.findIndex(
    (h) =>
      h === "first name" ||
      h === "firstname" ||
      h === "first_name" ||
      h === "first" ||
      h === "given name"
  );
  const lastName = normalized.findIndex(
    (h) =>
      h === "last name" ||
      h === "lastname" ||
      h === "last_name" ||
      h === "last" ||
      h === "surname" ||
      h === "family name"
  );

  return {
    fullName: fullName >= 0 ? fullName : undefined,
    firstName: firstName >= 0 ? firstName : undefined,
    lastName: lastName >= 0 ? lastName : undefined,
  };
}
