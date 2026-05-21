// src/lib/sync-pro-engine.ts
import { getSheetsClient } from "./google-auth";
import { withRetry } from "./retry";
import {
  normalizeEmail,
  findEmailColumnIndex,
  findNameColumns,
  columnIndexToLetter,
} from "./email-utils";
import { extractSpreadsheetId } from "./url-parser";
import { writePresentInColumn } from "./present-in-writer";
import { sheets_v4 } from "googleapis";
import type {
  ProBatchResult,
  ProColumnStats,
  ProConflict,
  ProLinkedSheet,
  ProSection,
  ProSectionResult,
  ProTabResult,
} from "./sync-pro-types";

/**
 * Sheet Syncer Pro engine.
 *
 * Each section is independent. Within a section:
 *   1. Read every linked sheet's email + name + mapped propagate columns.
 *   2. Cross-reference emails to build {email: which sheets contain it}.
 *   3. Propagation pass: for each logical propagate column, fill blanks
 *      from non-blank values. Never overwrite. Surface conflicts when
 *      multiple sheets have different non-blank values.
 *   4. Apply per-sheet writes (one batchUpdate per sheet).
 *   5. Write Present In column into each source sheet (optional).
 *   6. Write Master ✅/❌ tab into the master sheet.
 */

// Lowered from 5 to stay under Google Sheets' 60 read/min per-user quota,
// matching the main sync engine.
const MAX_CONCURRENCY = 3;

const HEADER_BG = { red: 0.93, green: 0.93, blue: 0.96 };

async function withConcurrencyLimit<T>(
  tasks: (() => Promise<T>)[],
  limit: number
): Promise<T[]> {
  const results: T[] = [];
  const executing = new Set<Promise<void>>();
  for (const task of tasks) {
    const p = task().then((result) => {
      results.push(result);
      executing.delete(p);
    });
    executing.add(p);
    if (executing.size >= limit) {
      await Promise.race(executing);
    }
  }
  await Promise.all(executing);
  return results;
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    return value.toLocaleString("en-US", {
      useGrouping: false,
      maximumFractionDigits: 20,
    });
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim();
}

/**
 * Resolve a 1-based column letter (or "auto") to a 0-based column index.
 * `auto` returns -1 and the caller should fall back to header detection.
 */
function letterToIndex(letter: string): number {
  if (!letter || letter.toLowerCase() === "auto") return -1;
  const upper = letter.toUpperCase();
  let idx = 0;
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    idx = idx * 26 + (code - 64);
  }
  return idx - 1;
}

/**
 * Merge names across sheets for the Master tab.
 *
 * DUPLICATED FROM src/lib/sync-engine.ts so the basic sync stays 100%
 * untouched (per spec / user direction). If a bug is found in either
 * copy, fix both.
 *
 * Strategy:
 *  1. Count every name occurrence (case-insensitive).
 *  2. If name A is a word-prefix of name B (e.g. "Laura" ⊂ "Laura
 *     Pegoraro"), fold A's count into B so the more complete name wins.
 *  3. Pick winner by mergedCount → words.length → firstSeenIdx.
 */
function mergeNamesForEmail(namesByEmail: Map<string, string[]>): Map<string, string> {
  const out = new Map<string, string>();
  for (const [email, names] of namesByEmail) {
    interface Candidate {
      name: string;
      words: string[];
      count: number;
      mergedCount: number;
      firstSeenIdx: number;
    }
    const candidates = new Map<string, Candidate>();
    let nextIdx = 0;
    for (const name of names) {
      const key = name.toLowerCase();
      const existing = candidates.get(key);
      if (existing) {
        existing.count++;
        existing.mergedCount++;
      } else {
        candidates.set(key, {
          name,
          words: key.split(/\s+/).filter(Boolean),
          count: 1,
          mergedCount: 1,
          firstSeenIdx: nextIdx++,
        });
      }
    }
    if (candidates.size === 0) continue;

    const list = Array.from(candidates.values());
    for (const a of list) {
      for (const b of list) {
        if (a === b) continue;
        if (b.words.length <= a.words.length) continue;
        const isPrefix = a.words.every((w, i) => b.words[i] === w);
        if (isPrefix) b.mergedCount += a.count;
      }
    }
    list.sort((a, b) => {
      if (b.mergedCount !== a.mergedCount) return b.mergedCount - a.mergedCount;
      if (b.words.length !== a.words.length) return b.words.length - a.words.length;
      return a.firstSeenIdx - b.firstSeenIdx;
    });
    out.set(email, list[0].name);
  }
  return out;
}

// ===========================================================================
// Per-sheet read + section run will be implemented in subsequent tasks.
// ===========================================================================

export async function runSyncProSection(
  refreshToken: string,
  masterSheetId: string,
  section: ProSection
): Promise<ProSectionResult> {
  // Placeholder — implemented incrementally over Tasks 5-7.
  const result: ProSectionResult = {
    sectionId: section.id,
    sectionName: section.name,
    masterSpreadsheetUrl: "",
    masterTabName: section.masterTabName || `Pro: ${section.name}`,
    totalUniqueEmails: 0,
    totalCellsFilled: 0,
    totalConflicts: 0,
    linkedSheets: [],
    columnStats: [],
    conflicts: [],
    presentInWritten: false,
    error: "runSyncProSection: not yet implemented",
  };
  // Silence unused-import warnings while skeleton compiles.
  void refreshToken;
  void masterSheetId;
  void getSheetsClient;
  void withRetry;
  void withConcurrencyLimit;
  void clean;
  void letterToIndex;
  void normalizeEmail;
  void findEmailColumnIndex;
  void findNameColumns;
  void columnIndexToLetter;
  void extractSpreadsheetId;
  void writePresentInColumn;
  void mergeNamesForEmail;
  void HEADER_BG;
  void MAX_CONCURRENCY;
  const _unused: sheets_v4.Sheets | null = null;
  void _unused;
  void ({} as ProLinkedSheet);
  void ({} as ProTabResult);
  void ({} as ProColumnStats);
  void ({} as ProConflict);
  return result;
}

export async function runSyncProBatch(
  refreshToken: string,
  masterSheetId: string,
  sections: ProSection[]
): Promise<ProBatchResult> {
  const results: ProSectionResult[] = [];
  for (const section of sections) {
    try {
      results.push(await runSyncProSection(refreshToken, masterSheetId, section));
    } catch (err) {
      results.push({
        sectionId: section.id,
        sectionName: section.name,
        masterSpreadsheetUrl: "",
        masterTabName: section.masterTabName || `Pro: ${section.name}`,
        totalUniqueEmails: 0,
        totalCellsFilled: 0,
        totalConflicts: 0,
        linkedSheets: [],
        columnStats: [],
        conflicts: [],
        presentInWritten: false,
        error: err instanceof Error ? err.message : "Section failed",
      });
    }
  }
  return { sections: results };
}
