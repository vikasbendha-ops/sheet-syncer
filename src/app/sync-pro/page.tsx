// src/app/sync-pro/page.tsx
"use client";

import { useCallback, useEffect, useState } from "react";
import type {
  ProConflict,
  ProColumnStats,
  ProLinkedSheet,
  ProPropagateColumn,
  ProSection,
  ProSectionResult,
  ProTabResult,
} from "@/lib/sync-pro-types";

// ---------- Local UI state shapes ----------

interface SourceTabState {
  loadedTabs: string[]; // tabs fetched from Google for the URL
  loadedHeaders: string[]; // headers for the picked tab (after Fetch tabs)
  loading: boolean;
}

interface LinkedSheetState extends ProLinkedSheet {
  ui: SourceTabState;
}

interface SectionState
  extends Omit<ProSection, "linkedSheets"> {
  linkedSheets: LinkedSheetState[];
}

function newId(): string {
  return `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function emptyLinkedSheet(): LinkedSheetState {
  return {
    url: "",
    nickname: "",
    tabName: "",
    emailColumn: "auto",
    columnMapping: {},
    ui: { loadedTabs: [], loadedHeaders: [], loading: false },
  };
}

function makeSection(index: number): SectionState {
  const name = `Section ${index + 1}`;
  return {
    id: newId(),
    name,
    masterTabName: `Pro: ${name}`,
    writePresentIn: true,
    propagateColumns: [],
    linkedSheets: [emptyLinkedSheet()],
  };
}

function toConfigSection(s: SectionState): ProSection {
  return {
    id: s.id,
    name: s.name,
    masterTabName: s.masterTabName,
    writePresentIn: s.writePresentIn,
    propagateColumns: s.propagateColumns,
    linkedSheets: s.linkedSheets.map(({ ui: _ui, ...rest }) => {
      void _ui;
      return rest;
    }),
  };
}

async function fetchTabsForUrl(url: string): Promise<string[]> {
  const res = await fetch(`/api/sheets/tabs?url=${encodeURIComponent(url)}`);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to fetch tabs");
  return data.tabs ?? [];
}

async function fetchHeadersForUrlTab(
  url: string,
  tabName: string
): Promise<string[]> {
  // Reuses /api/sheets/tabs? — not quite, headers aren't exposed by any
  // existing route. We add a small client-side fallback: fetch headers via
  // /api/sheets/headers when added; until then keep loadedHeaders empty
  // and let the UI fall back to a free-text mapping input.
  void url;
  void tabName;
  return [];
}

export default function SyncProPage() {
  const [sections, setSections] = useState<SectionState[]>([makeSection(0)]);
  const [results, setResults] = useState<Map<string, ProSectionResult>>(
    new Map()
  );
  const [errors, setErrors] = useState<Map<string, string>>(new Map());
  const [runningSections, setRunningSections] = useState<Set<string>>(
    new Set()
  );

  const [hydrated, setHydrated] = useState(false);
  const [hasSaved, setHasSaved] = useState(false);
  const [noMasterSheet, setNoMasterSheet] = useState(false);

  // ---------- Per-section state helpers ----------

  function setSectionRunning(id: string, on: boolean) {
    setRunningSections((prev) => {
      const next = new Set(prev);
      if (on) next.add(id);
      else next.delete(id);
      return next;
    });
  }

  function setSectionError(id: string, msg: string | null) {
    setErrors((prev) => {
      const next = new Map(prev);
      if (msg) next.set(id, msg);
      else next.delete(id);
      return next;
    });
  }

  function setSectionResult(id: string, result: ProSectionResult | null) {
    setResults((prev) => {
      const next = new Map(prev);
      if (result) next.set(id, result);
      else next.delete(id);
      return next;
    });
  }

  // ---------- Hydrate from server on mount ----------

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/sync-pro/config");
        if (res.status === 400) {
          const data = await res.json();
          if (data.code === "no_master_sheet") setNoMasterSheet(true);
        } else if (res.ok) {
          const data = (await res.json()) as { sections: ProSection[] };
          if (data.sections && data.sections.length > 0) {
            setHasSaved(true);
            const rehydrated: SectionState[] = data.sections.map(
              (s, idx): SectionState => ({
                id: s.id || newId(),
                name: s.name || `Section ${idx + 1}`,
                masterTabName:
                  s.masterTabName || `Pro: ${s.name || `Section ${idx + 1}`}`,
                writePresentIn:
                  typeof s.writePresentIn === "boolean"
                    ? s.writePresentIn
                    : true,
                propagateColumns: s.propagateColumns ?? [],
                linkedSheets:
                  s.linkedSheets && s.linkedSheets.length > 0
                    ? s.linkedSheets.map((l) => ({
                        ...l,
                        ui: {
                          loadedTabs: [],
                          loadedHeaders: [],
                          loading: false,
                        },
                      }))
                    : [emptyLinkedSheet()],
              })
            );
            setSections(rehydrated);

            // Eagerly load tab lists for any linked sheet that already has a URL
            await Promise.all(
              rehydrated.map(async (sec, sIdx) => {
                await Promise.all(
                  sec.linkedSheets.map(async (linked, lIdx) => {
                    if (!linked.url) return;
                    try {
                      const tabs = await fetchTabsForUrl(linked.url);
                      setSections((prev) => {
                        const next = prev.slice();
                        const sec2 = { ...next[sIdx] };
                        const ls = sec2.linkedSheets.slice();
                        ls[lIdx] = {
                          ...ls[lIdx],
                          ui: {
                            ...ls[lIdx].ui,
                            loadedTabs: tabs,
                            loading: false,
                          },
                        };
                        sec2.linkedSheets = ls;
                        next[sIdx] = sec2;
                        return next;
                      });
                    } catch {
                      // ignore — user can refetch
                    }
                  })
                );
              })
            );
          }
        }
      } catch {
        // empty form
      }
      setHydrated(true);
    })();
  }, []);

  // ---------- Persist to server on changes (debounced) ----------

  useEffect(() => {
    if (!hydrated || noMasterSheet) return;
    const hasAnything = sections.some(
      (s) =>
        s.name ||
        s.linkedSheets.some((l) => l.url || l.nickname) ||
        s.propagateColumns.length > 0
    );
    if (!hasAnything) return;

    const handle = setTimeout(() => {
      const payload = { sections: sections.map(toConfigSection) };
      fetch("/api/sync-pro/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      })
        .then((r) => {
          if (r.ok) setHasSaved(true);
        })
        .catch(() => {
          // ignore
        });
    }, 800);
    return () => clearTimeout(handle);
  }, [sections, hydrated, noMasterSheet]);

  // ---------- Mutators ----------

  const updateSection = useCallback(
    (idx: number, patch: Partial<SectionState>) => {
      setSections((prev) => {
        const next = prev.slice();
        next[idx] = { ...next[idx], ...patch };
        return next;
      });
    },
    []
  );

  function addSection() {
    setSections((prev) => [...prev, makeSection(prev.length)]);
  }

  function removeSection(idx: number) {
    setSections((prev) => {
      if (prev.length === 1) return [makeSection(0)];
      const removed = prev[idx];
      if (removed) {
        setSectionResult(removed.id, null);
        setSectionError(removed.id, null);
        setSectionRunning(removed.id, false);
      }
      return prev.filter((_, i) => i !== idx);
    });
  }

  async function handleClearSaved() {
    setErrors(new Map());
    setResults(new Map());
    try {
      await fetch("/api/sync-pro/config", { method: "DELETE" });
    } catch {
      // best-effort
    }
    setSections([makeSection(0)]);
    setHasSaved(false);
  }

  // Stubs filled in by Task 12 / 13.
  function handleRunSection(_idx: number) {
    void _idx;
  }
  async function handleRunAll(e: React.FormEvent) {
    e.preventDefault();
  }

  // Silence "unused" warnings on imports referenced only inside future tasks
  void updateSection;
  void addSection;
  void removeSection;
  void handleRunSection;
  void handleClearSaved;
  void fetchHeadersForUrlTab;
  void results;
  void errors;
  void ({} as ProConflict);
  void ({} as ProColumnStats);
  void ({} as ProPropagateColumn);
  void ({} as ProTabResult);

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Sync Pro
        </h1>
        <p className="text-muted text-sm mt-1">
          Per-section, multi-sheet Present In with column propagation.
          Sections are independent — each runs on its own button. (UI under
          construction.)
        </p>
      </div>

      {noMasterSheet && (
        <div className="bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          Set a master sheet on the Sheets page to use Sync Pro.
        </div>
      )}

      {hasSaved && !noMasterSheet && (
        <div className="flex items-center justify-between gap-3 bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          <span>
            Configuration saved to your master sheet — it&apos;ll auto-load
            next time.
          </span>
          <button
            type="button"
            onClick={handleClearSaved}
            className="text-danger hover:underline font-medium cursor-pointer whitespace-nowrap"
          >
            Clear saved
          </button>
        </div>
      )}

      <form onSubmit={handleRunAll} className="space-y-5">
        {sections.map((s, idx) => (
          <div
            key={s.id}
            className="bg-card border border-border rounded-lg p-5 text-sm text-muted"
          >
            <span className="font-medium text-foreground">{s.name}</span>{" "}
            (section {idx + 1} placeholder — full UI lands in Task 11)
          </div>
        ))}

        <div className="flex items-center justify-between gap-3 flex-wrap">
          <button
            type="button"
            onClick={addSection}
            className="text-sm font-medium px-3 py-1.5 rounded-md border border-border bg-background hover:bg-card transition-colors cursor-pointer"
          >
            + Add section
          </button>
          {sections.length > 1 && (
            <button
              type="submit"
              disabled={runningSections.size > 0}
              className="bg-primary text-white px-5 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
            >
              {runningSections.size > 0
                ? "Running..."
                : `Run all ${sections.length} sections`}
            </button>
          )}
        </div>
      </form>
    </div>
  );
}
