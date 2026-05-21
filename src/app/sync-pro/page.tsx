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

  function validateSection(s: SectionState): string | null {
    if (!s.name.trim()) return "Section needs a name.";
    if (!s.masterTabName.trim()) return "Master tab name is required.";
    if (!s.linkedSheets.length) return "Add at least one linked sheet.";
    for (let i = 0; i < s.linkedSheets.length; i++) {
      const l = s.linkedSheets[i];
      if (!l.url) return `Linked sheet #${i + 1}: URL is required.`;
      if (!l.nickname.trim())
        return `Linked sheet #${i + 1}: nickname is required.`;
      if (!l.tabName)
        return `Linked sheet #${i + 1}: pick a tab (use Fetch tabs first).`;
    }
    return null;
  }

  async function handleRunSection(idx: number) {
    const section = sections[idx];
    if (!section) return;
    const err = validateSection(section);
    if (err) {
      setSectionError(section.id, err);
      return;
    }
    setSectionError(section.id, null);
    setSectionRunning(section.id, true);
    try {
      const payload = { sections: [toConfigSection(section)] };
      const res = await fetch("/api/sync-pro", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        setSectionError(section.id, data.error || "Failed to run");
        return;
      }
      const sec = data.sections?.[0];
      if (sec) setSectionResult(section.id, sec as ProSectionResult);
    } catch {
      setSectionError(section.id, "Network error");
    } finally {
      setSectionRunning(section.id, false);
    }
  }

  async function handleRunAll(e: React.FormEvent) {
    e.preventDefault();
    let anyFailed = false;
    for (const s of sections) {
      const err = validateSection(s);
      if (err) {
        setSectionError(s.id, err);
        anyFailed = true;
      } else {
        setSectionError(s.id, null);
      }
    }
    if (anyFailed) return;
    setRunningSections(new Set(sections.map((s) => s.id)));
    try {
      const payload = { sections: sections.map(toConfigSection) };
      const res = await fetch("/api/sync-pro", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        const msg = data.error || "Failed to run";
        for (const s of sections) setSectionError(s.id, msg);
        return;
      }
      const newResults = new Map<string, ProSectionResult>();
      for (const r of (data.sections ?? []) as ProSectionResult[]) {
        newResults.set(r.sectionId, r);
      }
      setResults((prev) => {
        const merged = new Map(prev);
        for (const [k, v] of newResults) merged.set(k, v);
        return merged;
      });
    } catch {
      for (const s of sections) setSectionError(s.id, "Network error");
    } finally {
      setRunningSections(new Set());
    }
  }

  // Silence "unused" warnings on imports referenced only inside future tasks
  void handleClearSaved;
  void fetchHeadersForUrlTab;
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
        {sections.map((section, sIdx) => (
          <SectionCard
            key={section.id}
            section={section}
            index={sIdx}
            total={sections.length}
            running={runningSections.has(section.id)}
            error={errors.get(section.id) ?? null}
            result={results.get(section.id) ?? null}
            onChangeSection={(patch) => updateSection(sIdx, patch)}
            onRemove={() => removeSection(sIdx)}
            onAddLinkedSheet={() =>
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                sec.linkedSheets = [
                  ...sec.linkedSheets,
                  emptyLinkedSheet(),
                ];
                next[sIdx] = sec;
                return next;
              })
            }
            onRemoveLinkedSheet={(lIdx) =>
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                sec.linkedSheets =
                  sec.linkedSheets.length === 1
                    ? [emptyLinkedSheet()]
                    : sec.linkedSheets.filter((_, i) => i !== lIdx);
                next[sIdx] = sec;
                return next;
              })
            }
            onChangeLinkedSheet={(lIdx, patch) =>
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                const ls = sec.linkedSheets.slice();
                ls[lIdx] = { ...ls[lIdx], ...patch };
                sec.linkedSheets = ls;
                next[sIdx] = sec;
                return next;
              })
            }
            onFetchTabs={async (lIdx) => {
              const linked = sections[sIdx]?.linkedSheets[lIdx];
              if (!linked?.url) return;
              setSectionError(section.id, null);
              setSections((prev) => {
                const next = prev.slice();
                const sec = { ...next[sIdx] };
                const ls = sec.linkedSheets.slice();
                ls[lIdx] = {
                  ...ls[lIdx],
                  ui: { ...ls[lIdx].ui, loading: true },
                };
                sec.linkedSheets = ls;
                next[sIdx] = sec;
                return next;
              });
              try {
                const tabs = await fetchTabsForUrl(linked.url);
                setSections((prev) => {
                  const next = prev.slice();
                  const sec = { ...next[sIdx] };
                  const ls = sec.linkedSheets.slice();
                  ls[lIdx] = {
                    ...ls[lIdx],
                    ui: {
                      ...ls[lIdx].ui,
                      loadedTabs: tabs,
                      loading: false,
                    },
                  };
                  sec.linkedSheets = ls;
                  next[sIdx] = sec;
                  return next;
                });
              } catch (err) {
                setSectionError(
                  section.id,
                  err instanceof Error
                    ? err.message
                    : "Failed to fetch tabs"
                );
                setSections((prev) => {
                  const next = prev.slice();
                  const sec = { ...next[sIdx] };
                  const ls = sec.linkedSheets.slice();
                  ls[lIdx] = {
                    ...ls[lIdx],
                    ui: { ...ls[lIdx].ui, loading: false },
                  };
                  sec.linkedSheets = ls;
                  next[sIdx] = sec;
                  return next;
                });
              }
            }}
            onRunSection={() => handleRunSection(sIdx)}
          />
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

// ---------- SectionCard ----------

function SectionCard({
  section,
  index,
  total,
  running,
  error,
  result,
  onChangeSection,
  onRemove,
  onAddLinkedSheet,
  onRemoveLinkedSheet,
  onChangeLinkedSheet,
  onFetchTabs,
  onRunSection,
}: {
  section: SectionState;
  index: number;
  total: number;
  running: boolean;
  error: string | null;
  result: ProSectionResult | null;
  onChangeSection: (patch: Partial<SectionState>) => void;
  onRemove: () => void;
  onAddLinkedSheet: () => void;
  onRemoveLinkedSheet: (linkedIdx: number) => void;
  onChangeLinkedSheet: (
    linkedIdx: number,
    patch: Partial<LinkedSheetState>
  ) => void;
  onFetchTabs: (linkedIdx: number) => void;
  onRunSection: () => void;
}) {
  // When the user edits the section name, auto-update the master tab name
  // unless the user has already manually edited it away from the default.
  const defaultMasterTabName = `Pro: ${section.name}`;
  return (
    <div className="bg-card border border-border rounded-lg p-5 space-y-5">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div className="flex items-center gap-2 flex-1 min-w-0">
          <span className="text-xs text-muted uppercase tracking-wide">
            Section {index + 1}
          </span>
          <input
            type="text"
            value={section.name}
            onChange={(e) => {
              const newName = e.target.value;
              const masterTabName =
                section.masterTabName === defaultMasterTabName
                  ? `Pro: ${newName}`
                  : section.masterTabName;
              onChangeSection({ name: newName, masterTabName });
            }}
            placeholder="Section name"
            className="flex-1 min-w-0 border border-border rounded-md px-2 py-1 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </div>
        {total > 1 && (
          <button
            type="button"
            onClick={onRemove}
            className="text-xs text-danger hover:underline cursor-pointer whitespace-nowrap"
          >
            Remove section
          </button>
        )}
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <label className="flex flex-col gap-1 text-xs">
          <span className="text-muted uppercase tracking-wide">
            Master tab name
          </span>
          <input
            type="text"
            value={section.masterTabName}
            onChange={(e) =>
              onChangeSection({ masterTabName: e.target.value })
            }
            placeholder="Pro: Section 1"
            className="border border-border rounded-md px-2 py-1 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </label>
        <label className="flex items-center gap-2 text-sm self-end">
          <input
            type="checkbox"
            checked={section.writePresentIn}
            onChange={(e) =>
              onChangeSection({ writePresentIn: e.target.checked })
            }
          />
          Also write Present In column into each source
        </label>
      </div>

      {/* Propagate columns */}
      <div className="space-y-2">
        <p className="text-xs font-medium text-muted uppercase tracking-wide">
          Propagate columns ({section.propagateColumns.length})
        </p>
        <div className="flex flex-wrap gap-2 items-center">
          {section.propagateColumns.map((col, cIdx) => (
            <span
              key={`${col.name}-${cIdx}`}
              className="inline-flex items-center gap-1 bg-background border border-border rounded-md px-2 py-1 text-sm"
            >
              <code>{col.name}</code>
              <button
                type="button"
                onClick={() => {
                  const removedName = col.name;
                  const nextCols = section.propagateColumns.filter(
                    (_, i) => i !== cIdx
                  );
                  const nextLinked = section.linkedSheets.map((l) => {
                    const { [removedName]: _removed, ...rest } =
                      l.columnMapping;
                    void _removed;
                    return { ...l, columnMapping: rest };
                  });
                  onChangeSection({
                    propagateColumns: nextCols,
                    linkedSheets: nextLinked,
                  });
                }}
                className="text-danger hover:underline cursor-pointer"
                aria-label={`Remove ${col.name}`}
              >
                ×
              </button>
            </span>
          ))}
          <AddPropagateColumn
            onAdd={(name) => {
              const trimmed = name.trim();
              if (!trimmed) return;
              if (
                section.propagateColumns.some(
                  (c) => c.name.toLowerCase() === trimmed.toLowerCase()
                )
              ) {
                return;
              }
              onChangeSection({
                propagateColumns: [
                  ...section.propagateColumns,
                  { name: trimmed },
                ],
              });
            }}
          />
        </div>
        {section.propagateColumns.length === 0 && (
          <p className="text-xs text-muted">
            No propagate columns yet. Add one — e.g.{" "}
            <code className="bg-background px-1 rounded">Phone</code> — and
            map it to each linked sheet&apos;s actual header below. Pro will
            fill blanks across sheets for these columns.
          </p>
        )}
      </div>

      {/* Linked sheets list */}
      <div className="space-y-3">
        <p className="text-xs font-medium text-muted uppercase tracking-wide">
          Linked sheets ({section.linkedSheets.length})
        </p>
        {section.linkedSheets.map((linked, lIdx) => (
          <LinkedSheetRow
            key={lIdx}
            linked={linked}
            index={lIdx}
            total={section.linkedSheets.length}
            propagateColumns={section.propagateColumns}
            onChange={(patch) => onChangeLinkedSheet(lIdx, patch)}
            onRemove={() => onRemoveLinkedSheet(lIdx)}
            onFetch={() => onFetchTabs(lIdx)}
          />
        ))}
        <button
          type="button"
          onClick={onAddLinkedSheet}
          className="text-xs font-medium px-3 py-1.5 rounded-md border border-dashed border-border bg-background hover:bg-card transition-colors cursor-pointer"
        >
          + Add another linked sheet
        </button>
      </div>

      {/* Per-section Run + inline error */}
      <div className="border-t border-border pt-4 flex items-center justify-between gap-3 flex-wrap">
        {error ? (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2 flex-1 min-w-0">
            {error}
          </p>
        ) : (
          <span className="text-xs text-muted">
            Runs this section only. Other sections are untouched.
          </span>
        )}
        <button
          type="button"
          onClick={onRunSection}
          disabled={running}
          className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap"
        >
          {running ? "Running..." : "Run section"}
        </button>
      </div>

      {result && <InlineResult result={result} />}
    </div>
  );
}

// ---------- LinkedSheetRow ----------

function LinkedSheetRow({
  linked,
  index,
  total,
  propagateColumns,
  onChange,
  onRemove,
  onFetch,
}: {
  linked: LinkedSheetState;
  index: number;
  total: number;
  propagateColumns: ProPropagateColumn[];
  onChange: (patch: Partial<LinkedSheetState>) => void;
  onRemove: () => void;
  onFetch: () => void;
}) {
  return (
    <div className="border border-border rounded-md p-3 space-y-3 bg-background/50">
      <div className="flex items-center justify-between gap-2">
        <span className="text-xs text-muted">Sheet #{index + 1}</span>
        {total > 1 && (
          <button
            type="button"
            onClick={onRemove}
            className="text-xs text-danger hover:underline cursor-pointer"
          >
            Remove sheet
          </button>
        )}
      </div>

      <input
        type="text"
        value={linked.nickname}
        onChange={(e) => onChange({ nickname: e.target.value })}
        placeholder="Nickname (e.g. Newsletter)"
        className="w-full border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
      />

      <div className="flex flex-col sm:flex-row gap-2">
        <input
          type="url"
          value={linked.url}
          onChange={(e) =>
            onChange({
              url: e.target.value,
              tabName: "",
              columnMapping: {},
              ui: { loadedTabs: [], loadedHeaders: [], loading: false },
            })
          }
          placeholder="https://docs.google.com/spreadsheets/d/..."
          className="flex-1 min-w-0 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
        <button
          type="button"
          onClick={onFetch}
          disabled={!linked.url || linked.ui.loading}
          className="bg-foreground text-background px-3 py-2 rounded-md text-sm font-medium hover:opacity-80 disabled:opacity-50 transition-opacity cursor-pointer whitespace-nowrap"
        >
          {linked.ui.loading ? "Loading..." : "Fetch tabs"}
        </button>
      </div>

      {linked.ui.loadedTabs.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <label className="flex flex-col gap-1 text-xs">
            <span className="text-muted uppercase tracking-wide">Tab</span>
            <select
              value={linked.tabName}
              onChange={(e) =>
                onChange({ tabName: e.target.value, columnMapping: {} })
              }
              className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            >
              <option value="">— select —</option>
              {linked.ui.loadedTabs.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </label>
          <label className="flex flex-col gap-1 text-xs">
            <span className="text-muted uppercase tracking-wide">
              Email column
            </span>
            <input
              type="text"
              value={linked.emailColumn}
              onChange={(e) =>
                onChange({ emailColumn: e.target.value || "auto" })
              }
              placeholder="auto"
              className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            />
          </label>
        </div>
      )}

      {propagateColumns.length > 0 && linked.ui.loadedTabs.length > 0 && (
        <div className="space-y-2 border-t border-border pt-3">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Column mapping
          </p>
          <p className="text-[11px] text-muted">
            Type the header text from <code>{linked.tabName || "this tab"}</code>{" "}
            that should feed each section column. Leave blank to skip this
            sheet for that column.
          </p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            {propagateColumns.map((col) => (
              <label
                key={col.name}
                className="flex items-center gap-2 text-xs"
              >
                <span className="font-medium w-24 truncate">{col.name}</span>
                <input
                  type="text"
                  value={linked.columnMapping[col.name] ?? ""}
                  onChange={(e) => {
                    const v = e.target.value;
                    onChange({
                      columnMapping: {
                        ...linked.columnMapping,
                        [col.name]: v ? v : null,
                      },
                    });
                  }}
                  placeholder="(skip)"
                  className="flex-1 min-w-0 border border-border rounded-md px-2 py-1 text-xs bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
                />
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- AddPropagateColumn ----------

function AddPropagateColumn({ onAdd }: { onAdd: (name: string) => void }) {
  const [editing, setEditing] = useState(false);
  const [value, setValue] = useState("");
  if (!editing) {
    return (
      <button
        type="button"
        onClick={() => setEditing(true)}
        className="text-xs font-medium px-2 py-1 rounded-md border border-dashed border-border bg-background hover:bg-card transition-colors cursor-pointer"
      >
        + Add column
      </button>
    );
  }
  return (
    <span className="inline-flex items-center gap-1">
      <input
        autoFocus
        type="text"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            onAdd(value);
            setValue("");
            setEditing(false);
          } else if (e.key === "Escape") {
            setValue("");
            setEditing(false);
          }
        }}
        onBlur={() => {
          if (value.trim()) onAdd(value);
          setValue("");
          setEditing(false);
        }}
        placeholder="Phone"
        className="border border-border rounded-md px-2 py-1 text-xs bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary w-24"
      />
    </span>
  );
}

// ---------- InlineResult ----------

function InlineResult({ result }: { result: ProSectionResult }) {
  const failed = !!result.error;
  const visibleConflicts = result.conflicts.slice(0, 10);
  const hiddenConflicts = Math.max(
    0,
    result.conflicts.length - visibleConflicts.length
  );
  return (
    <div
      className={`border-t pt-4 ${
        failed ? "border-danger/30" : "border-border"
      }`}
    >
      <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2 mb-3">
        <h3 className="font-medium text-sm">Last run</h3>
        {result.masterSpreadsheetUrl && (
          <a
            href={result.masterSpreadsheetUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="text-sm text-primary hover:underline font-medium"
          >
            Open {result.masterTabName} →
          </a>
        )}
      </div>

      {failed && (
        <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2 mb-3">
          {result.error}
        </p>
      )}

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-sm">
        <Stat label="Unique emails" value={result.totalUniqueEmails} />
        <Stat
          label="Cells filled"
          value={result.totalCellsFilled}
          accent="success"
        />
        <Stat
          label="Conflicts"
          value={result.totalConflicts}
          accent={result.totalConflicts > 0 ? "danger" : "muted"}
        />
        <Stat
          label="Present In"
          value={result.presentInWritten ? 1 : 0}
          accent="muted"
          hint={result.presentInWritten ? "written" : "skipped"}
        />
      </div>

      {result.columnStats.length > 0 && (
        <div className="mt-3 space-y-1">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Per column
          </p>
          {result.columnStats.map((c) => (
            <div
              key={c.name}
              className="flex items-center justify-between text-xs py-1 border-b border-border last:border-0"
            >
              <code className="font-medium">{c.name}</code>
              <span className="text-muted">
                {c.cellsFilled} filled · {c.conflicts} conflict
                {c.conflicts === 1 ? "" : "s"} · {c.skippedSheets} skipped
                sheet{c.skippedSheets === 1 ? "" : "s"}
              </span>
            </div>
          ))}
        </div>
      )}

      {result.linkedSheets.length > 0 && (
        <div className="mt-3 space-y-1">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Per sheet
          </p>
          {result.linkedSheets.map((t, i) => (
            <div
              key={i}
              className="flex flex-col gap-0.5 text-xs py-1 border-b border-border last:border-0"
            >
              <div className="flex items-center justify-between gap-2">
                <span className="font-medium">{t.nickname}</span>
                {t.error ? (
                  <span className="text-danger">{t.error}</span>
                ) : (
                  <span className="text-muted">
                    {t.rowsRead} rows · {t.emailsFound} emails ·{" "}
                    {t.cellsFilled} cells filled
                  </span>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {visibleConflicts.length > 0 && (
        <div className="mt-3 space-y-1">
          <p className="text-xs font-medium text-muted uppercase tracking-wide">
            Conflicts (first {visibleConflicts.length})
          </p>
          <div className="space-y-1">
            {visibleConflicts.map((c, i) => (
              <div
                key={i}
                className="text-xs text-muted py-1 border-b border-border last:border-0"
              >
                <code className="font-medium">{c.email}</code> ·{" "}
                <code>{c.column}</code>:{" "}
                {c.values
                  .map((v) => `${v.nickname}="${v.value}"`)
                  .join(" vs ")}
              </div>
            ))}
          </div>
          {hiddenConflicts > 0 && (
            <p className="text-[11px] text-muted">
              … and {hiddenConflicts} more (open the master tab to see all).
            </p>
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Stat ----------

function Stat({
  label,
  value,
  accent,
  hint,
}: {
  label: string;
  value: number;
  accent?: "success" | "danger" | "muted";
  hint?: string;
}) {
  const valueClass =
    accent === "success"
      ? "text-success"
      : accent === "danger"
        ? "text-danger"
        : accent === "muted"
          ? "text-muted"
          : "text-foreground";
  return (
    <div className="bg-background border border-border rounded-md px-3 py-2">
      <p className="text-xs text-muted">{label}</p>
      <p className={`text-lg font-semibold ${valueClass}`}>{value}</p>
      {hint && <p className="text-[10px] text-muted mt-0.5">{hint}</p>}
    </div>
  );
}
