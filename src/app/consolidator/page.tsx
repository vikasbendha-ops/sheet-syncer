"use client";

import { useState, useEffect, useCallback } from "react";

interface TabResult {
  tabName: string;
  totalRows: number;
  rowsWithEmail: number;
  rowsWithoutEmail: number;
  columnsContributed: number;
  error?: string;
}

interface SourceResult {
  spreadsheetUrl: string;
  tabs: TabResult[];
  error?: string;
}

interface SectionResult {
  sectionId: string;
  sectionName: string;
  outputSpreadsheetUrl: string;
  outputTabName: string;
  totalSourceRows: number;
  totalOutputRows: number;
  rowsWithoutEmail: number;
  totalColumns: number;
  emailDuplicateValues: number;
  emailDuplicateCells: number;
  phoneDuplicateValues: number;
  phoneDuplicateCells: number;
  renewalRulesInstalled: boolean;
  sources: SourceResult[];
  error?: string;
}

interface SourceState {
  url: string;
  tabs: string[];
  selected: string[];
  loading: boolean;
}

interface SectionState {
  id: string;
  name: string;
  sources: SourceState[];
  outputUrl: string;
  outputTabName: string;
}

interface SavedSource {
  url: string;
  tabs: string[];
}

interface SavedSection {
  id: string;
  name: string;
  sources: SavedSource[];
  outputUrl: string;
  outputTabName: string;
}

function makeSourceState(url = "", selected: string[] = []): SourceState {
  return { url, tabs: [], selected, loading: false };
}

function makeSection(index: number): SectionState {
  return {
    id: `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    name: `Section ${index + 1}`,
    sources: [makeSourceState()],
    outputUrl: "",
    outputTabName: index === 0 ? "Consolidated" : `Consolidated ${index + 1}`,
  };
}

async function fetchTabsForUrl(url: string): Promise<string[]> {
  const res = await fetch(`/api/sheets/tabs?url=${encodeURIComponent(url)}`);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to fetch tabs");
  return data.tabs ?? [];
}

function serializeSection(s: SectionState) {
  return {
    id: s.id,
    name: s.name,
    outputUrl: s.outputUrl,
    outputTabName: s.outputTabName,
    sources: s.sources
      .filter((src) => src.url && src.selected.length > 0)
      .map((src) => ({ url: src.url, tabs: src.selected })),
  };
}

function validateSection(s: SectionState): string | null {
  if (!s.outputUrl) {
    return `Output spreadsheet URL is required.`;
  }
  if (!s.outputTabName.trim()) {
    return `Output tab name is required.`;
  }
  const hasAnyTabs = s.sources.some(
    (src) => src.url && src.selected.length > 0
  );
  if (!hasAnyTabs) {
    return `At least one source spreadsheet with tabs selected is required.`;
  }
  return null;
}

export default function ConsolidatorPage() {
  const [sections, setSections] = useState<SectionState[]>([makeSection(0)]);

  // Per-section runtime state (keyed by section.id).
  // Sections are independent — running one doesn't affect the others'
  // result/error/loading state.
  const [results, setResults] = useState<Map<string, SectionResult>>(
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

  function setSectionRunning(id: string, running: boolean) {
    setRunningSections((prev) => {
      const next = new Set(prev);
      if (running) next.add(id);
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

  function setSectionResult(id: string, result: SectionResult | null) {
    setResults((prev) => {
      const next = new Map(prev);
      if (result) next.set(id, result);
      else next.delete(id);
      return next;
    });
  }

  // Hydrate from server on mount
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/consolidator/config");
        if (res.status === 400) {
          const data = await res.json();
          if (data.code === "no_master_sheet") setNoMasterSheet(true);
        } else if (res.ok) {
          const data = (await res.json()) as { sections: SavedSection[] };
          if (data.sections && data.sections.length > 0) {
            setHasSaved(true);
            const rehydrated: SectionState[] = data.sections.map(
              (s, idx): SectionState => ({
                id: s.id || makeSection(idx).id,
                name: s.name || `Section ${idx + 1}`,
                sources:
                  s.sources && s.sources.length > 0
                    ? s.sources.map((src) =>
                        makeSourceState(src.url || "", src.tabs || [])
                      )
                    : [makeSourceState()],
                outputUrl: s.outputUrl || "",
                outputTabName: s.outputTabName || "Consolidated",
              })
            );
            setSections(rehydrated);

            // Fetch tabs for every source that has a URL
            await Promise.all(
              rehydrated.map(async (section, sIdx) => {
                await Promise.all(
                  section.sources.map(async (src, srcIdx) => {
                    if (!src.url) return;
                    try {
                      const tabs = await fetchTabsForUrl(src.url);
                      setSections((prev) => {
                        const next = prev.slice();
                        const sec = { ...next[sIdx] };
                        const srcs = sec.sources.slice();
                        const wantedSelected = srcs[srcIdx]?.selected ?? [];
                        srcs[srcIdx] = {
                          url: src.url,
                          tabs,
                          selected: wantedSelected.filter((t) =>
                            tabs.includes(t)
                          ),
                          loading: false,
                        };
                        sec.sources = srcs;
                        next[sIdx] = sec;
                        return next;
                      });
                    } catch {
                      // leave source with no tabs; user can refetch
                    }
                  })
                );
              })
            );
          }
        }
      } catch {
        // ignore — empty form
      }
      setHydrated(true);
    })();
  }, []);

  // Persist to server on changes (debounced)
  useEffect(() => {
    if (!hydrated || noMasterSheet) return;
    const hasAnything = sections.some(
      (s) =>
        s.outputUrl ||
        s.outputTabName !== "Consolidated" ||
        s.sources.some((src) => src.url || src.selected.length > 0)
    );
    if (!hasAnything) return;

    const handle = setTimeout(() => {
      const payload = {
        sections: sections.map((s) => ({
          id: s.id,
          name: s.name,
          outputUrl: s.outputUrl,
          outputTabName: s.outputTabName,
          sources: s.sources.map((src) => ({
            url: src.url,
            tabs: src.selected,
          })),
        })),
      };
      fetch("/api/consolidator/config", {
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

  // ---------- Section / source mutators ----------

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

  const updateSource = useCallback(
    (sIdx: number, srcIdx: number, patch: Partial<SourceState>) => {
      setSections((prev) => {
        const next = prev.slice();
        const sec = { ...next[sIdx] };
        const srcs = sec.sources.slice();
        srcs[srcIdx] = { ...srcs[srcIdx], ...patch };
        sec.sources = srcs;
        next[sIdx] = sec;
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
      if (prev.length === 1) {
        return [makeSection(0)];
      }
      // Clean up runtime state for the removed section
      const removed = prev[idx];
      if (removed) {
        setSectionResult(removed.id, null);
        setSectionError(removed.id, null);
        setSectionRunning(removed.id, false);
      }
      return prev.filter((_, i) => i !== idx);
    });
  }

  function addSource(sectionIdx: number) {
    setSections((prev) => {
      const next = prev.slice();
      const sec = { ...next[sectionIdx] };
      sec.sources = [...sec.sources, makeSourceState()];
      next[sectionIdx] = sec;
      return next;
    });
  }

  function removeSource(sectionIdx: number, sourceIdx: number) {
    setSections((prev) => {
      const next = prev.slice();
      const sec = { ...next[sectionIdx] };
      sec.sources =
        sec.sources.length === 1
          ? [makeSourceState()]
          : sec.sources.filter((_, i) => i !== sourceIdx);
      next[sectionIdx] = sec;
      return next;
    });
  }

  async function handleFetchTabs(sectionIdx: number, sourceIdx: number) {
    const src = sections[sectionIdx]?.sources[sourceIdx];
    if (!src?.url) return;
    const section = sections[sectionIdx];
    if (section) setSectionError(section.id, null);
    updateSource(sectionIdx, sourceIdx, { loading: true });
    try {
      const tabs = await fetchTabsForUrl(src.url);
      updateSource(sectionIdx, sourceIdx, {
        tabs,
        selected: tabs, // default-select all on fresh fetch
        loading: false,
      });
    } catch (err) {
      if (section) {
        setSectionError(
          section.id,
          err instanceof Error ? err.message : "Failed to fetch tabs"
        );
      }
      updateSource(sectionIdx, sourceIdx, {
        tabs: [],
        selected: [],
        loading: false,
      });
    }
  }

  function toggleTab(sectionIdx: number, sourceIdx: number, tab: string) {
    setSections((prev) => {
      const next = prev.slice();
      const sec = { ...next[sectionIdx] };
      const srcs = sec.sources.slice();
      const src = srcs[sourceIdx];
      srcs[sourceIdx] = {
        ...src,
        selected: src.selected.includes(tab)
          ? src.selected.filter((t) => t !== tab)
          : [...src.selected, tab],
      };
      sec.sources = srcs;
      next[sectionIdx] = sec;
      return next;
    });
  }

  function selectAllTabs(sectionIdx: number, sourceIdx: number, all: boolean) {
    setSections((prev) => {
      const next = prev.slice();
      const sec = { ...next[sectionIdx] };
      const srcs = sec.sources.slice();
      const src = srcs[sourceIdx];
      srcs[sourceIdx] = {
        ...src,
        selected: all ? src.tabs.slice() : [],
      };
      sec.sources = srcs;
      next[sectionIdx] = sec;
      return next;
    });
  }

  async function handleClearSaved() {
    setErrors(new Map());
    setResults(new Map());
    try {
      await fetch("/api/consolidator/config", { method: "DELETE" });
    } catch {
      // best-effort
    }
    setSections([makeSection(0)]);
    setHasSaved(false);
  }

  // ---------- Run handlers ----------

  /** Run a single section's consolidation. Independent from other sections. */
  async function handleRunSection(idx: number) {
    const section = sections[idx];
    if (!section) return;
    const validationError = validateSection(section);
    if (validationError) {
      setSectionError(section.id, validationError);
      return;
    }
    setSectionError(section.id, null);
    setSectionRunning(section.id, true);
    try {
      const payload = { sections: [serializeSection(section)] };
      const res = await fetch("/api/consolidator", {
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
      if (sec) {
        setSectionResult(section.id, sec);
      }
    } catch {
      setSectionError(section.id, "Network error");
    } finally {
      setSectionRunning(section.id, false);
    }
  }

  /** Convenience: run every section sequentially in one batch call. */
  async function handleRunAll(e: React.FormEvent) {
    e.preventDefault();

    // Validate each section. If any fails, surface the error inline on that
    // section's card; abort the batch run.
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
      const payload = { sections: sections.map(serializeSection) };
      const res = await fetch("/api/consolidator", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        const msg = data.error || "Failed to run";
        for (const s of sections) {
          setSectionError(s.id, msg);
        }
        return;
      }
      const newResults = new Map<string, SectionResult>();
      for (const r of (data.sections ?? []) as SectionResult[]) {
        newResults.set(r.sectionId, r);
      }
      setResults((prev) => {
        const merged = new Map(prev);
        for (const [k, v] of newResults) merged.set(k, v);
        return merged;
      });
    } catch {
      for (const s of sections) {
        setSectionError(s.id, "Network error");
      }
    } finally {
      setRunningSections(new Set());
    }
  }

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Consolidator
        </h1>
        <p className="text-muted text-sm mt-1">
          Each <span className="font-medium">section</span> merges rows from
          one or more source spreadsheets (each with its own picked tabs) and
          writes a single output tab into the spreadsheet you choose. The
          output is the <span className="font-medium">union of every column</span>{" "}
          seen across all picked tabs — every source row is kept (no
          merging or deletion, so nothing is lost). After writing, duplicate
          Email and Phone values get highlighted (first occurrence light
          green, the rest light red), and if a Renewal Date column is
          present the four-tier renewal proximity highlighting is installed
          on the output tab. Each section is independent — use its own{" "}
          <span className="font-medium">Run section</span> button, or hit{" "}
          <span className="font-medium">Run all</span> at the bottom to do
          them in one go.
        </p>
      </div>

      {noMasterSheet && (
        <div className="bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          Set a master sheet on the Sheets page to save your configuration
          between sessions.
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
            result={results.get(section.id) ?? null}
            error={errors.get(section.id) ?? null}
            running={runningSections.has(section.id)}
            onChangeSection={(patch) => updateSection(sIdx, patch)}
            onRemove={() => removeSection(sIdx)}
            onAddSource={() => addSource(sIdx)}
            onRemoveSource={(srcIdx) => removeSource(sIdx, srcIdx)}
            onChangeSource={(srcIdx, patch) =>
              updateSource(sIdx, srcIdx, patch)
            }
            onFetchTabs={(srcIdx) => handleFetchTabs(sIdx, srcIdx)}
            onToggleTab={(srcIdx, tab) => toggleTab(sIdx, srcIdx, tab)}
            onSelectAll={(srcIdx, all) => selectAllTabs(sIdx, srcIdx, all)}
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

// ---------- Section card ----------

function SectionCard({
  section,
  index,
  total,
  result,
  error,
  running,
  onChangeSection,
  onRemove,
  onAddSource,
  onRemoveSource,
  onChangeSource,
  onFetchTabs,
  onToggleTab,
  onSelectAll,
  onRunSection,
}: {
  section: SectionState;
  index: number;
  total: number;
  result: SectionResult | null;
  error: string | null;
  running: boolean;
  onChangeSection: (patch: Partial<SectionState>) => void;
  onRemove: () => void;
  onAddSource: () => void;
  onRemoveSource: (sourceIdx: number) => void;
  onChangeSource: (sourceIdx: number, patch: Partial<SourceState>) => void;
  onFetchTabs: (sourceIdx: number) => void;
  onToggleTab: (sourceIdx: number, tab: string) => void;
  onSelectAll: (sourceIdx: number, all: boolean) => void;
  onRunSection: () => void;
}) {
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
            onChange={(e) => onChangeSection({ name: e.target.value })}
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

      {/* Source spreadsheets */}
      <div className="space-y-3">
        <p className="text-xs font-medium text-muted uppercase tracking-wide">
          Source spreadsheets ({section.sources.length})
        </p>
        {section.sources.map((src, srcIdx) => (
          <SourceRow
            key={srcIdx}
            source={src}
            sourceIdx={srcIdx}
            total={section.sources.length}
            onChange={(patch) => onChangeSource(srcIdx, patch)}
            onRemove={() => onRemoveSource(srcIdx)}
            onFetch={() => onFetchTabs(srcIdx)}
            onToggle={(tab) => onToggleTab(srcIdx, tab)}
            onSelectAll={(all) => onSelectAll(srcIdx, all)}
          />
        ))}
        <button
          type="button"
          onClick={onAddSource}
          className="text-xs font-medium px-3 py-1.5 rounded-md border border-dashed border-border bg-background hover:bg-card transition-colors cursor-pointer"
        >
          + Add another source spreadsheet
        </button>
      </div>

      {/* Output */}
      <div className="space-y-3 border-t border-border pt-4">
        <p className="text-xs font-medium text-muted uppercase tracking-wide">
          Output
        </p>
        <div className="flex flex-col sm:flex-row gap-2">
          <input
            type="url"
            value={section.outputUrl}
            onChange={(e) => onChangeSection({ outputUrl: e.target.value })}
            placeholder="Output spreadsheet URL (https://docs.google.com/spreadsheets/d/...)"
            className="flex-1 min-w-0 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
          <input
            type="text"
            value={section.outputTabName}
            onChange={(e) =>
              onChangeSection({ outputTabName: e.target.value })
            }
            placeholder="Tab name"
            className="sm:w-48 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </div>
        <p className="text-xs text-muted">
          Re-running this section overwrites the{" "}
          <code className="bg-background px-1 rounded">
            {section.outputTabName || "Consolidated"}
          </code>{" "}
          tab in the chosen output spreadsheet.
        </p>
      </div>

      {/* Per-section Run + error */}
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

      {/* Inline result for this section */}
      {result && <InlineResult result={result} />}
    </div>
  );
}

// ---------- Inline result block (rendered inside SectionCard) ----------

function InlineResult({ result }: { result: SectionResult }) {
  const failed = !!result.error;
  return (
    <div
      className={`border-t pt-4 ${
        failed ? "border-danger/30" : "border-border"
      }`}
    >
      <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2 mb-3">
        <h3 className="font-medium text-sm">Last run</h3>
        {result.outputSpreadsheetUrl && (
          <a
            href={result.outputSpreadsheetUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="text-sm text-primary hover:underline font-medium"
          >
            Open output →
          </a>
        )}
      </div>

      {failed ? (
        <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
          {result.error}
        </p>
      ) : (
        <>
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 text-sm">
            <Stat label="Source rows" value={result.totalSourceRows} />
            <Stat
              label="Output rows"
              value={result.totalOutputRows}
              accent="success"
            />
            <Stat
              label="Columns"
              value={result.totalColumns}
              accent="muted"
            />
            <Stat
              label="Email dupes"
              value={result.emailDuplicateCells}
              accent={result.emailDuplicateCells > 0 ? "danger" : "muted"}
              hint={`${result.emailDuplicateValues} distinct value${result.emailDuplicateValues === 1 ? "" : "s"}`}
            />
            <Stat
              label="Phone dupes"
              value={result.phoneDuplicateCells}
              accent={result.phoneDuplicateCells > 0 ? "danger" : "muted"}
              hint={`${result.phoneDuplicateValues} distinct value${result.phoneDuplicateValues === 1 ? "" : "s"}`}
            />
            <Stat
              label="Rows w/o email"
              value={result.rowsWithoutEmail}
              accent={result.rowsWithoutEmail > 0 ? "danger" : "muted"}
            />
          </div>
          {result.renewalRulesInstalled && (
            <p className="text-xs text-success bg-success/10 rounded-md px-3 py-2">
              Renewal Date column detected — 4-tier conditional formatting
              installed (past / 0-4d / 5-14d / 15-30d). Sheets re-evaluates
              it daily.
            </p>
          )}

          <div className="space-y-3 mt-3">
            {result.sources.map((src, i) => (
              <div key={i} className="border-t border-border pt-3">
                <div className="flex items-center justify-between gap-2 mb-1">
                  <a
                    href={src.spreadsheetUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-xs text-primary hover:underline truncate"
                  >
                    {src.spreadsheetUrl}
                  </a>
                </div>
                {src.error ? (
                  <p className="text-xs text-danger">{src.error}</p>
                ) : (
                  <div className="space-y-1">
                    {src.tabs.map((t, j) => (
                      <div
                        key={j}
                        className="flex flex-col gap-0.5 text-xs py-1"
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="font-medium">{t.tabName}</span>
                          {t.error ? (
                            <span className="text-danger">{t.error}</span>
                          ) : (
                            <span className="text-muted">
                              {t.rowsWithEmail} w/ email ·{" "}
                              {t.rowsWithoutEmail} w/o · {t.totalRows} total
                              {t.columnsContributed > 0 && (
                                <>
                                  {" "}·{" "}
                                  <span className="text-primary">
                                    +{t.columnsContributed} col
                                    {t.columnsContributed === 1 ? "" : "s"}
                                  </span>
                                </>
                              )}
                            </span>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// ---------- Source row ----------

function SourceRow({
  source,
  sourceIdx,
  total,
  onChange,
  onRemove,
  onFetch,
  onToggle,
  onSelectAll,
}: {
  source: SourceState;
  sourceIdx: number;
  total: number;
  onChange: (patch: Partial<SourceState>) => void;
  onRemove: () => void;
  onFetch: () => void;
  onToggle: (tab: string) => void;
  onSelectAll: (all: boolean) => void;
}) {
  return (
    <div className="border border-border rounded-md p-3 space-y-3 bg-background/50">
      <div className="flex items-center justify-between gap-2">
        <span className="text-xs text-muted">Source #{sourceIdx + 1}</span>
        {total > 1 && (
          <button
            type="button"
            onClick={onRemove}
            className="text-xs text-danger hover:underline cursor-pointer"
          >
            Remove source
          </button>
        )}
      </div>

      <div className="flex flex-col sm:flex-row gap-2">
        <input
          type="url"
          value={source.url}
          onChange={(e) =>
            onChange({
              url: e.target.value,
              tabs: [],
              selected: [],
            })
          }
          placeholder="https://docs.google.com/spreadsheets/d/..."
          className="flex-1 min-w-0 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
        <button
          type="button"
          onClick={onFetch}
          disabled={!source.url || source.loading}
          className="bg-foreground text-background px-3 py-2 rounded-md text-sm font-medium hover:opacity-80 disabled:opacity-50 transition-opacity cursor-pointer whitespace-nowrap"
        >
          {source.loading ? "Loading..." : "Fetch tabs"}
        </button>
      </div>

      {source.tabs.length > 0 && (
        <div>
          <div className="flex items-center justify-between mb-2">
            <p className="text-xs text-muted">
              {source.selected.length} of {source.tabs.length} selected
            </p>
            <div className="flex gap-3 text-xs">
              <button
                type="button"
                onClick={() => onSelectAll(true)}
                className="text-primary hover:underline cursor-pointer"
              >
                Select all
              </button>
              <button
                type="button"
                onClick={() => onSelectAll(false)}
                className="text-muted hover:underline cursor-pointer"
              >
                Select none
              </button>
            </div>
          </div>
          <div className="flex flex-wrap gap-2">
            {source.tabs.map((tab) => (
              <button
                key={tab}
                type="button"
                onClick={() => onToggle(tab)}
                className={`px-3 py-1.5 rounded-md text-sm border transition-colors cursor-pointer ${
                  source.selected.includes(tab)
                    ? "bg-primary text-white border-primary"
                    : "bg-background border-border text-foreground hover:border-primary/50"
                }`}
              >
                {tab}
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- Reusable stat tile ----------

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
