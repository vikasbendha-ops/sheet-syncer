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
  uniqueRows: number;
  duplicatesMerged: number;
  rowsWithoutEmail: number;
  totalColumns: number;
  sources: SourceResult[];
  error?: string;
}

interface BatchResult {
  sections: SectionResult[];
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

export default function ConsolidatorPage() {
  const [sections, setSections] = useState<SectionState[]>([makeSection(0)]);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<BatchResult | null>(null);
  const [hydrated, setHydrated] = useState(false);
  const [hasSaved, setHasSaved] = useState(false);
  const [noMasterSheet, setNoMasterSheet] = useState(false);

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
            // Rehydrate state, then asynchronously refetch tabs for each source.
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
        // Always keep at least one section — clear it instead of removing.
        return [makeSection(0)];
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
    setResult(null);
    setError("");
    updateSource(sectionIdx, sourceIdx, { loading: true });
    try {
      const tabs = await fetchTabsForUrl(src.url);
      updateSource(sectionIdx, sourceIdx, {
        tabs,
        selected: tabs, // default-select all on fresh fetch
        loading: false,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to fetch tabs");
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
    setError("");
    try {
      await fetch("/api/consolidator/config", { method: "DELETE" });
    } catch {
      // best-effort
    }
    setSections([makeSection(0)]);
    setResult(null);
    setHasSaved(false);
  }

  async function handleRun(e: React.FormEvent) {
    e.preventDefault();

    // Front-end validation: every section needs an outputUrl + at least one source with selected tabs.
    for (let i = 0; i < sections.length; i++) {
      const s = sections[i];
      if (!s.outputUrl) {
        setError(`Section "${s.name}" needs an Output spreadsheet URL.`);
        return;
      }
      if (!s.outputTabName.trim()) {
        setError(`Section "${s.name}" needs an Output tab name.`);
        return;
      }
      const hasAnyTabs = s.sources.some(
        (src) => src.url && src.selected.length > 0
      );
      if (!hasAnyTabs) {
        setError(
          `Section "${s.name}" needs at least one source spreadsheet with tabs selected.`
        );
        return;
      }
    }

    setError("");
    setRunning(true);
    setResult(null);
    try {
      const payload = {
        sections: sections.map((s) => ({
          id: s.id,
          name: s.name,
          outputUrl: s.outputUrl,
          outputTabName: s.outputTabName,
          sources: s.sources
            .filter((src) => src.url && src.selected.length > 0)
            .map((src) => ({ url: src.url, tabs: src.selected })),
        })),
      };
      const res = await fetch("/api/consolidator", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to run");
        return;
      }
      setResult(data);
    } catch {
      setError("Network error");
    } finally {
      setRunning(false);
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
          seen across all picked tabs — first-seen order, headers matched
          case-insensitively. Rows are deduped by{" "}
          <code className="bg-card px-1 rounded">Email</code>; when the same
          email appears more than once, the first non-blank value per column
          wins (so a column blank in one source picks up its value from
          another). Add multiple sections to run several consolidations one
          after another.
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

      <form onSubmit={handleRun} className="space-y-5">
        {sections.map((section, sIdx) => (
          <SectionCard
            key={section.id}
            section={section}
            index={sIdx}
            total={sections.length}
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

          {error && (
            <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2 flex-1">
              {error}
            </p>
          )}

          <button
            type="submit"
            disabled={running}
            className="bg-primary text-white px-5 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
          >
            {running
              ? "Running..."
              : sections.length === 1
                ? "Run consolidator"
                : `Run all ${sections.length} sections`}
          </button>
        </div>
      </form>

      {result && (
        <div className="space-y-3">
          <h2 className="font-medium">Results</h2>
          {result.sections.map((r, i) => (
            <ResultCard key={r.sectionId + i} result={r} />
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Section card ----------

function SectionCard({
  section,
  index,
  total,
  onChangeSection,
  onRemove,
  onAddSource,
  onRemoveSource,
  onChangeSource,
  onFetchTabs,
  onToggleTab,
  onSelectAll,
}: {
  section: SectionState;
  index: number;
  total: number;
  onChangeSection: (patch: Partial<SectionState>) => void;
  onRemove: () => void;
  onAddSource: () => void;
  onRemoveSource: (sourceIdx: number) => void;
  onChangeSource: (sourceIdx: number, patch: Partial<SourceState>) => void;
  onFetchTabs: (sourceIdx: number) => void;
  onToggleTab: (sourceIdx: number, tab: string) => void;
  onSelectAll: (sourceIdx: number, all: boolean) => void;
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

// ---------- Result card ----------

function ResultCard({ result }: { result: SectionResult }) {
  const failed = !!result.error;
  return (
    <div
      className={`border rounded-lg p-5 space-y-4 ${
        failed
          ? "bg-danger/5 border-danger/30"
          : "bg-card border-border"
      }`}
    >
      <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2">
        <div>
          <h3 className="font-medium">{result.sectionName}</h3>
          <p className="text-xs text-muted">
            Output tab:{" "}
            <code className="bg-background px-1 rounded">
              {result.outputTabName}
            </code>
          </p>
        </div>
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
        <p className="text-sm text-danger">{result.error}</p>
      ) : (
        <>
          <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 text-sm">
            <Stat label="Source rows" value={result.totalSourceRows} />
            <Stat label="Unique" value={result.uniqueRows} accent="success" />
            <Stat
              label="Duplicates merged"
              value={result.duplicatesMerged}
              accent="muted"
            />
            <Stat
              label="Rows w/o email"
              value={result.rowsWithoutEmail}
              accent={result.rowsWithoutEmail > 0 ? "danger" : "muted"}
            />
            <Stat
              label="Columns"
              value={result.totalColumns}
              accent="muted"
            />
          </div>

          <div className="space-y-3">
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

function Stat({
  label,
  value,
  accent,
}: {
  label: string;
  value: number;
  accent?: "success" | "danger" | "muted";
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
    </div>
  );
}
