"use client";

import { useState, useEffect } from "react";

interface TabResult {
  tabName: string;
  totalRows: number;
  rowsWithEmail: number;
  rowsWithoutEmail: number;
  missingColumns: string[];
  error?: string;
}

interface Result {
  spreadsheetUrl: string;
  totalSourceRows: number;
  uniqueRows: number;
  duplicatesMerged: number;
  rowsWithoutEmail: number;
  tabs: TabResult[];
}

interface SheetState {
  url: string;
  tabs: string[];
  selected: string[];
  loading: boolean;
}

const initialSheet: SheetState = {
  url: "",
  tabs: [],
  selected: [],
  loading: false,
};

const EXPECTED_COLUMNS = ["Nome", "Cognome", "Email", "Telefono Cellulare"];

export default function ConsolidatorPage() {
  const [source, setSource] = useState<SheetState>(initialSheet);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<Result | null>(null);
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
          if (data.code === "no_master_sheet") {
            setNoMasterSheet(true);
          }
        } else if (res.ok) {
          const data = (await res.json()) as {
            sourceUrl: string;
            sourceTabs: string[];
          };
          if (data.sourceUrl) {
            setHasSaved(true);
            await fetchTabsForUrl(data.sourceUrl, data.sourceTabs);
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
    if (!source.url && source.selected.length === 0) return;

    const handle = setTimeout(() => {
      fetch("/api/consolidator/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sourceUrl: source.url,
          sourceTabs: source.selected,
        }),
      })
        .then((r) => {
          if (r.ok) setHasSaved(true);
        })
        .catch(() => {
          // ignore
        });
    }, 800);
    return () => clearTimeout(handle);
  }, [source.url, source.selected, hydrated, noMasterSheet]);

  async function fetchTabsForUrl(
    url: string,
    desiredSelection: string[] | null
  ) {
    setError("");
    setSource({ url, tabs: [], selected: [], loading: true });
    try {
      const res = await fetch(
        `/api/sheets/tabs?url=${encodeURIComponent(url)}`
      );
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to fetch tabs");
        setSource({ url, tabs: [], selected: [], loading: false });
        return;
      }
      const tabs: string[] = data.tabs ?? [];
      const selected =
        desiredSelection && desiredSelection.length > 0
          ? desiredSelection.filter((t) => tabs.includes(t))
          : tabs;
      setSource({ url, tabs, selected, loading: false });
    } catch {
      setError("Network error");
      setSource({ url, tabs: [], selected: [], loading: false });
    }
  }

  function handleFetchTabs() {
    if (!source.url) return;
    setResult(null);
    fetchTabsForUrl(source.url, null);
  }

  function toggle(tab: string) {
    setSource({
      ...source,
      selected: source.selected.includes(tab)
        ? source.selected.filter((t) => t !== tab)
        : [...source.selected, tab],
    });
  }

  function selectAll() {
    setSource({ ...source, selected: source.tabs });
  }

  function selectNone() {
    setSource({ ...source, selected: [] });
  }

  async function handleClearSaved() {
    setError("");
    try {
      await fetch("/api/consolidator/config", { method: "DELETE" });
    } catch {
      // best-effort
    }
    setSource(initialSheet);
    setResult(null);
    setHasSaved(false);
  }

  async function handleRun(e: React.FormEvent) {
    e.preventDefault();
    if (source.selected.length === 0) {
      setError("Select at least one tab");
      return;
    }
    setError("");
    setRunning(true);
    setResult(null);
    try {
      const res = await fetch("/api/consolidator", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sourceUrl: source.url,
          sourceTabs: source.selected,
        }),
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
          Pick a Google Sheet and the tabs to merge. The tool reads{" "}
          {EXPECTED_COLUMNS.map((c, i) => (
            <span key={c}>
              <code className="bg-card px-1 rounded">{c}</code>
              {i < EXPECTED_COLUMNS.length - 1 ? ", " : ""}
            </span>
          ))}{" "}
          from each tab, dedupes by email (rows with a phone number win), and
          writes a single <code className="bg-card px-1 rounded">Consolidated</code>{" "}
          tab into your master sheet. Re-running overwrites it.
        </p>
      </div>

      {noMasterSheet && (
        <div className="bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          Set a master sheet on the Sheets page to use this feature.
        </div>
      )}

      {hasSaved && !noMasterSheet && (
        <div className="flex items-center justify-between gap-3 bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          <span>
            Selection saved to your master sheet — it&apos;ll auto-load next
            time.
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

      <form
        onSubmit={handleRun}
        className="bg-card border border-border rounded-lg p-5 space-y-6"
      >
        <div className="space-y-3">
          <label className="block text-sm font-medium">
            Source spreadsheet
          </label>
          <div className="flex flex-col sm:flex-row gap-2">
            <input
              type="url"
              required
              value={source.url}
              onChange={(e) =>
                setSource({
                  ...source,
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
              onClick={handleFetchTabs}
              disabled={!source.url || source.loading}
              className="bg-foreground text-background px-4 py-2 rounded-md text-sm font-medium hover:opacity-80 disabled:opacity-50 transition-opacity cursor-pointer whitespace-nowrap"
            >
              {source.loading ? "Loading..." : "Fetch Tabs"}
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
                    onClick={selectAll}
                    className="text-primary hover:underline cursor-pointer"
                  >
                    Select all
                  </button>
                  <button
                    type="button"
                    onClick={selectNone}
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
                    onClick={() => toggle(tab)}
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

        {error && (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
            {error}
          </p>
        )}

        {source.tabs.length > 0 && (
          <button
            type="submit"
            disabled={running || source.selected.length === 0}
            className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
          >
            {running ? "Consolidating..." : "Run consolidator"}
          </button>
        )}
      </form>

      {result && (
        <div className="bg-card border border-border rounded-lg p-5 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2">
            <h2 className="font-medium">Results</h2>
            <a
              href={result.spreadsheetUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="text-sm text-primary hover:underline font-medium"
            >
              Open the Consolidated tab →
            </a>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-sm">
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
          </div>

          <div className="space-y-2">
            {result.tabs.map((t, i) => (
              <div
                key={i}
                className="flex flex-col gap-1 text-sm py-2 border-b border-border last:border-0"
              >
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-1">
                  <span className="font-medium">{t.tabName}</span>
                  {t.error ? (
                    <span className="text-danger text-xs">{t.error}</span>
                  ) : (
                    <span className="text-muted text-xs sm:text-sm">
                      {t.rowsWithEmail} w/ email · {t.rowsWithoutEmail} w/o
                      email · {t.totalRows} total
                    </span>
                  )}
                </div>
                {t.missingColumns.length > 0 && !t.error && (
                  <span className="text-xs text-danger">
                    Missing columns: {t.missingColumns.join(", ")}
                  </span>
                )}
              </div>
            ))}
          </div>
        </div>
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
