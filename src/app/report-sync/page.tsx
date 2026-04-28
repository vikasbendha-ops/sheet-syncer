"use client";

import { useState, useEffect } from "react";

interface TabResult {
  tabName: string;
  totalRows: number;
  matchedByEmail: number;
  matchedByPhone: number;
  unmatched: number;
  error?: string;
}

interface DestinationResult {
  spreadsheetUrl: string;
  tabs: TabResult[];
  error?: string;
}

interface RunResult {
  destinations: DestinationResult[];
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

export default function ReportSyncPage() {
  const [source, setSource] = useState<SheetState>(initialSheet);
  const [destinations, setDestinations] = useState<SheetState[]>([
    { ...initialSheet },
  ]);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<RunResult | null>(null);
  const [hydrated, setHydrated] = useState(false);
  const [hasSaved, setHasSaved] = useState(false);
  const [noMasterSheet, setNoMasterSheet] = useState(false);

  // Hydrate from server on mount
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/report-sync/config");
        if (res.status === 400) {
          const data = await res.json();
          if (data.code === "no_master_sheet") {
            setNoMasterSheet(true);
          }
        } else if (res.ok) {
          const data = (await res.json()) as {
            sourceUrl: string;
            sourceTabs: string[];
            destinations: { url: string; tabs: string[] }[];
          };
          const promises: Promise<void>[] = [];
          if (data.sourceUrl) {
            setHasSaved(true);
            promises.push(
              fetchTabsForUrl(data.sourceUrl, data.sourceTabs, (s) =>
                setSource(s)
              )
            );
          }
          if (data.destinations && data.destinations.length > 0) {
            setHasSaved(true);
            const destStates: SheetState[] = data.destinations.map(() => ({
              ...initialSheet,
            }));
            setDestinations(destStates);
            data.destinations.forEach((dest, i) => {
              if (dest.url) {
                promises.push(
                  fetchTabsForUrl(dest.url, dest.tabs, (s) =>
                    setDestinations((prev) => {
                      const next = [...prev];
                      next[i] = s;
                      return next;
                    })
                  )
                );
              }
            });
          }
          await Promise.all(promises);
        }
      } catch {
        // ignore — falls back to empty form
      }
      setHydrated(true);
    })();
  }, []);

  // Persist to server on changes (debounced)
  useEffect(() => {
    if (!hydrated || noMasterSheet) return;
    const hasAnything =
      source.url ||
      source.selected.length > 0 ||
      destinations.some((d) => d.url || d.selected.length > 0);
    if (!hasAnything) return;

    const handle = setTimeout(() => {
      fetch("/api/report-sync/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sourceUrl: source.url,
          sourceTabs: source.selected,
          destinations: destinations.map((d) => ({
            url: d.url,
            tabs: d.selected,
          })),
        }),
      })
        .then((r) => {
          if (r.ok) setHasSaved(true);
        })
        .catch(() => {
          // ignore — user can retry on next change
        });
    }, 800);
    return () => clearTimeout(handle);
  }, [source.url, source.selected, destinations, hydrated, noMasterSheet]);

  async function fetchTabsForUrl(
    url: string,
    desiredSelection: string[] | null,
    setState: (s: SheetState) => void
  ) {
    setError("");
    setState({ url, tabs: [], selected: [], loading: true });
    try {
      const res = await fetch(`/api/sheets/tabs?url=${encodeURIComponent(url)}`);
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to fetch tabs");
        setState({ url, tabs: [], selected: [], loading: false });
        return;
      }
      const tabs: string[] = data.tabs ?? [];
      const selected =
        desiredSelection && desiredSelection.length > 0
          ? desiredSelection.filter((t) => tabs.includes(t))
          : tabs;
      setState({ url, tabs, selected, loading: false });
    } catch {
      setError("Network error");
      setState({ url, tabs: [], selected: [], loading: false });
    }
  }

  function fetchSourceTabs() {
    if (!source.url) return;
    setResult(null);
    fetchTabsForUrl(source.url, null, (s) => setSource(s));
  }

  function fetchDestinationTabs(index: number) {
    const dest = destinations[index];
    if (!dest?.url) return;
    setResult(null);
    fetchTabsForUrl(dest.url, null, (s) =>
      setDestinations((prev) => {
        const next = [...prev];
        next[index] = s;
        return next;
      })
    );
  }

  function toggleSourceTab(tab: string) {
    setSource({
      ...source,
      selected: source.selected.includes(tab)
        ? source.selected.filter((t) => t !== tab)
        : [...source.selected, tab],
    });
  }

  function toggleDestinationTab(index: number, tab: string) {
    setDestinations((prev) => {
      const next = [...prev];
      const cur = next[index];
      next[index] = {
        ...cur,
        selected: cur.selected.includes(tab)
          ? cur.selected.filter((t) => t !== tab)
          : [...cur.selected, tab],
      };
      return next;
    });
  }

  function setDestinationUrl(index: number, url: string) {
    setDestinations((prev) => {
      const next = [...prev];
      next[index] = { url, tabs: [], selected: [], loading: false };
      return next;
    });
  }

  function addDestination() {
    setDestinations((prev) => [...prev, { ...initialSheet }]);
  }

  function removeDestination(index: number) {
    setDestinations((prev) => prev.filter((_, i) => i !== index));
  }

  async function handleClearSaved() {
    setError("");
    try {
      await fetch("/api/report-sync/config", { method: "DELETE" });
    } catch {
      // best-effort — still reset UI
    }
    setSource({ ...initialSheet });
    setDestinations([{ ...initialSheet }]);
    setResult(null);
    setHasSaved(false);
  }

  async function handleRun(e: React.FormEvent) {
    e.preventDefault();
    if (source.selected.length === 0) {
      setError("Select at least one tab from the source spreadsheet");
      return;
    }
    if (destinations.length === 0) {
      setError("Add at least one destination spreadsheet");
      return;
    }
    for (let i = 0; i < destinations.length; i++) {
      if (!destinations[i].url || destinations[i].selected.length === 0) {
        setError(`Select at least one tab for destination ${i + 1}`);
        return;
      }
    }
    setError("");
    setRunning(true);
    setResult(null);
    try {
      const res = await fetch("/api/report-sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sourceUrl: source.url,
          sourceTabs: source.selected,
          destinations: destinations.map((d) => ({
            url: d.url,
            tabs: d.selected,
          })),
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
          Report Sync
        </h1>
        <p className="text-muted text-sm mt-1">
          Pick one source spreadsheet and one or more destination spreadsheets.
          For every row in each destination, the tool finds a matching record
          in the source by{" "}
          <span className="font-medium">Email</span> (preferred) or{" "}
          <span className="font-medium">Phone</span> and copies{" "}
          <span className="font-medium">ESITO</span>,{" "}
          <span className="font-medium">MOTIVAZIONE</span>,{" "}
          <span className="font-medium">NOTE VENDITRICE</span>, and{" "}
          <span className="font-medium">DATA RECALL</span>. Existing values are
          preserved when no match is found. Re-running is safe (idempotent).
        </p>
      </div>

      {noMasterSheet && (
        <div className="bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          Set a master sheet on the Sheets page to save these links across
          sessions.
        </div>
      )}

      {hasSaved && !noMasterSheet && (
        <div className="flex items-center justify-between gap-3 bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          <span>
            Links saved to your master sheet — they&apos;ll auto-load next time.
          </span>
          <button
            type="button"
            onClick={handleClearSaved}
            className="text-danger hover:underline font-medium cursor-pointer whitespace-nowrap"
          >
            Clear saved links
          </button>
        </div>
      )}

      <form
        onSubmit={handleRun}
        className="bg-card border border-border rounded-lg p-5 space-y-6"
      >
        <SheetSection
          label="Source spreadsheet (provides ESITO, MOTIVAZIONE, NOTE VENDITRICE, DATA RECALL)"
          state={source}
          onUrlChange={(url) =>
            setSource({ url, tabs: [], selected: [], loading: false })
          }
          onFetch={fetchSourceTabs}
          onToggle={toggleSourceTab}
        />

        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-medium">
              Destination spreadsheets ({destinations.length})
            </h2>
            <button
              type="button"
              onClick={addDestination}
              className="text-primary hover:underline text-sm font-medium cursor-pointer"
            >
              + Add destination
            </button>
          </div>

          {destinations.map((dest, i) => (
            <div
              key={i}
              className="border border-border rounded-md p-4 space-y-3 relative"
            >
              <div className="flex items-center justify-between">
                <span className="text-xs text-muted font-medium">
                  Destination {i + 1}
                </span>
                {destinations.length > 1 && (
                  <button
                    type="button"
                    onClick={() => removeDestination(i)}
                    className="text-danger hover:underline text-xs font-medium cursor-pointer"
                  >
                    Remove
                  </button>
                )}
              </div>
              <SheetSection
                label=""
                state={dest}
                onUrlChange={(url) => setDestinationUrl(i, url)}
                onFetch={() => fetchDestinationTabs(i)}
                onToggle={(tab) => toggleDestinationTab(i, tab)}
              />
            </div>
          ))}
        </div>

        {error && (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
            {error}
          </p>
        )}

        <button
          type="submit"
          disabled={
            running ||
            source.selected.length === 0 ||
            destinations.length === 0 ||
            destinations.some((d) => d.selected.length === 0)
          }
          className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
        >
          {running ? "Syncing..." : "Run Report sync"}
        </button>
      </form>

      {result && (
        <div className="space-y-4">
          {result.destinations.map((d, i) => {
            const totalEmail = d.tabs.reduce(
              (a, t) => a + t.matchedByEmail,
              0
            );
            const totalPhone = d.tabs.reduce(
              (a, t) => a + t.matchedByPhone,
              0
            );
            const totalUnmatched = d.tabs.reduce(
              (a, t) => a + t.unmatched,
              0
            );
            return (
              <div
                key={i}
                className="bg-card border border-border rounded-lg p-5 space-y-4"
              >
                <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2">
                  <h2 className="font-medium">Destination {i + 1}</h2>
                  <a
                    href={d.spreadsheetUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-sm text-primary hover:underline font-medium"
                  >
                    Open the spreadsheet →
                  </a>
                </div>

                {d.error ? (
                  <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
                    {d.error}
                  </p>
                ) : (
                  <>
                    <p className="text-sm text-muted">
                      <span className="font-medium text-success">
                        {totalEmail}
                      </span>{" "}
                      matched by email ·{" "}
                      <span className="font-medium text-success">
                        {totalPhone}
                      </span>{" "}
                      matched by phone ·{" "}
                      <span className="font-medium text-danger">
                        {totalUnmatched}
                      </span>{" "}
                      unmatched
                    </p>

                    <div className="space-y-2">
                      {d.tabs.map((t, j) => (
                        <div
                          key={j}
                          className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-1 text-sm py-2 border-b border-border last:border-0"
                        >
                          <span className="font-medium">{t.tabName}</span>
                          {t.error ? (
                            <span className="text-danger text-xs">
                              {t.error}
                            </span>
                          ) : (
                            <span className="text-muted text-xs sm:text-sm">
                              {t.matchedByEmail} email · {t.matchedByPhone}{" "}
                              phone · {t.unmatched} unmatched · {t.totalRows}{" "}
                              total
                            </span>
                          )}
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function SheetSection({
  label,
  state,
  onUrlChange,
  onFetch,
  onToggle,
}: {
  label: string;
  state: SheetState;
  onUrlChange: (url: string) => void;
  onFetch: () => void;
  onToggle: (tab: string) => void;
}) {
  return (
    <div className="space-y-3">
      {label && (
        <label className="block text-sm font-medium">{label}</label>
      )}
      <div className="flex flex-col sm:flex-row gap-2">
        <input
          type="url"
          value={state.url}
          onChange={(e) => onUrlChange(e.target.value)}
          placeholder="https://docs.google.com/spreadsheets/d/..."
          className="flex-1 min-w-0 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
        <button
          type="button"
          onClick={onFetch}
          disabled={!state.url || state.loading}
          className="bg-foreground text-background px-4 py-2 rounded-md text-sm font-medium hover:opacity-80 disabled:opacity-50 transition-opacity cursor-pointer whitespace-nowrap"
        >
          {state.loading ? "Loading..." : "Fetch Tabs"}
        </button>
      </div>

      {state.tabs.length > 0 && (
        <div>
          <div className="flex flex-wrap gap-2">
            {state.tabs.map((tab) => (
              <button
                key={tab}
                type="button"
                onClick={() => onToggle(tab)}
                className={`px-3 py-1.5 rounded-md text-sm border transition-colors cursor-pointer ${
                  state.selected.includes(tab)
                    ? "bg-primary text-white border-primary"
                    : "bg-background border-border text-foreground hover:border-primary/50"
                }`}
              >
                {tab}
              </button>
            ))}
          </div>
          <p className="text-xs text-muted mt-1">
            {state.selected.length} of {state.tabs.length} selected
          </p>
        </div>
      )}
    </div>
  );
}
