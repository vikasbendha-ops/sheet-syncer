"use client";

import { useState } from "react";

interface TabResult {
  tabName: string;
  totalRows: number;
  matchedByEmail: number;
  matchedByPhone: number;
  unmatched: number;
  error?: string;
}

interface Result {
  spreadsheetUrl: string;
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

export default function BizTutorSyncPage() {
  const [source, setSource] = useState<SheetState>(initialSheet);
  const [lookup, setLookup] = useState<SheetState>(initialSheet);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<Result | null>(null);

  async function fetchTabs(
    state: SheetState,
    setState: (s: SheetState) => void
  ) {
    if (!state.url) return;
    setError("");
    setResult(null);
    setState({ ...state, loading: true, tabs: [], selected: [] });
    try {
      const res = await fetch(
        `/api/sheets/tabs?url=${encodeURIComponent(state.url)}`
      );
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to fetch tabs");
        setState({ ...state, loading: false });
        return;
      }
      setState({
        ...state,
        tabs: data.tabs ?? [],
        selected: data.tabs ?? [],
        loading: false,
      });
    } catch {
      setError("Network error");
      setState({ ...state, loading: false });
    }
  }

  function toggle(
    state: SheetState,
    setState: (s: SheetState) => void,
    tab: string
  ) {
    setState({
      ...state,
      selected: state.selected.includes(tab)
        ? state.selected.filter((t) => t !== tab)
        : [...state.selected, tab],
    });
  }

  async function handleRun(e: React.FormEvent) {
    e.preventDefault();
    if (source.selected.length === 0 || lookup.selected.length === 0) {
      setError("Select at least one tab from each spreadsheet");
      return;
    }
    setError("");
    setRunning(true);
    setResult(null);
    try {
      const res = await fetch("/api/biz-tutor-sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sourceUrl: source.url,
          sourceTabs: source.selected,
          lookupUrl: lookup.url,
          lookupTabs: lookup.selected,
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

  const totalEmail =
    result?.tabs.reduce((a, t) => a + t.matchedByEmail, 0) ?? 0;
  const totalPhone =
    result?.tabs.reduce((a, t) => a + t.matchedByPhone, 0) ?? 0;
  const totalUnmatched =
    result?.tabs.reduce((a, t) => a + t.unmatched, 0) ?? 0;

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          BIZ TUTOR Sync
        </h1>
        <p className="text-muted text-sm mt-1">
          Pick a destination sheet and a lookup sheet. For each row in the
          destination, the tool finds a matching record in the lookup sheet by{" "}
          <span className="font-medium">Email</span> (preferred) or{" "}
          <span className="font-medium">Phone</span> and fills in the{" "}
          <span className="font-medium">BIZ TUTOR</span> column. Existing values
          are preserved when no match is found. Re-running is safe (idempotent).
        </p>
      </div>

      <form
        onSubmit={handleRun}
        className="bg-card border border-border rounded-lg p-5 space-y-6"
      >
        <SheetSection
          label="Destination spreadsheet (BIZ TUTOR will be filled here)"
          state={source}
          setState={setSource}
          onFetch={() => fetchTabs(source, setSource)}
          onToggle={(tab) => toggle(source, setSource, tab)}
        />

        <SheetSection
          label="Lookup spreadsheet (source of BIZ TUTOR values)"
          state={lookup}
          setState={setLookup}
          onFetch={() => fetchTabs(lookup, setLookup)}
          onToggle={(tab) => toggle(lookup, setLookup, tab)}
        />

        {error && (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
            {error}
          </p>
        )}

        {source.tabs.length > 0 && lookup.tabs.length > 0 && (
          <button
            type="submit"
            disabled={
              running ||
              source.selected.length === 0 ||
              lookup.selected.length === 0
            }
            className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
          >
            {running ? "Syncing..." : "Run BIZ TUTOR sync"}
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
              Open the spreadsheet →
            </a>
          </div>

          <p className="text-sm text-muted">
            <span className="font-medium text-success">{totalEmail}</span>{" "}
            matched by email ·{" "}
            <span className="font-medium text-success">{totalPhone}</span>{" "}
            matched by phone ·{" "}
            <span className="font-medium text-danger">{totalUnmatched}</span>{" "}
            unmatched
          </p>

          <div className="space-y-2">
            {result.tabs.map((t, i) => (
              <div
                key={i}
                className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-1 text-sm py-2 border-b border-border last:border-0"
              >
                <span className="font-medium">{t.tabName}</span>
                {t.error ? (
                  <span className="text-danger text-xs">{t.error}</span>
                ) : (
                  <span className="text-muted text-xs sm:text-sm">
                    {t.matchedByEmail} email · {t.matchedByPhone} phone ·{" "}
                    {t.unmatched} unmatched · {t.totalRows} total
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

function SheetSection({
  label,
  state,
  setState,
  onFetch,
  onToggle,
}: {
  label: string;
  state: SheetState;
  setState: (s: SheetState) => void;
  onFetch: () => void;
  onToggle: (tab: string) => void;
}) {
  return (
    <div className="space-y-3">
      <label className="block text-sm font-medium">{label}</label>
      <div className="flex flex-col sm:flex-row gap-2">
        <input
          type="url"
          required
          value={state.url}
          onChange={(e) =>
            setState({
              ...state,
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
