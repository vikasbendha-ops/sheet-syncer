"use client";

import { useState } from "react";

interface Result {
  spreadsheetUrl: string;
  tabs: Array<{
    tabName: string;
    totalRows: number;
    matched: number;
    unmatched: number;
    error?: string;
  }>;
}

export default function EmailFinderPage() {
  const [url, setUrl] = useState("");
  const [tabs, setTabs] = useState<string[]>([]);
  const [selectedTabs, setSelectedTabs] = useState<string[]>([]);
  const [loadingTabs, setLoadingTabs] = useState(false);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<Result | null>(null);

  async function handleFetchTabs() {
    if (!url) return;
    setError("");
    setLoadingTabs(true);
    setTabs([]);
    setSelectedTabs([]);
    setResult(null);

    try {
      const res = await fetch(`/api/sheets/tabs?url=${encodeURIComponent(url)}`);
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to fetch tabs");
        return;
      }
      setTabs(data.tabs ?? []);
      setSelectedTabs(data.tabs ?? []);
    } catch {
      setError("Network error");
    } finally {
      setLoadingTabs(false);
    }
  }

  function toggleTab(tab: string) {
    setSelectedTabs((prev) =>
      prev.includes(tab) ? prev.filter((t) => t !== tab) : [...prev, tab]
    );
  }

  async function handleRun(e: React.FormEvent) {
    e.preventDefault();
    if (selectedTabs.length === 0) {
      setError("Select at least one tab");
      return;
    }
    setError("");
    setRunning(true);
    setResult(null);

    try {
      const res = await fetch("/api/email-finder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url, tabs: selectedTabs }),
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

  const totalMatched = result?.tabs.reduce((a, t) => a + t.matched, 0) ?? 0;
  const totalUnmatched = result?.tabs.reduce((a, t) => a + t.unmatched, 0) ?? 0;

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Email Finder
        </h1>
        <p className="text-muted text-sm mt-1">
          Paste a Google Sheets URL, pick tabs. For each row, the tool looks up
          the person&apos;s name in your master sheet and writes their email
          into a new <code className="bg-card px-1 rounded">Email</code> column
          right after the name. Names with no match are highlighted red.
        </p>
      </div>

      <form
        onSubmit={handleRun}
        className="bg-card border border-border rounded-lg p-5 space-y-4"
      >
        <div>
          <label htmlFor="url" className="block text-sm font-medium mb-1">
            Google Sheets URL
          </label>
          <div className="flex flex-col sm:flex-row gap-2">
            <input
              id="url"
              type="url"
              required
              value={url}
              onChange={(e) => {
                setUrl(e.target.value);
                setTabs([]);
                setSelectedTabs([]);
                setResult(null);
              }}
              placeholder="https://docs.google.com/spreadsheets/d/..."
              className="flex-1 min-w-0 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            />
            <button
              type="button"
              onClick={handleFetchTabs}
              disabled={!url || loadingTabs}
              className="bg-foreground text-background px-4 py-2 rounded-md text-sm font-medium hover:opacity-80 disabled:opacity-50 transition-opacity cursor-pointer whitespace-nowrap"
            >
              {loadingTabs ? "Loading..." : "Fetch Spreadsheets"}
            </button>
          </div>
        </div>

        {tabs.length > 0 && (
          <div>
            <label className="block text-sm font-medium mb-2">
              Select spreadsheets to process
            </label>
            <div className="flex flex-wrap gap-2">
              {tabs.map((tab) => (
                <button
                  key={tab}
                  type="button"
                  onClick={() => toggleTab(tab)}
                  className={`px-3 py-1.5 rounded-md text-sm border transition-colors cursor-pointer ${
                    selectedTabs.includes(tab)
                      ? "bg-primary text-white border-primary"
                      : "bg-background border-border text-foreground hover:border-primary/50"
                  }`}
                >
                  {tab}
                </button>
              ))}
            </div>
            <p className="text-xs text-muted mt-1">
              {selectedTabs.length} of {tabs.length} spreadsheet(s) selected
            </p>
          </div>
        )}

        {error && (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
            {error}
          </p>
        )}

        {tabs.length > 0 && (
          <button
            type="submit"
            disabled={running || selectedTabs.length === 0}
            className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
          >
            {running ? "Finding emails..." : "Find emails & write"}
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
            <span className="font-medium text-success">{totalMatched}</span>{" "}
            matched ·{" "}
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
                    {t.matched} matched · {t.unmatched} unmatched · {t.totalRows} total
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
