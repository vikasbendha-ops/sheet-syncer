"use client";

import { useState } from "react";

interface TabResult {
  tabName: string;
  totalRows: number;
  emailDuplicateCells: number;
  phoneDuplicateCells: number;
  missingColumns: string[];
  detectedEmailHeader: string | null;
  detectedPhoneHeader: string | null;
  error?: string;
}

interface Result {
  spreadsheetUrl: string;
  totalDuplicateEmails: number;
  totalDuplicatePhones: number;
  totalDuplicateCells: number;
  tabs: TabResult[];
}

const EXPECTED_COLUMNS = ["Email", "Telefono Cellulare"];

export default function DuplicateFinderPage() {
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
      const res = await fetch("/api/duplicate-finder", {
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

  const nothingFound =
    result &&
    result.totalDuplicateEmails === 0 &&
    result.totalDuplicatePhones === 0;

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Duplicate Finder
        </h1>
        <p className="text-muted text-sm mt-1">
          Paste a Google Sheets URL and pick the tabs to scan. The tool checks{" "}
          {EXPECTED_COLUMNS.map((c, i) => (
            <span key={c}>
              <code className="bg-card px-1 rounded">{c}</code>
              {i < EXPECTED_COLUMNS.length - 1 ? " and " : ""}
            </span>
          ))}{" "}
          across all picked tabs as one dataset. The first occurrence of a
          duplicated value is colored{" "}
          <span className="inline-block w-3 h-3 rounded-sm align-middle border border-border" style={{ backgroundColor: "rgb(199, 235, 199)" }} />
          {" "}<span className="font-medium">light green</span>; every later
          occurrence is colored{" "}
          <span className="inline-block w-3 h-3 rounded-sm align-middle border border-border" style={{ backgroundColor: "rgb(255, 209, 209)" }} />
          {" "}<span className="font-medium">light red</span>. Values that
          appear only once are left untouched. Phone numbers match regardless
          of formatting (spaces, dashes, country code).
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
              {loadingTabs ? "Loading..." : "Fetch Tabs"}
            </button>
          </div>
        </div>

        {tabs.length > 0 && (
          <div>
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs text-muted">
                {selectedTabs.length} of {tabs.length} selected
              </p>
              <div className="flex gap-3 text-xs">
                <button
                  type="button"
                  onClick={() => setSelectedTabs(tabs)}
                  className="text-primary hover:underline cursor-pointer"
                >
                  Select all
                </button>
                <button
                  type="button"
                  onClick={() => setSelectedTabs([])}
                  className="text-muted hover:underline cursor-pointer"
                >
                  Select none
                </button>
              </div>
            </div>
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
            {running ? "Scanning..." : "Find duplicates"}
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

          {nothingFound ? (
            <p className="text-sm text-success bg-success/10 rounded-md px-3 py-2">
              No duplicates found. Every email and phone in the picked tabs is
              unique.
            </p>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-sm">
              <Stat
                label="Duplicated emails"
                value={result.totalDuplicateEmails}
                accent={result.totalDuplicateEmails > 0 ? "danger" : "muted"}
                hint="distinct values seen >1 time"
              />
              <Stat
                label="Duplicated phones"
                value={result.totalDuplicatePhones}
                accent={result.totalDuplicatePhones > 0 ? "danger" : "muted"}
                hint="distinct values seen >1 time"
              />
              <Stat
                label="Cells painted red"
                value={result.totalDuplicateCells}
                accent={result.totalDuplicateCells > 0 ? "danger" : "muted"}
                hint="total duplicate cells (excluding first occurrence)"
              />
            </div>
          )}

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
                      {t.emailDuplicateCells} email dupes ·{" "}
                      {t.phoneDuplicateCells} phone dupes · {t.totalRows} rows
                    </span>
                  )}
                </div>
                {!t.error && (
                  <p className="text-[10px] text-muted">
                    Matched email column:{" "}
                    {t.detectedEmailHeader ? (
                      <code className="bg-background px-1 rounded">
                        {t.detectedEmailHeader}
                      </code>
                    ) : (
                      <span className="text-danger">not found</span>
                    )}{" "}
                    · Matched phone column:{" "}
                    {t.detectedPhoneHeader ? (
                      <code className="bg-background px-1 rounded">
                        {t.detectedPhoneHeader}
                      </code>
                    ) : (
                      <span className="text-danger">not found</span>
                    )}
                  </p>
                )}
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
