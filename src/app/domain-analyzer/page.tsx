"use client";

import { useState } from "react";

interface AnalyzeResponse {
  success: boolean;
  spreadsheetUrl: string;
  writtenTabs: string[];
  errors: string[];
}

export default function DomainAnalyzerPage() {
  const [url, setUrl] = useState("");
  const [emailColumn, setEmailColumn] = useState("auto");
  const [tabs, setTabs] = useState<string[]>([]);
  const [selectedTabs, setSelectedTabs] = useState<string[]>([]);
  const [loadingTabs, setLoadingTabs] = useState(false);
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<AnalyzeResponse | null>(null);

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

  async function handleAnalyze(e: React.FormEvent) {
    e.preventDefault();
    if (selectedTabs.length === 0) {
      setError("Select at least one tab");
      return;
    }
    setError("");
    setAnalyzing(true);
    setResult(null);

    try {
      const res = await fetch("/api/domain-analyzer", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url, tabs: selectedTabs, emailColumn }),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to analyze");
        return;
      }
      setResult(data);
    } catch {
      setError("Network error");
    } finally {
      setAnalyzing(false);
    }
  }

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Domain Analyzer
        </h1>
        <p className="text-muted text-sm mt-1">
          Paste a Google Sheets URL, pick spreadsheets, and the email-domain
          breakdown (table + chart) will be written to the bottom of each one.
        </p>
      </div>

      <form
        onSubmit={handleAnalyze}
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
          <>
            <div>
              <label className="block text-sm font-medium mb-2">
                Select spreadsheets to analyze
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

            <div>
              <label
                htmlFor="emailColumn"
                className="block text-sm font-medium mb-1"
              >
                Email Column
              </label>
              <input
                id="emailColumn"
                type="text"
                value={emailColumn}
                onChange={(e) => setEmailColumn(e.target.value)}
                placeholder="auto, A, B, C..."
                className="w-full sm:max-w-xs border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
              />
              <p className="text-xs text-muted mt-1">
                &quot;auto&quot; detects the email column from headers
              </p>
            </div>
          </>
        )}

        {error && (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
            {error}
          </p>
        )}

        {tabs.length > 0 && (
          <button
            type="submit"
            disabled={analyzing || selectedTabs.length === 0}
            className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
          >
            {analyzing
              ? "Writing analysis..."
              : `Analyze & write to sheet`}
          </button>
        )}
      </form>

      {result && (
        <div
          className={`rounded-lg p-5 space-y-3 ${
            result.success
              ? "bg-success/10 border border-success/30"
              : "bg-danger/10 border border-danger/30"
          }`}
        >
          <h2 className="font-medium text-base">
            {result.success ? "Analysis written" : "Completed with issues"}
          </h2>
          {result.writtenTabs.length > 0 && (
            <p className="text-sm">
              Wrote analysis to {result.writtenTabs.length} tab(s):{" "}
              <span className="font-medium">
                {result.writtenTabs.join(", ")}
              </span>
            </p>
          )}
          <a
            href={result.spreadsheetUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-block text-sm text-primary hover:underline font-medium"
          >
            Open the spreadsheet →
          </a>
          {result.errors.length > 0 && (
            <ul className="text-sm text-danger list-disc list-inside space-y-1">
              {result.errors.map((e, i) => (
                <li key={i}>{e}</li>
              ))}
            </ul>
          )}
        </div>
      )}
    </div>
  );
}
