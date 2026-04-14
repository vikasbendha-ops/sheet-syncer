"use client";

import { useState } from "react";

interface AddSheetFormProps {
  onAdded: () => void;
}

export default function AddSheetForm({ onAdded }: AddSheetFormProps) {
  const [url, setUrl] = useState("");
  const [nickname, setNickname] = useState("");
  const [emailColumn, setEmailColumn] = useState("auto");
  const [tabs, setTabs] = useState<string[]>([]);
  const [selectedTabs, setSelectedTabs] = useState<string[]>([]);
  const [loadingTabs, setLoadingTabs] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function handleFetchTabs() {
    if (!url) return;
    setError("");
    setLoadingTabs(true);
    setTabs([]);
    setSelectedTabs([]);

    try {
      const res = await fetch(`/api/sheets/tabs?url=${encodeURIComponent(url)}`);
      const data = await res.json();

      if (!res.ok) {
        setError(data.error || "Failed to fetch tabs");
        return;
      }

      setTabs(data.tabs ?? []);
      // Auto-select all tabs
      setSelectedTabs(data.tabs ?? []);
    } catch {
      setError("Network error. Please try again.");
    } finally {
      setLoadingTabs(false);
    }
  }

  function toggleTab(tab: string) {
    setSelectedTabs((prev) =>
      prev.includes(tab) ? prev.filter((t) => t !== tab) : [...prev, tab]
    );
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError("");

    if (selectedTabs.length === 0) {
      setError("Select at least one tab");
      return;
    }

    setLoading(true);

    try {
      const res = await fetch("/api/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url, nickname, tabs: selectedTabs, emailColumn }),
      });

      const data = await res.json();

      if (!res.ok) {
        setError(data.error || "Failed to add sheet");
        return;
      }

      setUrl("");
      setNickname("");
      setEmailColumn("auto");
      setTabs([]);
      setSelectedTabs([]);
      onAdded();
    } catch {
      setError("Network error. Please try again.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <form onSubmit={handleSubmit} className="bg-card border border-border rounded-lg p-5 space-y-4">
      <h3 className="font-medium text-sm text-muted uppercase tracking-wider">
        Add a Google Sheet
      </h3>

      {/* Step 1: URL + Fetch Tabs */}
      <div>
        <label htmlFor="url" className="block text-sm font-medium mb-1">
          Google Sheets URL
        </label>
        <div className="flex gap-2">
          <input
            id="url"
            type="url"
            required
            value={url}
            onChange={(e) => {
              setUrl(e.target.value);
              setTabs([]);
              setSelectedTabs([]);
            }}
            placeholder="https://docs.google.com/spreadsheets/d/..."
            className="flex-1 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
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

      {/* Step 2: Select Tabs */}
      {tabs.length > 0 && (
        <div>
          <label className="block text-sm font-medium mb-2">
            Select tabs to include
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
            {selectedTabs.length} of {tabs.length} tab(s) selected
          </p>
        </div>
      )}

      {/* Step 3: Nickname + Email Column */}
      {tabs.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label htmlFor="nickname" className="block text-sm font-medium mb-1">
              Nickname
            </label>
            <input
              id="nickname"
              type="text"
              required
              value={nickname}
              onChange={(e) => setNickname(e.target.value)}
              placeholder="e.g. Signups, Newsletter"
              className="w-full border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            />
            {selectedTabs.length > 1 && (
              <p className="text-xs text-muted mt-1">
                Each tab will appear as &quot;{nickname || "Nickname"} - TabName&quot;
              </p>
            )}
          </div>

          <div>
            <label htmlFor="emailColumn" className="block text-sm font-medium mb-1">
              Email Column
            </label>
            <input
              id="emailColumn"
              type="text"
              value={emailColumn}
              onChange={(e) => setEmailColumn(e.target.value)}
              placeholder="auto, A, B, C..."
              className="w-full border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            />
            <p className="text-xs text-muted mt-1">
              &quot;auto&quot; detects the email column from headers
            </p>
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
          disabled={loading || selectedTabs.length === 0}
          className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
        >
          {loading
            ? "Verifying & Adding..."
            : `Add ${selectedTabs.length} Tab${selectedTabs.length !== 1 ? "s" : ""}`}
        </button>
      )}
    </form>
  );
}
