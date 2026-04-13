"use client";

import { useState } from "react";

interface AddSheetFormProps {
  onAdded: () => void;
}

export default function AddSheetForm({ onAdded }: AddSheetFormProps) {
  const [url, setUrl] = useState("");
  const [nickname, setNickname] = useState("");
  const [emailColumn, setEmailColumn] = useState("auto");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    setLoading(true);

    try {
      const res = await fetch("/api/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url, nickname, emailColumn }),
      });

      const data = await res.json();

      if (!res.ok) {
        setError(data.error || "Failed to add sheet");
        return;
      }

      setUrl("");
      setNickname("");
      setEmailColumn("auto");
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

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <div className="sm:col-span-2">
          <label htmlFor="url" className="block text-sm font-medium mb-1">
            Google Sheets URL
          </label>
          <input
            id="url"
            type="url"
            required
            value={url}
            onChange={(e) => setUrl(e.target.value)}
            placeholder="https://docs.google.com/spreadsheets/d/..."
            className="w-full border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </div>

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

      {error && (
        <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
          {error}
        </p>
      )}

      <button
        type="submit"
        disabled={loading}
        className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
      >
        {loading ? "Verifying & Adding..." : "Add Sheet"}
      </button>
    </form>
  );
}
