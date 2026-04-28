"use client";

import { useEffect, useState, useCallback } from "react";

interface MasterSheetCardProps {
  onChange?: (masterSheetId: string | null) => void;
}

export default function MasterSheetCard({ onChange }: MasterSheetCardProps) {
  const [masterSheetId, setMasterSheetId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [editing, setEditing] = useState(false);
  const [url, setUrl] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const fetchStatus = useCallback(async () => {
    try {
      const res = await fetch("/api/config/master-sheet");
      const data = await res.json();
      setMasterSheetId(data.masterSheetId ?? null);
      onChange?.(data.masterSheetId ?? null);
    } catch {
      // ignore
    } finally {
      setLoading(false);
    }
  }, [onChange]);

  useEffect(() => {
    fetchStatus();
  }, [fetchStatus]);

  async function handleSave(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    setSaving(true);
    try {
      const res = await fetch("/api/config/master-sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url }),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to save");
        return;
      }
      setMasterSheetId(data.masterSheetId);
      setEditing(false);
      setUrl("");
      onChange?.(data.masterSheetId);
    } catch {
      setError("Network error");
    } finally {
      setSaving(false);
    }
  }

  if (loading) return null;

  return (
    <div className="bg-card border border-border rounded-lg p-5 space-y-3">
      <div className="flex items-start justify-between gap-4">
        <div>
          <h2 className="font-medium">Master Sheet</h2>
          <p className="text-sm text-muted mt-1">
            The Google Sheet where merged email data is written.
          </p>
        </div>
        {!editing && masterSheetId && (
          <button
            onClick={() => {
              setEditing(true);
              setUrl("");
              setError("");
            }}
            className="text-sm text-primary hover:text-primary-hover font-medium transition-colors cursor-pointer"
          >
            Change
          </button>
        )}
      </div>

      {!editing && masterSheetId && (
        <a
          href={`https://docs.google.com/spreadsheets/d/${masterSheetId}/edit`}
          target="_blank"
          rel="noopener noreferrer"
          className="text-xs sm:text-sm text-primary hover:underline break-all inline-block"
        >
          https://docs.google.com/spreadsheets/d/{masterSheetId}/edit
        </a>
      )}

      {(editing || !masterSheetId) && (
        <form onSubmit={handleSave} className="space-y-3">
          {!masterSheetId && (
            <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
              No master sheet set. Paste a Google Sheets URL below to configure one.
            </p>
          )}
          <input
            type="url"
            required
            value={url}
            onChange={(e) => setUrl(e.target.value)}
            placeholder="https://docs.google.com/spreadsheets/d/..."
            className="w-full border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
          {error && (
            <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2">
              {error}
            </p>
          )}
          <div className="flex gap-2">
            <button
              type="submit"
              disabled={saving || !url}
              className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
            >
              {saving ? "Saving..." : "Save"}
            </button>
            {editing && (
              <button
                type="button"
                onClick={() => {
                  setEditing(false);
                  setError("");
                }}
                disabled={saving}
                className="px-4 py-2 rounded-md text-sm font-medium text-muted hover:text-foreground transition-colors cursor-pointer"
              >
                Cancel
              </button>
            )}
          </div>
        </form>
      )}
    </div>
  );
}
