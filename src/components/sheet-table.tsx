"use client";

import { LinkedSheet } from "@/types";
import { useState } from "react";

interface SheetTableProps {
  sheets: LinkedSheet[];
  onRemoved: () => void;
  onUpdated: () => void;
}

export default function SheetTable({ sheets, onRemoved, onUpdated }: SheetTableProps) {
  const [removing, setRemoving] = useState<number | null>(null);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);

  async function handleRemove(index: number) {
    if (!confirm(`Remove "${sheets[index].nickname}"?`)) return;

    setRemoving(index);
    try {
      const res = await fetch(`/api/sheets/${index}`, { method: "DELETE" });
      if (res.ok) {
        onRemoved();
      }
    } finally {
      setRemoving(null);
    }
  }

  if (sheets.length === 0) {
    return (
      <div className="bg-card border border-border rounded-lg p-8 text-center text-muted">
        No sheets linked yet. Add one below.
      </div>
    );
  }

  return (
    <div className="bg-card border border-border rounded-lg overflow-hidden">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-border bg-background">
            <th className="text-left px-4 py-3 font-medium text-muted">Nickname</th>
            <th className="text-left px-4 py-3 font-medium text-muted">Tab</th>
            <th className="text-left px-4 py-3 font-medium text-muted">Email Column</th>
            <th className="text-left px-4 py-3 font-medium text-muted">Last Synced</th>
            <th className="text-right px-4 py-3 font-medium text-muted">Actions</th>
          </tr>
        </thead>
        <tbody>
          {sheets.map((sheet, index) =>
            editingIndex === index ? (
              <EditRow
                key={index}
                sheet={sheet}
                index={index}
                onCancel={() => setEditingIndex(null)}
                onSaved={() => {
                  setEditingIndex(null);
                  onUpdated();
                }}
              />
            ) : (
              <tr key={index} className="border-b border-border last:border-b-0">
                <td className="px-4 py-3">
                  <div className="font-medium">{sheet.nickname}</div>
                  <div className="text-xs text-muted truncate max-w-xs" title={sheet.url}>
                    {sheet.url}
                  </div>
                </td>
                <td className="px-4 py-3 text-muted">{sheet.tabName}</td>
                <td className="px-4 py-3 text-muted">
                  {sheet.emailColumn === "auto" ? "Auto-detect" : `Column ${sheet.emailColumn}`}
                </td>
                <td className="px-4 py-3 text-muted">
                  {sheet.lastSynced
                    ? new Date(sheet.lastSynced).toLocaleString()
                    : "Never"}
                </td>
                <td className="px-4 py-3 text-right whitespace-nowrap">
                  <button
                    onClick={() => setEditingIndex(index)}
                    className="text-primary hover:text-primary-hover text-sm font-medium transition-colors cursor-pointer mr-4"
                  >
                    Edit
                  </button>
                  <button
                    onClick={() => handleRemove(index)}
                    disabled={removing === index}
                    className="text-danger hover:text-danger-hover text-sm font-medium disabled:opacity-50 transition-colors cursor-pointer"
                  >
                    {removing === index ? "Removing..." : "Remove"}
                  </button>
                </td>
              </tr>
            )
          )}
        </tbody>
      </table>
    </div>
  );
}

interface EditRowProps {
  sheet: LinkedSheet;
  index: number;
  onCancel: () => void;
  onSaved: () => void;
}

function EditRow({ sheet, index, onCancel, onSaved }: EditRowProps) {
  const [nickname, setNickname] = useState(sheet.nickname);
  const [tabName, setTabName] = useState(sheet.tabName);
  const [emailColumn, setEmailColumn] = useState(sheet.emailColumn);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  async function handleSave() {
    setError("");
    setSaving(true);
    try {
      const res = await fetch(`/api/sheets/${index}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ nickname, tabName, emailColumn }),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to save");
        return;
      }
      onSaved();
    } catch {
      setError("Network error");
    } finally {
      setSaving(false);
    }
  }

  return (
    <tr className="border-b border-border last:border-b-0 bg-background/50">
      <td className="px-4 py-3">
        <input
          value={nickname}
          onChange={(e) => setNickname(e.target.value)}
          className="w-full border border-border rounded-md px-2 py-1 text-sm bg-card focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
        <div className="text-xs text-muted truncate max-w-xs mt-1" title={sheet.url}>
          {sheet.url}
        </div>
      </td>
      <td className="px-4 py-3">
        <input
          value={tabName}
          onChange={(e) => setTabName(e.target.value)}
          className="w-full border border-border rounded-md px-2 py-1 text-sm bg-card focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
      </td>
      <td className="px-4 py-3">
        <input
          value={emailColumn}
          onChange={(e) => setEmailColumn(e.target.value)}
          placeholder="auto, A, B..."
          className="w-full border border-border rounded-md px-2 py-1 text-sm bg-card focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
      </td>
      <td className="px-4 py-3 text-muted">
        {sheet.lastSynced ? new Date(sheet.lastSynced).toLocaleString() : "Never"}
      </td>
      <td className="px-4 py-3 text-right whitespace-nowrap">
        <button
          onClick={handleSave}
          disabled={saving}
          className="text-primary hover:text-primary-hover text-sm font-medium disabled:opacity-50 transition-colors cursor-pointer mr-4"
        >
          {saving ? "Saving..." : "Save"}
        </button>
        <button
          onClick={onCancel}
          disabled={saving}
          className="text-muted hover:text-foreground text-sm font-medium transition-colors cursor-pointer"
        >
          Cancel
        </button>
        {error && (
          <div className="text-xs text-danger mt-1 text-left">{error}</div>
        )}
      </td>
    </tr>
  );
}
