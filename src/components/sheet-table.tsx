"use client";

import { LinkedSheet } from "@/types";
import { useState } from "react";

interface SheetTableProps {
  sheets: LinkedSheet[];
  onRemoved: () => void;
}

export default function SheetTable({ sheets, onRemoved }: SheetTableProps) {
  const [removing, setRemoving] = useState<number | null>(null);

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
          {sheets.map((sheet, index) => (
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
              <td className="px-4 py-3 text-right">
                <button
                  onClick={() => handleRemove(index)}
                  disabled={removing === index}
                  className="text-danger hover:text-danger-hover text-sm font-medium disabled:opacity-50 transition-colors cursor-pointer"
                >
                  {removing === index ? "Removing..." : "Remove"}
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
