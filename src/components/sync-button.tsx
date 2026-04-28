"use client";

import { useState } from "react";
import { SyncResult } from "@/types";

interface SyncButtonProps {
  onSynced?: (result: SyncResult) => void;
}

export default function SyncButton({ onSynced }: SyncButtonProps) {
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<SyncResult | null>(null);

  async function handleSync() {
    setLoading(true);
    setResult(null);

    try {
      const res = await fetch("/api/sync", { method: "POST" });
      const data: SyncResult = await res.json();
      setResult(data);
      onSynced?.(data);
    } catch {
      setResult({
        success: false,
        sheetsProcessed: 0,
        totalEmails: 0,
        errors: ["Network error"],
        timestamp: new Date().toISOString(),
      });
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="space-y-3">
      <button
        onClick={handleSync}
        disabled={loading}
        className="bg-primary text-white px-5 py-2.5 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
      >
        {loading ? "Syncing..." : "Sync Now"}
      </button>

      {result && (
        <div
          className={`text-sm rounded-md px-4 py-3 ${
            result.success
              ? "bg-success/10 text-success"
              : "bg-danger/10 text-danger"
          }`}
        >
          {result.success ? (
            <p>
              Synced {result.sheetsProcessed} sheet(s) with {result.totalEmails}{" "}
              unique email(s).
            </p>
          ) : (
            <div>
              <p>Sync completed with errors:</p>
              <ul className="mt-1 list-disc list-inside">
                {(result.errors ?? []).map((err, i) => (
                  <li key={i}>{err}</li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
