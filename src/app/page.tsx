"use client";

import { useEffect, useState, useCallback } from "react";
import SyncButton from "@/components/sync-button";
import { LinkedSheet } from "@/types";

export default function Dashboard() {
  const [sheets, setSheets] = useState<LinkedSheet[]>([]);
  const [loading, setLoading] = useState(true);

  const masterSheetId = process.env.NEXT_PUBLIC_MASTER_SHEET_ID;

  const fetchSheets = useCallback(async () => {
    try {
      const res = await fetch("/api/sheets");
      const data = await res.json();
      setSheets(data.sheets ?? []);
    } catch {
      // silently fail on dashboard
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchSheets();
  }, [fetchSheets]);

  const lastSynced = sheets
    .map((s) => s.lastSynced)
    .filter(Boolean)
    .sort()
    .pop();

  const totalSheets = sheets.length;

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl font-semibold tracking-tight">Dashboard</h1>
        <p className="text-muted text-sm mt-1">
          Monitor your Google Sheets email sync
        </p>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-card border border-border rounded-lg p-5">
          <p className="text-sm text-muted">Linked Sheets</p>
          <p className="text-3xl font-semibold mt-1">
            {loading ? "-" : totalSheets}
          </p>
        </div>
        <div className="bg-card border border-border rounded-lg p-5">
          <p className="text-sm text-muted">Last Synced</p>
          <p className="text-lg font-medium mt-1">
            {loading
              ? "-"
              : lastSynced
                ? new Date(lastSynced).toLocaleString()
                : "Never"}
          </p>
        </div>
        <div className="bg-card border border-border rounded-lg p-5">
          <p className="text-sm text-muted">Master Sheet</p>
          {masterSheetId ? (
            <a
              href={`https://docs.google.com/spreadsheets/d/${masterSheetId}/edit`}
              target="_blank"
              rel="noopener noreferrer"
              className="text-primary hover:underline text-sm mt-2 inline-block"
            >
              Open in Google Sheets
            </a>
          ) : (
            <p className="text-sm text-muted mt-2">
              Set NEXT_PUBLIC_MASTER_SHEET_ID env var
            </p>
          )}
        </div>
      </div>

      {/* Sync */}
      <div className="bg-card border border-border rounded-lg p-5 space-y-4">
        <h2 className="font-medium">Manual Sync</h2>
        <p className="text-sm text-muted">
          Trigger a full sync of all linked sheets into the master sheet. The
          auto-sync cron runs every 15 minutes on Vercel.
        </p>
        <SyncButton onSynced={() => fetchSheets()} />
      </div>

      {/* Linked Sheets Quick View */}
      {!loading && sheets.length > 0 && (
        <div className="bg-card border border-border rounded-lg p-5">
          <h2 className="font-medium mb-3">Linked Sheets</h2>
          <div className="space-y-2">
            {sheets.map((sheet, i) => (
              <div
                key={i}
                className="flex items-center justify-between text-sm py-2 border-b border-border last:border-0"
              >
                <span className="font-medium">{sheet.nickname}</span>
                <span className="text-muted">
                  {sheet.lastSynced
                    ? `Synced ${new Date(sheet.lastSynced).toLocaleString()}`
                    : "Not synced"}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
