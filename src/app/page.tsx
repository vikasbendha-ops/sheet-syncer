"use client";

import { useEffect, useState, useCallback } from "react";
import Link from "next/link";
import SyncButton from "@/components/sync-button";
import { LinkedSheet } from "@/types";

export default function Dashboard() {
  const [sheets, setSheets] = useState<LinkedSheet[]>([]);
  const [masterSheetId, setMasterSheetId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  const fetchSheets = useCallback(async () => {
    try {
      const res = await fetch("/api/sheets");
      const data = await res.json();
      setSheets(data.sheets ?? []);
    } catch {
      // silently fail on dashboard
    }
  }, []);

  const fetchMasterSheet = useCallback(async () => {
    try {
      const res = await fetch("/api/config/master-sheet");
      const data = await res.json();
      setMasterSheetId(data.masterSheetId ?? null);
    } catch {
      // ignore
    }
  }, []);

  useEffect(() => {
    (async () => {
      await fetchMasterSheet();
      await fetchSheets();
      setLoading(false);
    })();
  }, [fetchMasterSheet, fetchSheets]);

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
          <p className="text-base sm:text-lg font-medium mt-1 break-words">
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
            <Link
              href="/sheets"
              className="text-primary hover:underline text-sm mt-2 inline-block"
            >
              Not configured — set up here
            </Link>
          )}
        </div>
      </div>

      {/* Sync */}
      <div className="bg-card border border-border rounded-lg p-5 space-y-4">
        <h2 className="font-medium">Manual Sync</h2>
        <p className="text-sm text-muted">
          Trigger a full sync of all linked sheets into the master sheet.
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
                className="flex flex-col sm:flex-row sm:items-center sm:justify-between text-sm py-2 border-b border-border last:border-0 gap-1"
              >
                <span className="font-medium">{sheet.nickname}</span>
                <span className="text-muted text-xs sm:text-sm">
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
