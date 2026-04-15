"use client";

import { useEffect, useState, useCallback } from "react";
import SheetTable from "@/components/sheet-table";
import AddSheetForm from "@/components/add-sheet-form";
import MasterSheetCard from "@/components/master-sheet-card";
import { LinkedSheet } from "@/types";

export default function SheetsPage() {
  const [sheets, setSheets] = useState<LinkedSheet[]>([]);
  const [loading, setLoading] = useState(true);
  const [masterSheetId, setMasterSheetId] = useState<string | null>(null);

  const fetchSheets = useCallback(async () => {
    try {
      const res = await fetch("/api/sheets");
      const data = await res.json();
      setSheets(data.sheets ?? []);
    } catch {
      // handle error silently
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (masterSheetId) {
      fetchSheets();
    } else {
      setLoading(false);
    }
  }, [fetchSheets, masterSheetId]);

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl font-semibold tracking-tight">Linked Sheets</h1>
        <p className="text-muted text-sm mt-1">
          Manage the Google Sheets that feed into your master sheet.
        </p>
      </div>

      <MasterSheetCard onChange={setMasterSheetId} />

      {!masterSheetId ? (
        <div className="bg-card border border-border rounded-lg p-8 text-center text-muted">
          Set a master sheet above to start linking Google Sheets.
        </div>
      ) : loading ? (
        <div className="text-center py-12 text-muted">Loading...</div>
      ) : (
        <>
          <SheetTable
            sheets={sheets}
            onRemoved={fetchSheets}
            onUpdated={fetchSheets}
          />
          <AddSheetForm onAdded={fetchSheets} />
        </>
      )}
    </div>
  );
}
