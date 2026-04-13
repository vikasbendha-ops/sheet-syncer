"use client";

import { useEffect, useState, useCallback } from "react";
import SheetTable from "@/components/sheet-table";
import AddSheetForm from "@/components/add-sheet-form";
import { LinkedSheet } from "@/types";

export default function SheetsPage() {
  const [sheets, setSheets] = useState<LinkedSheet[]>([]);
  const [loading, setLoading] = useState(true);

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
    fetchSheets();
  }, [fetchSheets]);

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl font-semibold tracking-tight">Linked Sheets</h1>
        <p className="text-muted text-sm mt-1">
          Manage the Google Sheets that feed into your master sheet. Each sheet
          must be shared with your service account email.
        </p>
      </div>

      {loading ? (
        <div className="text-center py-12 text-muted">Loading...</div>
      ) : (
        <>
          <SheetTable sheets={sheets} onRemoved={fetchSheets} />
          <AddSheetForm onAdded={fetchSheets} />
        </>
      )}
    </div>
  );
}
