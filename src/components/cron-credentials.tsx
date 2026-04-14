"use client";

import { useState } from "react";

export default function CronCredentials() {
  const [show, setShow] = useState(false);
  const [loading, setLoading] = useState(false);
  const [creds, setCreds] = useState<{
    refreshToken: string;
    masterSheetId: string | null;
  } | null>(null);
  const [copiedField, setCopiedField] = useState<string | null>(null);

  async function handleReveal() {
    setLoading(true);
    try {
      const res = await fetch("/api/auth/credentials");
      const data = await res.json();
      if (res.ok) {
        setCreds(data);
        setShow(true);
      }
    } finally {
      setLoading(false);
    }
  }

  async function handleCopy(value: string, field: string) {
    await navigator.clipboard.writeText(value);
    setCopiedField(field);
    setTimeout(() => setCopiedField(null), 2000);
  }

  return (
    <div className="bg-card border border-border rounded-lg p-5 space-y-3">
      <div>
        <h2 className="font-medium">Cron Job Setup</h2>
        <p className="text-sm text-muted mt-1">
          To enable the built-in 30-minute auto-sync, set these as environment
          variables on your host and restart the app. No external cron needed —
          the schedule runs inside the Node process.
        </p>
      </div>

      {!show ? (
        <button
          onClick={handleReveal}
          disabled={loading}
          className="text-sm text-primary hover:text-primary-hover transition-colors cursor-pointer"
        >
          {loading ? "Loading..." : "Reveal credentials"}
        </button>
      ) : creds ? (
        <div className="space-y-3">
          <CredentialField
            label="GOOGLE_REFRESH_TOKEN"
            value={creds.refreshToken}
            field="refresh"
            copiedField={copiedField}
            onCopy={handleCopy}
          />
          {creds.masterSheetId ? (
            <CredentialField
              label="MASTER_SHEET_ID"
              value={creds.masterSheetId}
              field="master"
              copiedField={copiedField}
              onCopy={handleCopy}
            />
          ) : (
            <p className="text-xs text-danger">
              MASTER_SHEET_ID is not set yet — configure the master sheet above first.
            </p>
          )}
          <p className="text-xs text-muted">
            Default schedule is every 30 minutes. To override, set{" "}
            <code className="bg-background px-1 py-0.5 rounded">CRON_SCHEDULE</code>{" "}
            to a standard cron expression (e.g.{" "}
            <code className="bg-background px-1 py-0.5 rounded">0 * * * *</code>{" "}
            for hourly).
          </p>
          <button
            onClick={() => setShow(false)}
            className="text-xs text-muted hover:text-foreground transition-colors cursor-pointer"
          >
            Hide
          </button>
        </div>
      ) : null}
    </div>
  );
}

function CredentialField({
  label,
  value,
  field,
  copiedField,
  onCopy,
}: {
  label: string;
  value: string;
  field: string;
  copiedField: string | null;
  onCopy: (value: string, field: string) => void;
}) {
  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <code className="text-xs font-mono text-muted">{label}</code>
        <button
          onClick={() => onCopy(value, field)}
          className="text-xs text-primary hover:text-primary-hover transition-colors cursor-pointer"
        >
          {copiedField === field ? "Copied!" : "Copy"}
        </button>
      </div>
      <code className="block bg-background border border-border rounded-md px-3 py-2 text-xs font-mono break-all select-all">
        {value}
      </code>
    </div>
  );
}
