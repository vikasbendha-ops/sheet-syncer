"use client";

import { useState } from "react";

export default function CronSetup() {
  const [token, setToken] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [copied, setCopied] = useState(false);

  async function handleReveal() {
    setLoading(true);
    try {
      const res = await fetch("/api/auth/refresh-token");
      const data = await res.json();
      if (data.refreshToken) {
        setToken(data.refreshToken);
      }
    } finally {
      setLoading(false);
    }
  }

  async function handleCopy() {
    if (!token) return;
    await navigator.clipboard.writeText(token);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  return (
    <div className="bg-card border border-border rounded-lg p-5 space-y-3">
      <h2 className="font-medium">Cron Job Setup</h2>
      <p className="text-sm text-muted">
        For the auto-sync cron to work on Vercel, set your refresh token as the{" "}
        <code className="bg-background px-1.5 py-0.5 rounded text-xs font-mono">
          GOOGLE_REFRESH_TOKEN
        </code>{" "}
        environment variable.
      </p>

      {!token ? (
        <button
          onClick={handleReveal}
          disabled={loading}
          className="text-sm text-primary hover:text-primary-hover transition-colors cursor-pointer"
        >
          {loading ? "Loading..." : "Reveal refresh token"}
        </button>
      ) : (
        <div className="space-y-2">
          <code className="block bg-background border border-border rounded-md px-3 py-2 text-xs font-mono break-all select-all">
            {token}
          </code>
          <button
            onClick={handleCopy}
            className="text-sm text-primary hover:text-primary-hover transition-colors cursor-pointer"
          >
            {copied ? "Copied!" : "Copy to clipboard"}
          </button>
        </div>
      )}
    </div>
  );
}
