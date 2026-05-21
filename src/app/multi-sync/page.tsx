// src/app/multi-sync/page.tsx
"use client";

import { useCallback, useEffect, useState } from "react";
import type {
  MultiLinkedSheet,
  MultiSyncSection,
  MultiSyncSectionResult,
} from "@/lib/multi-sync-types";

// ---------- Local UI state ----------

interface LinkedSheetUi {
  loadedTabs: string[];
  loading: boolean;
}

interface LinkedSheetState extends MultiLinkedSheet {
  ui: LinkedSheetUi;
}

interface SectionState extends Omit<MultiSyncSection, "linkedSheets"> {
  linkedSheets: LinkedSheetState[];
}

function newId(): string {
  return `sec_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function emptyLinkedSheet(): LinkedSheetState {
  return {
    url: "",
    nickname: "",
    tabName: "",
    emailColumn: "auto",
    lastSynced: "",
    ui: { loadedTabs: [], loading: false },
  };
}

function makeSection(slot: number): SectionState {
  return {
    id: newId(),
    name: `Section ${slot}`,
    slot,
    masterTabName: `Multi Master - ${slot}`,
    presentInColumnName: `Present In - ${slot}`,
    linkedSheets: [emptyLinkedSheet()],
  };
}

function toConfigSection(s: SectionState): MultiSyncSection {
  return {
    id: s.id,
    name: s.name,
    slot: s.slot,
    masterTabName: s.masterTabName,
    presentInColumnName: s.presentInColumnName,
    linkedSheets: s.linkedSheets.map(({ ui: _ui, ...rest }) => {
      void _ui;
      return rest;
    }),
  };
}

async function fetchTabsForUrl(url: string): Promise<string[]> {
  const res = await fetch(`/api/sheets/tabs?url=${encodeURIComponent(url)}`);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to fetch tabs");
  return data.tabs ?? [];
}

async function claimSlot(): Promise<number> {
  const res = await fetch("/api/multi-sync/claim-slot", { method: "POST" });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to claim slot");
  return data.slot as number;
}

function formatLastSynced(iso: string): string {
  if (!iso) return "Never";
  try {
    return new Date(iso).toLocaleString();
  } catch {
    return iso;
  }
}

export default function MultiSyncPage() {
  const [sections, setSections] = useState<SectionState[]>([]);
  const [results, setResults] = useState<Map<string, MultiSyncSectionResult>>(
    new Map()
  );
  const [errors, setErrors] = useState<Map<string, string>>(new Map());
  const [runningSections, setRunningSections] = useState<Set<string>>(
    new Set()
  );

  const [hydrated, setHydrated] = useState(false);
  const [hasSaved, setHasSaved] = useState(false);
  const [noMasterSheet, setNoMasterSheet] = useState(false);
  const [bootstrapping, setBootstrapping] = useState(false);

  function setSectionRunning(id: string, on: boolean) {
    setRunningSections((prev) => {
      const next = new Set(prev);
      if (on) next.add(id);
      else next.delete(id);
      return next;
    });
  }
  function setSectionError(id: string, msg: string | null) {
    setErrors((prev) => {
      const next = new Map(prev);
      if (msg) next.set(id, msg);
      else next.delete(id);
      return next;
    });
  }
  function setSectionResult(id: string, r: MultiSyncSectionResult | null) {
    setResults((prev) => {
      const next = new Map(prev);
      if (r) next.set(id, r);
      else next.delete(id);
      return next;
    });
  }

  // Hydrate from server on mount
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/multi-sync/config");
        if (res.status === 400) {
          const data = await res.json();
          if (data.code === "no_master_sheet") setNoMasterSheet(true);
        } else if (res.ok) {
          const data = (await res.json()) as { sections: MultiSyncSection[] };
          if (data.sections && data.sections.length > 0) {
            setHasSaved(true);
            const rehydrated: SectionState[] = data.sections.map(
              (s): SectionState => ({
                id: s.id || newId(),
                name: s.name,
                slot: s.slot,
                masterTabName: s.masterTabName,
                presentInColumnName: s.presentInColumnName,
                linkedSheets:
                  s.linkedSheets && s.linkedSheets.length > 0
                    ? s.linkedSheets.map((l) => ({
                        ...l,
                        ui: { loadedTabs: [], loading: false },
                      }))
                    : [emptyLinkedSheet()],
              })
            );
            setSections(rehydrated);

            // Eagerly fetch tabs for any linked sheet with a URL
            await Promise.all(
              rehydrated.map(async (sec, sIdx) => {
                await Promise.all(
                  sec.linkedSheets.map(async (linked, lIdx) => {
                    if (!linked.url) return;
                    try {
                      const tabs = await fetchTabsForUrl(linked.url);
                      setSections((prev) => {
                        const next = prev.slice();
                        const ss = { ...next[sIdx] };
                        const ls = ss.linkedSheets.slice();
                        ls[lIdx] = {
                          ...ls[lIdx],
                          ui: { loadedTabs: tabs, loading: false },
                        };
                        ss.linkedSheets = ls;
                        next[sIdx] = ss;
                        return next;
                      });
                    } catch {
                      // ignore
                    }
                  })
                );
              })
            );
          }
        }
      } catch {
        // empty form
      }
      setHydrated(true);
    })();
  }, []);

  // Auto-save (debounced)
  useEffect(() => {
    if (!hydrated || noMasterSheet) return;
    if (sections.length === 0) return;
    const hasAnything = sections.some(
      (s) =>
        s.name ||
        s.linkedSheets.some((l) => l.url || l.nickname)
    );
    if (!hasAnything) return;

    const handle = setTimeout(() => {
      const payload = { sections: sections.map(toConfigSection) };
      fetch("/api/multi-sync/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      })
        .then((r) => {
          if (r.ok) setHasSaved(true);
        })
        .catch(() => {});
    }, 800);
    return () => clearTimeout(handle);
  }, [sections, hydrated, noMasterSheet]);

  // Mutators

  const updateSection = useCallback(
    (idx: number, patch: Partial<SectionState>) => {
      setSections((prev) => {
        const next = prev.slice();
        next[idx] = { ...next[idx], ...patch };
        return next;
      });
    },
    []
  );

  const updateLinkedSheet = useCallback(
    (sIdx: number, lIdx: number, patch: Partial<LinkedSheetState>) => {
      setSections((prev) => {
        const next = prev.slice();
        const sec = { ...next[sIdx] };
        const ls = sec.linkedSheets.slice();
        ls[lIdx] = { ...ls[lIdx], ...patch };
        sec.linkedSheets = ls;
        next[sIdx] = sec;
        return next;
      });
    },
    []
  );

  async function addSection() {
    if (noMasterSheet) return;
    setBootstrapping(true);
    try {
      const slot = await claimSlot();
      setSections((prev) => [...prev, makeSection(slot)]);
    } catch (err) {
      console.error(err);
      alert(
        err instanceof Error
          ? `Couldn't claim slot: ${err.message}`
          : "Couldn't claim slot"
      );
    } finally {
      setBootstrapping(false);
    }
  }

  function removeSection(idx: number) {
    setSections((prev) => {
      const removed = prev[idx];
      if (removed) {
        setSectionResult(removed.id, null);
        setSectionError(removed.id, null);
        setSectionRunning(removed.id, false);
      }
      return prev.filter((_, i) => i !== idx);
    });
  }

  function addLinkedSheet(sectionIdx: number) {
    setSections((prev) => {
      const next = prev.slice();
      const sec = { ...next[sectionIdx] };
      sec.linkedSheets = [...sec.linkedSheets, emptyLinkedSheet()];
      next[sectionIdx] = sec;
      return next;
    });
  }

  function removeLinkedSheet(sIdx: number, lIdx: number) {
    setSections((prev) => {
      const next = prev.slice();
      const sec = { ...next[sIdx] };
      sec.linkedSheets =
        sec.linkedSheets.length === 1
          ? [emptyLinkedSheet()]
          : sec.linkedSheets.filter((_, i) => i !== lIdx);
      next[sIdx] = sec;
      return next;
    });
  }

  async function handleFetchTabs(sIdx: number, lIdx: number) {
    const section = sections[sIdx];
    const linked = section?.linkedSheets[lIdx];
    if (!linked?.url || !section) return;
    setSectionError(section.id, null);
    updateLinkedSheet(sIdx, lIdx, {
      ui: { ...linked.ui, loading: true },
    });
    try {
      const tabs = await fetchTabsForUrl(linked.url);
      updateLinkedSheet(sIdx, lIdx, {
        ui: { loadedTabs: tabs, loading: false },
      });
    } catch (err) {
      setSectionError(
        section.id,
        err instanceof Error ? err.message : "Failed to fetch tabs"
      );
      updateLinkedSheet(sIdx, lIdx, {
        ui: { loadedTabs: [], loading: false },
      });
    }
  }

  async function handleClearSaved() {
    setErrors(new Map());
    setResults(new Map());
    try {
      await fetch("/api/multi-sync/config", { method: "DELETE" });
    } catch {
      // best-effort
    }
    setSections([]);
    setHasSaved(false);
  }

  function validateSection(s: SectionState): string | null {
    if (!s.name.trim()) return "Section needs a name.";
    if (!s.masterTabName.trim()) return "Master tab name is required.";
    if (!s.presentInColumnName.trim())
      return "Present In column name is required.";
    if (!s.linkedSheets.length) return "Add at least one linked sheet.";
    for (let i = 0; i < s.linkedSheets.length; i++) {
      const l = s.linkedSheets[i];
      if (!l.url) return `Linked sheet #${i + 1}: URL is required.`;
      if (!l.nickname.trim())
        return `Linked sheet #${i + 1}: nickname is required.`;
      if (!l.tabName)
        return `Linked sheet #${i + 1}: pick a tab (use Fetch tabs first).`;
    }
    return null;
  }

  async function handleRunSection(idx: number) {
    const section = sections[idx];
    if (!section) return;
    const err = validateSection(section);
    if (err) {
      setSectionError(section.id, err);
      return;
    }
    setSectionError(section.id, null);
    setSectionRunning(section.id, true);
    try {
      const payload = { sections: [toConfigSection(section)] };
      const res = await fetch("/api/multi-sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        setSectionError(section.id, data.error || "Failed to run");
        return;
      }
      const sec = data.sections?.[0];
      if (sec) setSectionResult(section.id, sec as MultiSyncSectionResult);
    } catch {
      setSectionError(section.id, "Network error");
    } finally {
      setSectionRunning(section.id, false);
    }
  }

  async function handleRunAll(e: React.FormEvent) {
    e.preventDefault();
    let anyFailed = false;
    for (const s of sections) {
      const err = validateSection(s);
      if (err) {
        setSectionError(s.id, err);
        anyFailed = true;
      } else {
        setSectionError(s.id, null);
      }
    }
    if (anyFailed) return;
    setRunningSections(new Set(sections.map((s) => s.id)));
    try {
      const payload = { sections: sections.map(toConfigSection) };
      const res = await fetch("/api/multi-sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) {
        const msg = data.error || "Failed to run";
        for (const s of sections) setSectionError(s.id, msg);
        return;
      }
      const newResults = new Map<string, MultiSyncSectionResult>();
      for (const r of (data.sections ?? []) as MultiSyncSectionResult[]) {
        newResults.set(r.sectionId, r);
      }
      setResults((prev) => {
        const merged = new Map(prev);
        for (const [k, v] of newResults) merged.set(k, v);
        return merged;
      });
    } catch {
      for (const s of sections) setSectionError(s.id, "Network error");
    } finally {
      setRunningSections(new Set());
    }
  }

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Multi Sync
        </h1>
        <p className="text-muted text-sm mt-1">
          The basic Present In sync, multiplied. Each section is independent:
          its own linked sheets, its own Master tab, and its own uniquely-named
          Present In column (default{" "}
          <code className="bg-card px-1 rounded">Present In - N</code>) that
          will never collide with the basic sync or other sections. Slot
          numbers are permanent — deleting a section doesn&apos;t free up its
          number.
        </p>
      </div>

      {noMasterSheet && (
        <div className="bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          Set a master sheet on the Sheets page to use Multi Sync.
        </div>
      )}

      {hasSaved && !noMasterSheet && (
        <div className="flex items-center justify-between gap-3 bg-card border border-border rounded-lg px-4 py-2.5 text-xs text-muted">
          <span>
            Configuration saved to your master sheet — it&apos;ll auto-load
            next time.
          </span>
          <button
            type="button"
            onClick={handleClearSaved}
            className="text-danger hover:underline font-medium cursor-pointer whitespace-nowrap"
          >
            Clear saved
          </button>
        </div>
      )}

      <form onSubmit={handleRunAll} className="space-y-5">
        {sections.map((section, sIdx) => (
          <SectionCard
            key={section.id}
            section={section}
            index={sIdx}
            total={sections.length}
            running={runningSections.has(section.id)}
            error={errors.get(section.id) ?? null}
            result={results.get(section.id) ?? null}
            onChangeSection={(patch) => updateSection(sIdx, patch)}
            onRemove={() => removeSection(sIdx)}
            onAddLinkedSheet={() => addLinkedSheet(sIdx)}
            onRemoveLinkedSheet={(lIdx) => removeLinkedSheet(sIdx, lIdx)}
            onChangeLinkedSheet={(lIdx, patch) =>
              updateLinkedSheet(sIdx, lIdx, patch)
            }
            onFetchTabs={(lIdx) => handleFetchTabs(sIdx, lIdx)}
            onRunSection={() => handleRunSection(sIdx)}
          />
        ))}

        {sections.length === 0 && !noMasterSheet && (
          <div className="bg-card border border-border rounded-lg p-8 text-center text-sm text-muted">
            No sections yet. Click <span className="font-medium">+ Add section</span>{" "}
            to create one.
          </div>
        )}

        <div className="flex items-center justify-between gap-3 flex-wrap">
          <button
            type="button"
            onClick={addSection}
            disabled={bootstrapping || noMasterSheet}
            className="text-sm font-medium px-3 py-1.5 rounded-md border border-border bg-background hover:bg-card disabled:opacity-50 transition-colors cursor-pointer"
          >
            {bootstrapping ? "Claiming slot..." : "+ Add section"}
          </button>
          {sections.length > 1 && (
            <button
              type="submit"
              disabled={runningSections.size > 0}
              className="bg-primary text-white px-5 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer"
            >
              {runningSections.size > 0
                ? "Running..."
                : `Run all ${sections.length} sections`}
            </button>
          )}
        </div>
      </form>
    </div>
  );
}

// ---------- SectionCard ----------

function SectionCard({
  section,
  index,
  total,
  running,
  error,
  result,
  onChangeSection,
  onRemove,
  onAddLinkedSheet,
  onRemoveLinkedSheet,
  onChangeLinkedSheet,
  onFetchTabs,
  onRunSection,
}: {
  section: SectionState;
  index: number;
  total: number;
  running: boolean;
  error: string | null;
  result: MultiSyncSectionResult | null;
  onChangeSection: (patch: Partial<SectionState>) => void;
  onRemove: () => void;
  onAddLinkedSheet: () => void;
  onRemoveLinkedSheet: (linkedIdx: number) => void;
  onChangeLinkedSheet: (
    linkedIdx: number,
    patch: Partial<LinkedSheetState>
  ) => void;
  onFetchTabs: (linkedIdx: number) => void;
  onRunSection: () => void;
}) {
  return (
    <div className="bg-card border border-border rounded-lg p-5 space-y-5">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div className="flex items-center gap-2 flex-1 min-w-0">
          <span className="text-xs text-muted uppercase tracking-wide">
            Section {index + 1}
          </span>
          <span className="text-[10px] text-muted bg-background px-1.5 py-0.5 rounded-md border border-border">
            slot {section.slot}
          </span>
          <input
            type="text"
            value={section.name}
            onChange={(e) => onChangeSection({ name: e.target.value })}
            placeholder="Section name"
            className="flex-1 min-w-0 border border-border rounded-md px-2 py-1 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </div>
        {total > 1 && (
          <button
            type="button"
            onClick={onRemove}
            className="text-xs text-danger hover:underline cursor-pointer whitespace-nowrap"
          >
            Remove section
          </button>
        )}
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <label className="flex flex-col gap-1 text-xs">
          <span className="text-muted uppercase tracking-wide">
            Master tab name
          </span>
          <input
            type="text"
            value={section.masterTabName}
            onChange={(e) =>
              onChangeSection({ masterTabName: e.target.value })
            }
            placeholder={`Multi Master - ${section.slot}`}
            className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </label>
        <label className="flex flex-col gap-1 text-xs">
          <span className="text-muted uppercase tracking-wide">
            Present In column name
          </span>
          <input
            type="text"
            value={section.presentInColumnName}
            onChange={(e) =>
              onChangeSection({ presentInColumnName: e.target.value })
            }
            placeholder={`Present In - ${section.slot}`}
            className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
          />
        </label>
      </div>

      {/* Linked sheets list */}
      <div className="space-y-3">
        <p className="text-xs font-medium text-muted uppercase tracking-wide">
          Linked sheets ({section.linkedSheets.length})
        </p>
        {section.linkedSheets.map((linked, lIdx) => (
          <LinkedSheetRow
            key={lIdx}
            linked={linked}
            index={lIdx}
            total={section.linkedSheets.length}
            onChange={(patch) => onChangeLinkedSheet(lIdx, patch)}
            onRemove={() => onRemoveLinkedSheet(lIdx)}
            onFetch={() => onFetchTabs(lIdx)}
          />
        ))}
        <button
          type="button"
          onClick={onAddLinkedSheet}
          className="text-xs font-medium px-3 py-1.5 rounded-md border border-dashed border-border bg-background hover:bg-card transition-colors cursor-pointer"
        >
          + Add another linked sheet
        </button>
      </div>

      {/* Per-section Run + inline error */}
      <div className="border-t border-border pt-4 flex items-center justify-between gap-3 flex-wrap">
        {error ? (
          <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2 flex-1 min-w-0">
            {error}
          </p>
        ) : (
          <span className="text-xs text-muted">
            Runs this section only. Other sections are untouched.
          </span>
        )}
        <button
          type="button"
          onClick={onRunSection}
          disabled={running}
          className="bg-primary text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-primary-hover disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap"
        >
          {running ? "Running..." : "Run section"}
        </button>
      </div>

      {result && <InlineResult result={result} />}
    </div>
  );
}

// ---------- LinkedSheetRow ----------

function LinkedSheetRow({
  linked,
  index,
  total,
  onChange,
  onRemove,
  onFetch,
}: {
  linked: LinkedSheetState;
  index: number;
  total: number;
  onChange: (patch: Partial<LinkedSheetState>) => void;
  onRemove: () => void;
  onFetch: () => void;
}) {
  return (
    <div className="border border-border rounded-md p-3 space-y-3 bg-background/50">
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-baseline gap-2">
          <span className="text-xs text-muted">Sheet #{index + 1}</span>
          <span className="text-[10px] text-muted">
            last synced: {formatLastSynced(linked.lastSynced)}
          </span>
        </div>
        {total > 1 && (
          <button
            type="button"
            onClick={onRemove}
            className="text-xs text-danger hover:underline cursor-pointer"
          >
            Remove sheet
          </button>
        )}
      </div>

      <input
        type="text"
        value={linked.nickname}
        onChange={(e) => onChange({ nickname: e.target.value })}
        placeholder="Nickname (e.g. Newsletter)"
        className="w-full border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
      />

      <div className="flex flex-col sm:flex-row gap-2">
        <input
          type="url"
          value={linked.url}
          onChange={(e) =>
            onChange({
              url: e.target.value,
              tabName: "",
              ui: { loadedTabs: [], loading: false },
            })
          }
          placeholder="https://docs.google.com/spreadsheets/d/..."
          className="flex-1 min-w-0 border border-border rounded-md px-3 py-2 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
        />
        <button
          type="button"
          onClick={onFetch}
          disabled={!linked.url || linked.ui.loading}
          className="bg-foreground text-background px-3 py-2 rounded-md text-sm font-medium hover:opacity-80 disabled:opacity-50 transition-opacity cursor-pointer whitespace-nowrap"
        >
          {linked.ui.loading ? "Loading..." : "Fetch tabs"}
        </button>
      </div>

      {linked.ui.loadedTabs.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <label className="flex flex-col gap-1 text-xs">
            <span className="text-muted uppercase tracking-wide">Tab</span>
            <select
              value={linked.tabName}
              onChange={(e) => onChange({ tabName: e.target.value })}
              className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            >
              <option value="">— select —</option>
              {linked.ui.loadedTabs.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </label>
          <label className="flex flex-col gap-1 text-xs">
            <span className="text-muted uppercase tracking-wide">
              Email column
            </span>
            <input
              type="text"
              value={linked.emailColumn}
              onChange={(e) =>
                onChange({ emailColumn: e.target.value || "auto" })
              }
              placeholder="auto"
              className="border border-border rounded-md px-2 py-1.5 text-sm bg-background focus:outline-none focus:ring-2 focus:ring-primary/30 focus:border-primary"
            />
          </label>
        </div>
      )}
    </div>
  );
}

// ---------- InlineResult ----------

function InlineResult({ result }: { result: MultiSyncSectionResult }) {
  const failed = !!result.error;
  return (
    <div
      className={`border-t pt-4 ${
        failed ? "border-danger/30" : "border-border"
      }`}
    >
      <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2 mb-3">
        <h3 className="font-medium text-sm">Last run</h3>
        {result.masterSpreadsheetUrl && (
          <a
            href={result.masterSpreadsheetUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="text-sm text-primary hover:underline font-medium"
          >
            Open {result.masterTabName} →
          </a>
        )}
      </div>

      {failed && (
        <p className="text-sm text-danger bg-danger/10 rounded-md px-3 py-2 mb-3">
          {result.error}
        </p>
      )}

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-sm">
        <Stat label="Unique emails" value={result.totalUniqueEmails} />
        <Stat label="Sheets" value={result.linkedSheets.length} />
        <Stat
          label="Present In"
          value={result.presentInWritten ? 1 : 0}
          accent="muted"
          hint={result.presentInColumnName}
        />
        <Stat
          label="Errors"
          value={result.linkedSheets.filter((t) => t.error).length}
          accent={
            result.linkedSheets.some((t) => t.error) ? "danger" : "muted"
          }
        />
      </div>

      <div className="mt-3 space-y-1">
        {result.linkedSheets.map((t, i) => (
          <div
            key={i}
            className="flex items-center justify-between text-xs py-1 border-b border-border last:border-0"
          >
            <span className="font-medium">{t.nickname}</span>
            {t.error ? (
              <span className="text-danger">{t.error}</span>
            ) : (
              <span className="text-muted">
                {t.rowsRead} rows · {t.emailsFound} emails
              </span>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

// ---------- Stat ----------

function Stat({
  label,
  value,
  accent,
  hint,
}: {
  label: string;
  value: number;
  accent?: "success" | "danger" | "muted";
  hint?: string;
}) {
  const valueClass =
    accent === "success"
      ? "text-success"
      : accent === "danger"
        ? "text-danger"
        : accent === "muted"
          ? "text-muted"
          : "text-foreground";
  return (
    <div className="bg-background border border-border rounded-md px-3 py-2">
      <p className="text-xs text-muted">{label}</p>
      <p className={`text-lg font-semibold ${valueClass}`}>{value}</p>
      {hint && (
        <p className="text-[10px] text-muted mt-0.5 truncate">{hint}</p>
      )}
    </div>
  );
}
