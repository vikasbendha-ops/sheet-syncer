"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useRef, useState } from "react";
import AuthStatus from "./auth-status";

const navLinks = [
  { href: "/", label: "Dashboard" },
  { href: "/sheets", label: "Sheets" },
  { href: "/domain-analyzer", label: "Domain Analyzer" },
  { href: "/email-finder", label: "Email Finder" },
  { href: "/renewal-sync", label: "Renewal Sync" },
  { href: "/biz-tutor-sync", label: "BIZ Tutor" },
];

export default function Nav() {
  const [open, setOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);
  const pathname = usePathname();

  useEffect(() => {
    if (!open) return;
    function onClick(e: MouseEvent) {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    }
    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") setOpen(false);
    }
    document.addEventListener("mousedown", onClick);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onClick);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  return (
    <header className="border-b border-border bg-card">
      <div className="max-w-5xl mx-auto px-3 sm:px-4 flex items-center justify-between h-14 gap-3">
        <Link
          href="/"
          className="font-semibold text-base sm:text-lg tracking-tight shrink-0"
        >
          Sheet Syncer
        </Link>
        <div className="flex items-center gap-3">
          <AuthStatus />
          <div className="relative" ref={menuRef}>
            <button
              type="button"
              onClick={() => setOpen((v) => !v)}
              aria-label={open ? "Close menu" : "Open menu"}
              aria-expanded={open}
              className="p-2 -mr-2 rounded-md hover:bg-background transition-colors cursor-pointer text-foreground"
            >
              {open ? (
                <svg
                  width="20"
                  height="20"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  aria-hidden="true"
                >
                  <line x1="6" y1="6" x2="18" y2="18" />
                  <line x1="18" y1="6" x2="6" y2="18" />
                </svg>
              ) : (
                <svg
                  width="20"
                  height="20"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  aria-hidden="true"
                >
                  <line x1="4" y1="6" x2="20" y2="6" />
                  <line x1="4" y1="12" x2="20" y2="12" />
                  <line x1="4" y1="18" x2="20" y2="18" />
                </svg>
              )}
            </button>
            {open && (
              <nav className="absolute right-0 top-full mt-2 w-56 bg-card border border-border rounded-md shadow-lg overflow-hidden z-50">
                {navLinks.map((link) => {
                  const active =
                    link.href === "/"
                      ? pathname === "/"
                      : pathname === link.href ||
                        pathname.startsWith(link.href + "/");
                  return (
                    <Link
                      key={link.href}
                      href={link.href}
                      onClick={() => setOpen(false)}
                      className={`block px-4 py-2.5 text-sm transition-colors ${
                        active
                          ? "bg-background text-foreground font-medium"
                          : "text-muted hover:bg-background hover:text-foreground"
                      }`}
                    >
                      {link.label}
                    </Link>
                  );
                })}
              </nav>
            )}
          </div>
        </div>
      </div>
    </header>
  );
}
