import Link from "next/link";
import AuthStatus from "./auth-status";

export default function Nav() {
  return (
    <header className="border-b border-border bg-card">
      <div className="max-w-5xl mx-auto px-3 sm:px-4 flex items-center justify-between h-14 gap-3">
        <Link
          href="/"
          className="font-semibold text-base sm:text-lg tracking-tight shrink-0"
        >
          Sheet Syncer
        </Link>
        <div className="flex items-center gap-3 sm:gap-6 min-w-0">
          <nav className="flex gap-3 sm:gap-6 text-sm shrink-0">
            <Link
              href="/"
              className="text-muted hover:text-foreground transition-colors"
            >
              Dashboard
            </Link>
            <Link
              href="/sheets"
              className="text-muted hover:text-foreground transition-colors"
            >
              Sheets
            </Link>
            <Link
              href="/domain-analyzer"
              className="text-muted hover:text-foreground transition-colors whitespace-nowrap"
            >
              <span className="hidden sm:inline">Domain Analyzer</span>
              <span className="sm:hidden">Domains</span>
            </Link>
            <Link
              href="/email-finder"
              className="text-muted hover:text-foreground transition-colors whitespace-nowrap"
            >
              <span className="hidden sm:inline">Email Finder</span>
              <span className="sm:hidden">Finder</span>
            </Link>
          </nav>
          <AuthStatus />
        </div>
      </div>
    </header>
  );
}
