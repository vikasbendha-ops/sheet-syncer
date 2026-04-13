import Link from "next/link";

export default function Nav() {
  return (
    <header className="border-b border-border bg-card">
      <div className="max-w-5xl mx-auto px-4 flex items-center justify-between h-14">
        <Link href="/" className="font-semibold text-lg tracking-tight">
          Sheet Syncer
        </Link>
        <nav className="flex gap-6 text-sm">
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
        </nav>
      </div>
    </header>
  );
}
