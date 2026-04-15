import Link from "next/link";

export default function Footer() {
  return (
    <footer className="border-t border-border bg-card mt-12">
      <div className="max-w-5xl mx-auto px-3 sm:px-4 py-6 flex flex-col sm:flex-row items-center justify-between gap-3 text-sm text-muted">
        <p>© {new Date().getFullYear()} Sheet Syncer</p>
        <nav className="flex gap-4 sm:gap-6">
          <Link href="/privacy" className="hover:text-foreground transition-colors">
            Privacy
          </Link>
          <Link href="/terms" className="hover:text-foreground transition-colors">
            Terms
          </Link>
        </nav>
      </div>
    </footer>
  );
}
