"use client";

import { useEffect, useState } from "react";

interface AuthState {
  authenticated: boolean;
  email?: string;
}

export default function AuthStatus() {
  const [auth, setAuth] = useState<AuthState | null>(null);
  const [loggingOut, setLoggingOut] = useState(false);

  useEffect(() => {
    fetch("/api/auth/status")
      .then((res) => res.json())
      .then(setAuth)
      .catch(() => setAuth({ authenticated: false }));
  }, []);

  async function handleLogout() {
    setLoggingOut(true);
    await fetch("/api/auth/logout", { method: "POST" });
    setAuth({ authenticated: false });
    setLoggingOut(false);
  }

  if (!auth) return null;

  if (!auth.authenticated) {
    return (
      <a
        href="/api/auth/login"
        className="text-xs sm:text-sm bg-primary text-white px-2.5 sm:px-3 py-1.5 rounded-md hover:bg-primary-hover transition-colors whitespace-nowrap"
      >
        <span className="hidden sm:inline">Sign in with Google</span>
        <span className="sm:hidden">Sign in</span>
      </a>
    );
  }

  return (
    <div className="flex items-center gap-2 sm:gap-3 text-sm min-w-0">
      <span className="text-muted truncate max-w-[120px] sm:max-w-none hidden sm:inline">
        {auth.email}
      </span>
      <button
        onClick={handleLogout}
        disabled={loggingOut}
        className="text-muted hover:text-foreground transition-colors cursor-pointer whitespace-nowrap"
      >
        Sign out
      </button>
    </div>
  );
}
