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
        className="text-sm bg-primary text-white px-3 py-1.5 rounded-md hover:bg-primary-hover transition-colors"
      >
        Sign in with Google
      </a>
    );
  }

  return (
    <div className="flex items-center gap-3 text-sm">
      <span className="text-muted">{auth.email}</span>
      <button
        onClick={handleLogout}
        disabled={loggingOut}
        className="text-muted hover:text-foreground transition-colors cursor-pointer"
      >
        Sign out
      </button>
    </div>
  );
}
