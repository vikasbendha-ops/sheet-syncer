/**
 * Returns the public-facing base URL of the app.
 * Falls back to derive from GOOGLE_REDIRECT_URI, then to the request's origin.
 * Set APP_URL explicitly when running behind a reverse proxy (Hostinger, Nginx, etc.)
 * where the request URL would otherwise be the internal address (e.g. http://0.0.0.0:3000).
 */
export function getAppUrl(requestUrl?: string): string {
  if (process.env.APP_URL) {
    return stripTrailingSlash(process.env.APP_URL);
  }

  if (process.env.GOOGLE_REDIRECT_URI) {
    try {
      const u = new URL(process.env.GOOGLE_REDIRECT_URI);
      return `${u.protocol}//${u.host}`;
    } catch {
      // ignore
    }
  }

  if (requestUrl) {
    try {
      return new URL(requestUrl).origin;
    } catch {
      // ignore
    }
  }

  return "http://localhost:3000";
}

function stripTrailingSlash(s: string): string {
  return s.endsWith("/") ? s.slice(0, -1) : s;
}
