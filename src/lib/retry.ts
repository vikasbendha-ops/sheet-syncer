/**
 * Retries a Google API call with exponential backoff when a rate-limit error
 * (HTTP 429) or transient 5xx is encountered. Other errors propagate.
 */
export async function withRetry<T>(
  fn: () => Promise<T>,
  options: { maxAttempts?: number; initialDelayMs?: number; maxDelayMs?: number } = {}
): Promise<T> {
  const maxAttempts = options.maxAttempts ?? 6;
  const initialDelay = options.initialDelayMs ?? 1000;
  const maxDelay = options.maxDelayMs ?? 32000;

  let delay = initialDelay;
  let lastErr: unknown;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      if (!isRetryable(err) || attempt === maxAttempts) {
        throw err;
      }
      const jitter = Math.random() * 500;
      await sleep(Math.min(delay + jitter, maxDelay));
      delay = Math.min(delay * 2, maxDelay);
    }
  }
  throw lastErr;
}

function isRetryable(err: unknown): boolean {
  if (typeof err !== "object" || err === null) return false;
  const e = err as {
    code?: number | string;
    status?: number;
    response?: { status?: number };
    errors?: Array<{ reason?: string }>;
  };
  const code = Number(e.code ?? e.status ?? e.response?.status ?? 0);
  if (code === 429) return true;
  if (code >= 500 && code < 600) return true;
  // Some Google API errors expose reason instead of status
  if (e.errors?.some((x) => x.reason === "rateLimitExceeded" || x.reason === "userRateLimitExceeded")) {
    return true;
  }
  return false;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
