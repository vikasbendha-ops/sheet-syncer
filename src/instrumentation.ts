export async function register() {
  // Only run in the Node.js runtime (skip Edge runtime)
  if (process.env.NEXT_RUNTIME !== "nodejs") return;

  const { startCron } = await import("@/lib/cron");
  startCron();
}
