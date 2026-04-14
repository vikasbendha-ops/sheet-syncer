import cron from "node-cron";
import { runFullSync } from "./sync-engine";

const DEFAULT_SCHEDULE = "*/30 * * * *"; // every 30 minutes

declare global {
  // eslint-disable-next-line no-var
  var __sheetSyncerCronStarted: boolean | undefined;
}

export function startCron(): void {
  // Guard against double-registration during Next.js hot reload
  if (global.__sheetSyncerCronStarted) {
    console.log("[cron] already started; skipping");
    return;
  }

  const schedule = process.env.CRON_SCHEDULE || DEFAULT_SCHEDULE;

  if (!cron.validate(schedule)) {
    console.error(`[cron] invalid CRON_SCHEDULE: "${schedule}"`);
    return;
  }

  cron.schedule(schedule, async () => {
    const startedAt = new Date().toISOString();
    console.log(`[cron] sync triggered at ${startedAt}`);

    const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
    const masterSheetId = process.env.MASTER_SHEET_ID;

    if (!refreshToken || !masterSheetId) {
      console.warn(
        "[cron] skipping: GOOGLE_REFRESH_TOKEN or MASTER_SHEET_ID env var not set"
      );
      return;
    }

    try {
      const result = await runFullSync(refreshToken, masterSheetId);
      console.log(
        `[cron] done: ${result.sheetsProcessed} sheets, ${result.totalEmails} emails, ${result.errors.length} errors`
      );
      if (result.errors.length > 0) {
        console.warn("[cron] errors:", result.errors);
      }
    } catch (err) {
      console.error("[cron] sync failed:", err);
    }
  });

  global.__sheetSyncerCronStarted = true;
  console.log(`[cron] scheduled with "${schedule}"`);
}
