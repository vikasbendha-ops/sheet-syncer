import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Privacy Policy — Sheet Syncer",
  description: "How Sheet Syncer handles your data",
};

const APP_NAME = "Sheet Syncer";
const CONTACT_EMAIL = "connect@vikasbendha.com";
const EFFECTIVE_DATE = "April 14, 2026";

export default function PrivacyPolicy() {
  return (
    <article className="prose-like space-y-6">
      <header>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Privacy Policy
        </h1>
        <p className="text-sm text-muted mt-1">
          Effective: {EFFECTIVE_DATE}
        </p>
      </header>

      <Section title="1. Overview">
        <p>
          {APP_NAME} (&quot;the App&quot;, &quot;we&quot;, &quot;us&quot;) is a
          web application that helps users consolidate email addresses from
          multiple Google Sheets into a single master Google Sheet. This
          Privacy Policy explains what data we access, how we use it, and how
          we protect it.
        </p>
      </Section>

      <Section title="2. Information We Access">
        <p>When you sign in with Google, we request access to the following:</p>
        <ul className="list-disc pl-6 space-y-1 mt-2">
          <li>
            <strong>Your email address and basic profile info</strong> (via the
            <code className="mx-1 text-xs bg-background px-1 py-0.5 rounded">
              userinfo.email
            </code>
            and
            <code className="mx-1 text-xs bg-background px-1 py-0.5 rounded">
              openid
            </code>
            scopes) — used solely to display your identity in the app interface.
          </li>
          <li>
            <strong>Your Google Sheets content</strong> (via the
            <code className="mx-1 text-xs bg-background px-1 py-0.5 rounded">
              spreadsheets
            </code>
            scope) — used to read email data from sheets you explicitly link in
            the app and to write the consolidated master sheet and a
            &quot;Present In&quot; column.
          </li>
        </ul>
      </Section>

      <Section title="3. How We Use Your Information">
        <ul className="list-disc pl-6 space-y-1">
          <li>
            To authenticate you and maintain your signed-in session.
          </li>
          <li>
            To read email addresses from Google Sheets you explicitly link.
          </li>
          <li>
            To write the merged data into the master Google Sheet you configure.
          </li>
          <li>
            To display your linked sheets and sync status in the app.
          </li>
        </ul>
        <p className="mt-2">
          We do not use your data for any other purpose. We do not sell, trade,
          rent, or share your data with third parties for advertising or
          marketing.
        </p>
      </Section>

      <Section title="4. Google API Services User Data Policy — Limited Use">
        <p>
          {APP_NAME}&apos;s use and transfer of information received from Google
          APIs adheres to the{" "}
          <a
            href="https://developers.google.com/terms/api-services-user-data-policy#additional_requirements_for_specific_api_scopes"
            target="_blank"
            rel="noopener noreferrer"
            className="text-primary hover:underline"
          >
            Google API Services User Data Policy
          </a>
          , including the Limited Use requirements. Specifically:
        </p>
        <ul className="list-disc pl-6 space-y-1 mt-2">
          <li>
            We only use Google user data to provide or improve
            user-facing features of the App that are prominent in the
            user experience.
          </li>
          <li>
            We do not transfer Google user data to third parties except as
            necessary to operate user-facing features (we do not transfer
            it to any third party).
          </li>
          <li>
            We do not use Google user data for serving advertisements.
          </li>
          <li>
            We do not allow humans to read Google user data unless we
            obtain explicit consent, it is necessary for security, to
            comply with law, or the data is aggregated and anonymized.
          </li>
        </ul>
      </Section>

      <Section title="5. Data Storage and Security">
        <p>
          We do not operate any database. All app configuration (the list of
          your linked sheets, master sheet reference, sync timestamps) is
          stored inside your own Google Sheet in a dedicated
          <code className="mx-1 text-xs bg-background px-1 py-0.5 rounded">
            _config
          </code>
          tab that you control.
        </p>
        <p className="mt-2">
          Your Google OAuth tokens (refresh token and access token) are stored
          only in an <strong>encrypted HTTP-only session cookie</strong>{" "}
          (AES-256-GCM) in your browser. Tokens are never written to a server
          database.
        </p>
        <p className="mt-2">
          All communication with Google APIs occurs over HTTPS. The app is
          hosted on Vercel and runs only while processing your requests.
        </p>
      </Section>

      <Section title="6. Data Retention and Deletion">
        <p>You can delete your data at any time:</p>
        <ul className="list-disc pl-6 space-y-1 mt-2">
          <li>
            <strong>Sign out</strong> from within the app — your session cookie
            and OAuth tokens are immediately destroyed.
          </li>
          <li>
            <strong>Revoke access</strong> to the App at any time in your{" "}
            <a
              href="https://myaccount.google.com/permissions"
              target="_blank"
              rel="noopener noreferrer"
              className="text-primary hover:underline"
            >
              Google Account permissions page
            </a>
            . This immediately invalidates our refresh token.
          </li>
          <li>
            <strong>Delete your data</strong> directly in Google Sheets — you
            own your spreadsheets and can remove any data the App has written.
          </li>
        </ul>
        <p className="mt-2">
          We retain no copy of your data outside your own Google Sheets and the
          browser session cookie.
        </p>
      </Section>

      <Section title="7. Third-Party Services">
        <p>
          The App relies on the following third parties to function:
        </p>
        <ul className="list-disc pl-6 space-y-1 mt-2">
          <li>
            <strong>Google APIs</strong> — for authentication and Sheets access
            (<a
              href="https://policies.google.com/privacy"
              target="_blank"
              rel="noopener noreferrer"
              className="text-primary hover:underline"
            >
              Google Privacy Policy
            </a>
            )
          </li>
          <li>
            <strong>Vercel</strong> — our hosting provider (
            <a
              href="https://vercel.com/legal/privacy-policy"
              target="_blank"
              rel="noopener noreferrer"
              className="text-primary hover:underline"
            >
              Vercel Privacy Policy
            </a>
            )
          </li>
        </ul>
      </Section>

      <Section title="8. Children's Privacy">
        <p>
          {APP_NAME} is not directed to children under 13, and we do not
          knowingly collect personal information from children under 13.
        </p>
      </Section>

      <Section title="9. Changes to This Policy">
        <p>
          We may update this Privacy Policy from time to time. The
          &quot;Effective&quot; date at the top will reflect the most recent
          update.
        </p>
      </Section>

      <Section title="10. Contact">
        <p>
          If you have questions about this Privacy Policy or our data
          practices, contact us at{" "}
          <a
            href={`mailto:${CONTACT_EMAIL}`}
            className="text-primary hover:underline"
          >
            {CONTACT_EMAIL}
          </a>
          .
        </p>
      </Section>
    </article>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="space-y-2 text-sm sm:text-base leading-relaxed">
      <h2 className="text-lg sm:text-xl font-semibold mt-6">{title}</h2>
      <div className="text-foreground/90">{children}</div>
    </section>
  );
}
