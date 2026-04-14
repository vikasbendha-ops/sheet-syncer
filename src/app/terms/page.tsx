import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Terms of Service — Sheet Syncer",
  description: "Terms governing use of Sheet Syncer",
};

const APP_NAME = "Sheet Syncer";
const CONTACT_EMAIL = "connect@vikasbendha.com";
const EFFECTIVE_DATE = "April 14, 2026";

export default function TermsOfService() {
  return (
    <article className="space-y-6">
      <header>
        <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          Terms of Service
        </h1>
        <p className="text-sm text-muted mt-1">
          Effective: {EFFECTIVE_DATE}
        </p>
      </header>

      <Section title="1. Acceptance of Terms">
        <p>
          By accessing or using {APP_NAME} (&quot;the App&quot;), you agree to
          be bound by these Terms of Service. If you do not agree, do not use
          the App.
        </p>
      </Section>

      <Section title="2. Description of Service">
        <p>
          {APP_NAME} is a utility web application that lets you merge email
          addresses from multiple Google Sheets into a master Google Sheet,
          with a reverse-lookup &quot;Present In&quot; column written back to
          each source sheet. The App requires a Google account and permission
          to read and write your Google Sheets.
        </p>
      </Section>

      <Section title="3. Google Account and Data">
        <p>
          To use the App, you must sign in with a Google account and grant the
          permissions requested on the OAuth consent screen. Your use of
          Google&apos;s services through the App is also governed by{" "}
          <a
            href="https://policies.google.com/terms"
            target="_blank"
            rel="noopener noreferrer"
            className="text-primary hover:underline"
          >
            Google&apos;s Terms of Service
          </a>
          . You retain full ownership of your Google Sheets and all content
          within them.
        </p>
      </Section>

      <Section title="4. User Responsibilities">
        <ul className="list-disc pl-6 space-y-1">
          <li>
            You are responsible for the Google Sheets you link to the App and
            any data they contain.
          </li>
          <li>
            You must have the legal right to access and modify any sheets you
            link. Do not link sheets you are not authorized to use.
          </li>
          <li>
            You are responsible for maintaining the security of your Google
            account.
          </li>
          <li>
            You agree not to use the App to violate any applicable law, to
            harvest personal data without consent, or to send spam.
          </li>
        </ul>
      </Section>

      <Section title="5. Prohibited Use">
        <p>You may not:</p>
        <ul className="list-disc pl-6 space-y-1 mt-2">
          <li>Reverse engineer, decompile, or attempt to extract source code beyond what is publicly available.</li>
          <li>Use the App to process data you do not have the right to process.</li>
          <li>Interfere with or disrupt the App or the servers and networks running it.</li>
          <li>Use automated means to access the App outside of normal interactive use.</li>
        </ul>
      </Section>

      <Section title="6. Intellectual Property">
        <p>
          The App&apos;s code, design, and branding are the property of the
          developer. Your content (the Google Sheets data) remains yours —
          {APP_NAME} claims no ownership over any data you process through
          the App.
        </p>
      </Section>

      <Section title="7. Disclaimer of Warranty">
        <p>
          The App is provided <strong>&quot;as is&quot;</strong> and{" "}
          <strong>&quot;as available&quot;</strong> without any warranty,
          express or implied, including but not limited to merchantability,
          fitness for a particular purpose, or non-infringement. We do not
          warrant that the App will be uninterrupted, error-free, or free of
          harmful components.
        </p>
      </Section>

      <Section title="8. Limitation of Liability">
        <p>
          To the fullest extent permitted by law, the developer shall not be
          liable for any indirect, incidental, special, consequential, or
          punitive damages, including loss of data, arising from your use of
          the App. <strong>Always keep a backup of any important Google
          Sheets before linking them.</strong>
        </p>
      </Section>

      <Section title="9. Termination">
        <p>
          You can stop using the App at any time by signing out or revoking
          access in your{" "}
          <a
            href="https://myaccount.google.com/permissions"
            target="_blank"
            rel="noopener noreferrer"
            className="text-primary hover:underline"
          >
            Google Account permissions page
          </a>
          . We may suspend or terminate access at our discretion, including
          if you violate these Terms.
        </p>
      </Section>

      <Section title="10. Changes to These Terms">
        <p>
          We may update these Terms from time to time. Continued use of the
          App after changes constitutes acceptance of the updated Terms.
        </p>
      </Section>

      <Section title="11. Contact">
        <p>
          Questions about these Terms? Contact{" "}
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
