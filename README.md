# Investment Team Updates

A phone-friendly web app for logging investment updates and emailing your team a polished summary.

## What it does

- Lets team members sign in with email plus a shared workspace password
- Saves each investment update in `data/investments.json`
- Sends a formatted summary email through Resend
- Can summarize uploaded PDF decks into investment notes using OpenAI
- Can manually check a configured Microsoft 365 mail folder for investment emails and PDF attachments
- Shows recent updates in a mobile-friendly dashboard
- Protects the update history behind a signed session cookie

## Quick start

1. Install Node.js 18 or newer.
2. Copy `.env.example` to `.env`.
3. Fill in the values in `.env`, especially:
   - `DATA_DIR`
   - `TEAM_PASSWORD`
   - `SESSION_SECRET`
   - `TEAM_ALLOWED_EMAILS`
   - `TEAM_EMAILS`
   - `FROM_EMAIL`
   - `UPDATE_REQUEST_FROM_EMAIL`
   - `UPDATE_REQUEST_REPLY_TO_EMAIL`
   - `RESEND_API_KEY`
   - `OPENAI_API_KEY`

```bash
node server.js
```

5. Open `http://localhost:3000`

## Environment variables

- `TEAM_PASSWORD`: shared password your team uses to sign in
- `SESSION_SECRET`: long random string used to sign login cookies
- `DATA_DIR`: where updates are stored on disk
- `TEAM_ALLOWED_EMAILS`: comma-separated list of allowed sign-in emails
- `TEAM_EMAILS`: default recipients for update emails
- `FROM_EMAIL`: sender identity for Resend
- `UPDATE_REQUEST_FROM_EMAIL`: sender identity for latest-update request emails, for example `Tyler@Beamanventures.com`; the domain must be verified in Resend
- `UPDATE_REQUEST_REPLY_TO_EMAIL`: inbox that receives replies to latest-update request emails, for example `Tyler@Beamanventures.com`
- `RESEND_API_KEY`: API key for Resend
- `OPENAI_API_KEY`: API key for deck summarization and investment email analysis
- `OPENAI_MODEL`: optional override for the summarization and analysis model
- `AI_EMAIL_INTAKE_ENABLED`: opt-in switch for Microsoft 365 investment email intake; defaults to disabled and must be set to `true` to enable the manual check
- `MICROSOFT_TENANT_ID`: Microsoft Entra tenant identifier used for client-credentials authentication
- `MICROSOFT_CLIENT_ID`: application client identifier used for Microsoft Graph authentication
- `MICROSOFT_CLIENT_SECRET`: application credential used for Microsoft Graph authentication; keep it out of source control and logs
- `MICROSOFT_MAILBOX_USER`: mailbox address the intake service checks directly
- `MICROSOFT_MAIL_FOLDER_NAME`: mailbox folder to check; defaults to `AI Investment Updates`
- `AI_EMAIL_MAX_MESSAGES_PER_RUN`: maximum recent messages requested per manual check; defaults to `10` and is constrained to `1` through `50`
- `AI_EMAIL_ALLOWED_SENDERS`: optional comma-separated list of exact sender email addresses accepted for analysis
- `AI_EMAIL_ALLOWED_DOMAINS`: optional comma-separated list of sender domains accepted for analysis
- `AI_EMAIL_HOUSE_DOMAINS`: comma-separated house domains that cannot provide positive company-domain match evidence; defaults to `beamanventures.com` and includes its subdomains

## Microsoft 365 investment email intake

Microsoft 365 intake is disabled by default. Microsoft 365 tenant setup, application consent, and Graph permissions are operational prerequisites managed outside this repository. Configure those prerequisites and the environment variables above before setting `AI_EMAIL_INTAKE_ENABLED=true`.

An editor starts each intake run manually from the AI Update Inbox by selecting `Check for new investment emails`. The app requests up to `AI_EMAIL_MAX_MESSAGES_PER_RUN` of the most recent messages in the configured mailbox folder and reports how many were processed, skipped, or failed.

A master editor can select `Preview intake window` to inspect that same ordered Graph message window without running intake. The preview shows message identifiers, sender and subject metadata, allowlist and intake-state outcomes, attachment metadata, and projected message/run attachment budgets. It does not reserve messages, write intake state, download attachment content, invoke analysis, or create proposals.

The intake accepts meaningful email-body text and non-inline PDF attachments. HTML email is normalized to text, and obvious signatures and quoted thread content are removed before analysis. Inline files and attachments other than PDFs are skipped.

Sender controls are optional. When `AI_EMAIL_ALLOWED_SENDERS` is populated, a message must come from one of those exact addresses. When `AI_EMAIL_ALLOWED_DOMAINS` is populated, its sender domain must be listed. When both settings are populated, both checks must pass. Leaving both empty permits any sender whose message is present in the configured folder.

House domains are separate from sender allowlists. A sender from `AI_EMAIL_HOUSE_DOMAINS` can still match an investment through explicit company or alias evidence in the subject, body, or attachment filename, but its email domain alone is never positive match evidence.

Processed messages are deduplicated using their Internet Message ID or Microsoft Graph message ID. PDF content is also deduplicated by a SHA-256 hash. Intake state and its analysis audit are stored in `ai-email-intake-state.json` under `DATA_DIR`, so later manual checks do not create duplicate proposals from previously processed sources.

One source email may describe multiple investment opportunities. Intake first partitions the email body and parsed PDFs into evidence-isolated opportunities, then runs deterministic matching and proposal analysis independently for each partition. Each pending proposal retains the common source-message key plus its own stable opportunity ID; that pair is the idempotency key for source reprocessing. The Proposal Inbox labels sibling opportunities from the same source email.

For a legacy pending proposal created before opportunity partitioning, a master editor can explicitly select `Reanalyze source into opportunities`. The server fetches the exact preserved Graph message, verifies its Internet Message ID and configured intake folder, reuses attachment hashes, and creates only pending replacement proposals. The legacy proposal is marked `superseded` only after replacements exist. Reanalysis never approves a proposal or creates an investment.

Automated intake must find explicit portfolio evidence before it creates a proposal, and the existing source-evidence safety checks still apply. Successful analysis creates a `pending` proposal only. It never approves or applies an investment update: an authorized human must review and approve or reject each proposal in the AI Update Inbox.

### Potential New Deal intake

When deterministic company and alias matching finds one confident existing investment, email content stays in the existing investment-update workflow. Competing deterministic matches create an ambiguous `Potential New Deal` review that cannot be approved until a master editor selects an existing investment or explicitly confirms that no existing match exists. The AI cannot override deterministic conflicts.

Potential New Deal proposals default to `Beaman Ventures`, but the entity must be confirmed before approval. Only a `master-editor` can approve a proposal as a new pipeline deal. Approval creates a `Private Investment` with status `New Lead`; it does not create capital activity, valuations, distributions, or ownership history. A proposed check size becomes the pipeline `amount` only when source-verified and explicitly confirmed. Total round size is never mapped to `amount`.

Target fund size, amount being raised, amount committed, amount remaining, historical/unfunded target difference, minimum LP commitment, total co-investment availability, and Beaman Ventures proposed check size are separate evidence fields. Fund minimums, fund targets, and total co-investment capacity never populate the proposed check field. A target-minus-historical-commitments difference is retained as historical evidence and never shown as currently remaining capital when fundraising is closed. When the email body supplies a newer explicit opportunity status than an attached deck, the email status is current and the superseded deck evidence remains attached to that status claim.

Microsoft 365 attachments are bounded to 25 MB per PDF, 10 MB per other supported file, 30 MB per message, 50 MB per run, and 20 attachments per message. Supported non-inline PDF, Office, text, CSV, and common image files are preserved in `DATA_DIR/uploads`. Only PDFs are parsed. Unsupported, excessive, or unavailable attachments remain visible on the proposal as unresolved records and are not described as parsed. Approval associates preserved files with the existing company-document vault without copying the binary.

Graph access remains read-only and requires application `Mail.Read` because message bodies and attachments are read. Do not grant `Mail.ReadWrite` or `Mail.Send`. As a separately approved Microsoft 365 hardening task, restrict the application to the intake mailbox with Exchange Online Application RBAC; this repository does not apply tenant configuration.

Email and attachment contents are untrusted model input. They are delimited as source data, cannot supply system or tool instructions, and cannot set approval, authorization, proposal status, or investment-creation fields.

## Deployment

This project includes `render.yaml` for an easy Render deployment.

1. Push the project to GitHub.
2. Create a new Render web service from the repo.
3. Add the environment variables from `.env.example`.
4. Keep `DATA_DIR=/var/data` so updates are written to the attached disk.
5. Deploy.

The app exposes a health endpoint at `/api/health`.

## Current storage model

Updates are stored in `data/investments.json` locally, or under the path in `DATA_DIR` when deployed. The included Render blueprint mounts a persistent disk at `/var/data` so updates survive restarts and deploys. For a larger team, a managed database is still the next upgrade I’d recommend.

## Phone use

Open the deployed app in your phone browser and add it to your home screen for an app-like experience.

## Deck summarization

Upload a deck as a PDF in the app, then click `Summarize deck into notes`. The app sends the PDF to OpenAI and fills the Notes field with a structured summary that can be included in the team email.

PowerPoint decks should be exported to PDF first for the most reliable results.
