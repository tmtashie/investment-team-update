# Investment Team Updates

A phone-friendly web app for logging investment updates and emailing your team a polished summary.

## What it does

- Lets team members sign in with email plus a shared workspace password
- Saves each investment update in `data/investments.json`
- Sends a formatted summary email through Resend
- Can summarize uploaded PDF decks into investment notes using OpenAI
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
- `RESEND_API_KEY`: API key for Resend
- `OPENAI_API_KEY`: API key for deck summarization
- `OPENAI_MODEL`: optional override for the summarization model

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
