# Planit

A desktop calendar planner with Doodle-style group scheduling and Microsoft 365 (Outlook) integration. Built with Electron.

## Features

- **Calendar** — month and week views of your Outlook calendar. Click any day/slot to create an event; invite people by email and real Outlook invitations go out.
- **Polls** — Doodle-style scheduling polls. Propose time slots, track everyone's yes/maybe/no votes, email the options to participants from your Outlook account, then "Book best slot" to turn the winner into a real calendar event with invites.
- **Group** — side-by-side free/busy grid for a list of people, powered by Microsoft 365 free/busy lookup, so you can spot a slot that works for everyone.
- **Mail** — your Outlook inbox: read, reply, and compose.
- **Demo mode** — everything works with sample data before you connect an account, so you can explore immediately.

## Run it

Requires [Node.js](https://nodejs.org) 18+.

```bash
npm install
npm start
```

To build an installer for your platform (`.exe` on Windows, `.dmg` on Mac):

```bash
npm run dist
```

The installer lands in the `dist/` folder.

## Connect your Microsoft account (~5 minutes, free)

Planit talks directly to Microsoft Graph — no third-party server ever sees your data. Microsoft requires every app to have an **app registration**, which is free and takes a few minutes:

1. Go to <https://entra.microsoft.com> (or portal.azure.com) and sign in with your Microsoft account.
2. Navigate to **Identity → Applications → App registrations → New registration**.
3. Name: `Planit` (anything works).
4. **Supported account types**: choose *"Accounts in any organizational directory and personal Microsoft accounts"* (this allows both work and personal accounts, including hotmail/outlook.com).
5. **Redirect URI**: choose platform **"Mobile and desktop applications"** and enter:
   ```
   http://localhost
   ```
6. Click **Register**, then copy the **Application (client) ID** from the overview page.
7. Still in the registration, open **Authentication** and make sure **"Allow public client flows"** is set to **Yes**.
8. In Planit: **Settings → Microsoft account**, paste the client ID, keep tenant `common`, click **Save**, then **Sign in to Microsoft**. Your browser opens; sign in and approve the permissions (read/write calendar, read/send mail).

That's it — your real calendar, inbox, and group availability replace the demo data.

### Permissions Planit asks for

`User.Read`, `Calendars.ReadWrite`, `Calendars.Read.Shared`, `Mail.Read`, `Mail.Send` — all delegated (only while you're signed in), nothing needs admin consent.

## Group calendars — how sharing works

The **Group** tab has three modes:

- **Group calendars**: anyone can create a named calendar (e.g. "Trip crew") and invite members with **write access** — each member gets an Outlook invite, and once accepted the calendar shows up in their Planit calendar picker (top-left of the Calendar view). Everyone plans their own events directly in it, so the group shares one living calendar. Events there can be Teams meetings too.

- **Shared events** (works for personal *and* work accounts): add your group members' emails, click **"Share my calendar with the group"** (this sends each member an Outlook calendar-sharing invite), and have each member do the same from their Planit or from Outlook (Calendar → Share). Once invites are accepted, everyone's planned events appear side by side, color-coded per person. This is the group calendar.
- **Free/busy** (work/school accounts only): a busy/free grid via the Graph `getSchedule` API — works instantly for people in the same Microsoft 365 organization with no sharing step, but Microsoft doesn't offer it between personal accounts.

## Microsoft Teams meetings

Any event — from the New event form or a poll's "Book best slot" — can be created as an online meeting: tick **"Microsoft Teams meeting"** and the event gets a join link (🎥 marker on the calendar; "Join Teams meeting" button when you open it). Work/school accounts get a Teams link; for personal accounts Planit automatically falls back to Microsoft's consumer meeting-link provider, since Teams-for-business links are a work-account feature.

If the in-app share button fails for a particular recipient (Microsoft's sharing API has gaps for some personal-account combinations), sharing from Outlook's own UI (Calendar → Share → enter their email → "Can view all details") achieves exactly the same thing — Planit picks up any accepted shared calendar automatically.

## Releasing updates to your colleagues

Planit auto-updates from **GitHub Releases**: every installed copy checks on launch (and every 4 hours), downloads the new version in the background, and installs it when the user restarts. Your release workflow:

**One-time setup**
1. Create a repo on github.com (e.g. `yourname/planit`) and push this project to it.
2. In `package.json`, set `build.publish.owner` to your GitHub username (and `repo` if you named it differently).
3. Create a token at github.com → Settings → Developer settings → Personal access tokens (classic, `repo` scope). In PowerShell: `setx GH_TOKEN "ghp_yourtoken"` (then open a new terminal).

**Every release**
1. Make your changes and test with `npm start`.
2. Bump `"version"` in `package.json` (e.g. `1.0.0` → `1.1.0`) — the updater only picks up higher versions.
3. Run `npm run release` — this builds the Windows installer and uploads it (plus the update manifest `latest.yml`) to a GitHub release automatically.
4. Publish the draft release on GitHub if it's left as draft. Done — colleagues get it on their next launch, no reinstall needed.

New colleagues still install once from the `Planit Setup x.y.z.exe` in that same GitHub release; from then on they're on the update train. Settings → "Check for updates now" lets anyone pull an update immediately.

## Where data lives

- Your Microsoft data stays between the app and Microsoft Graph; sign-in tokens are cached locally in Electron's user-data folder.
- Polls, settings, and offline demo events are stored as local JSON on your machine only.

## Project layout

```
src/main.js      Electron main process, IPC, local JSON stores
src/auth.js      Microsoft sign-in (MSAL, auth-code + PKCE via your browser)
src/graph.js     Microsoft Graph REST wrappers (calendar, schedule, mail)
src/preload.js   Secure bridge between UI and main process
renderer/        The app UI (no framework, plain HTML/CSS/JS)
```
