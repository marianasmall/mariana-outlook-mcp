# Setup: Outlook Mail + Calendar + Contacts access for Claude Code

**Instructions for Claude Code.** A user has pasted this document (or a link to it) into you so you can set up Microsoft account access for them. Follow it step by step. Do everything you can yourself; where a step requires clicking in a browser, give the user numbered instructions, then STOP and wait for them to say "done" before continuing. Do not skip ahead.

## What you are building

You will install `mariana-outlook-mcp`, an open-source MCP server (MIT license) that gives Claude Code direct access to Outlook Mail, Calendar, and Contacts across multiple Microsoft accounts — both M365 work/school accounts and personal accounts (Outlook.com, Hotmail). It is deliberately safety-first:

- **It cannot send email — enforced by Microsoft.** The server never requests the send permission, so its tokens are physically incapable of sending. The user drafts; the user sends from Outlook.
- **It cannot delete anything.** Mail moves to a "To Be Deleted" folder; calendar deletions prepend "DELETE - " to the event title. The user confirms in the Outlook UI.
- **Every write is logged** to an append-only local action log.
- **Multi-account:** each Microsoft account connects under a named alias and tools take an `account` parameter.
- **No client secret exists.** The only credential is a public Application ID (PKCE desktop flow).

Repo: https://github.com/marianasmall/mariana-outlook-mcp
Sibling for Google accounts: https://github.com/marianasmall/mariana-google-mcp

## Phase 1 — Install and connect

### Step 0 — Preflight (you do this)

Check that `node` (v18+), `npm`, and `git` are installed (on macOS: `brew install node git`). Ask the user which directory they keep code projects in; default to `~/Projects/`.

Also ask now: **which Microsoft accounts do they want to connect, and what short alias should each get?** (e.g. `work`, `consulting`, `personal`.)

### Step 1 — Azure app registration (user does this in a browser; you guide)

Walk the user through with numbered instructions, one checkpoint at a time. Microsoft's portal UI changes; adapt to what they actually see.

1. Go to https://portal.azure.com and sign in with a Microsoft work account (an M365 account creates this in its own tenant; the other accounts connect later regardless).
2. In the top search bar, search **"App registrations"** → open it → **New registration**.
3. Name: "Claude Code".
4. Supported account types: choose the option that includes BOTH organizational directories AND personal Microsoft accounts ("Accounts in any organizational directory ... and personal Microsoft accounts"). This is what lets work and personal accounts connect to the same app.
5. Redirect URI: platform dropdown **"Public client/native (mobile & desktop)"**, value: `http://localhost`
6. Click Register.
7. On the Overview page, copy the **Application (client) ID** (a GUID) and paste it back to you.
8. Left menu → **Authentication** → scroll to "Advanced settings" → set **"Allow public client flows"** to **Yes** → Save.

No client secret is created anywhere in this flow — that is correct, not a missed step.

### Step 2 — Install the server (you do this)

```bash
cd ~/Projects
git clone https://github.com/marianasmall/mariana-outlook-mcp.git
cd mariana-outlook-mcp
npm install
npm run build
```

Confirm `dist/index.js` exists before proceeding.

### Step 3 — Register the MCP server (you do this)

```bash
claude mcp add outlook --scope user \
  -e MS_CLIENT_ID="<application-client-id>" \
  -- node /FULL/PATH/TO/mariana-outlook-mcp/dist/index.js
```

Use the absolute path to the clone's `dist/index.js`. Fallback if `claude mcp add` is unavailable: add the equivalent entry to `~/.claude.json` under top-level `"mcpServers"` (type stdio, command node, args [path], env {MS_CLIENT_ID}).

Then tell the user to fully restart Claude Code (quit and relaunch — MCP servers only load on startup) and paste this document in again, saying "continue from Step 4".

### Step 4 — Connect each account (after restart)

For each account from Step 0:

1. Run the `microsoft_auth` tool with that account's alias. The first account connected becomes the default.
2. A browser window opens. The user picks/enters the matching Microsoft account and signs in.
3. A permissions consent screen lists the requested access (read/write mail, calendars, contacts — note: no send permission). Accept.
4. For an M365 WORK account, if the org blocks user consent, an admin must approve — if the user is their own admin: Microsoft Entra admin center → Enterprise applications → Consent and permissions (or just grant admin consent on the app registration's API permissions page).

Tokens are stored locally at `~/.config/mariana-outlook-mcp/` and refresh automatically with use.

### Step 5 — Verify (you do this)

1. Run `microsoft_status` — every account should show `tokenStatus: valid`.
2. Run an `outlook_search` for something recent on each account and show a couple of subject lines as proof of life.
3. Briefly show the user what they can now ask for: mail search/triage across accounts, drafted replies (they send), folder/category cleanup, inbox rules, calendar management, contact lookups.

## Phase 2 — Optional: unify a multi-account inbox

Same playbook as the Google sibling — if the user runs 3+ accounts, offer: (1) **Claude as the unified triage layer** — a "sweep my inboxes" routine across all accounts (works across BOTH this server and the Google sibling if installed — one sweep covers Gmail and Outlook together); (2) **consolidation** via Outlook rules/forwarding; (3) **shared category taxonomy** via `outlook_create_category` + `outlook_create_rule`. Recommend starting with (1).

## Troubleshooting

- **AADSTS7000218 / public client error:** "Allow public client flows" wasn't set to Yes (Step 1.8).
- **AADSTS50194 / wrong audience:** the registration was created single-tenant; change Supported account types to include personal accounts (Manifest: `signInAudience` = `AzureADandPersonalMicrosoftAccount`).
- **"Need admin approval" on a work account:** org requires admin consent (Step 4.4).
- **Account disconnects after months of no use:** personal-account refresh tokens expire ~90 days without use; re-run `microsoft_auth`.
- **Server not appearing after restart:** check the absolute path and `claude mcp list`.

## Security rules (state these to the user at the end)

- The Application ID is public by design — there is no secret to leak in this setup.
- Access tokens live only on this machine (`~/.config/mariana-outlook-mcp/`); another computer needs its own `microsoft_auth`.
- Revoking access: https://account.live.com/consent/Manage (personal) or https://myaccount.microsoft.com → App permissions (work), or delete the local tokens directory.
