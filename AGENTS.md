# AGENTS.md — mariana-outlook-mcp

## What this repo is

Custom MCP server giving Claude Code safety-first access to Microsoft accounts (Outlook Mail, Calendar, Contacts) via Microsoft Graph. **Sibling of `mariana-google-mcp`** — same architecture, same safety guarantees, Microsoft instead of Google. When changing one sibling, check whether the change belongs in both.

## Safety architecture (do not weaken)

1. **No send**: `Mail.Send` scope is never requested (`src/auth.ts` SCOPES). Never add it.
2. **No delete**: mail soft-deletes to a "To Be Deleted" folder; calendar prepends "DELETE - ". Never call Graph DELETE on mail/events.
3. **Action log**: every mutating tool calls `logAction()` (JSONL at `~/.config/mariana-outlook-mcp/action-log.jsonl`).
4. **PKCE public client**: no client secret exists anywhere. `MS_CLIENT_ID` env var is the only credential config.

## Where things live

- `src/index.ts` — server + `microsoft_auth` / `microsoft_status`
- `src/auth.ts` — PKCE OAuth flow (non-blocking consent, live token health, refresh-token rotation)
- `src/graph.ts` — minimal fetch wrapper for Graph v1.0
- `src/mail.ts` / `src/calendar.ts` / `src/contacts.ts` — tool modules
- `src/config.ts` — account registry, email-hash token paths (`~/.config/mariana-outlook-mcp/`)
- `SETUP-PROMPT.md` — paste-into-Claude-Code install doc (public; person-agnostic — keep it that way)

## Rules

- Tokens are stored under sha256(email) filenames; the config file never holds plaintext addresses.
- Build with `npm run build` (tsc, strict). `dist/` is gitignored — consumers build from source.
- Tool descriptions are user-facing contracts; keep caveats (attendee notifications, moved-message IDs, 3MB attachment cap) in them.
- README tool table must be updated in the same commit as tool changes.

## Gotchas

- Graph `$search` disallows `$orderby`; `$filter` on calendarView uses single-quote escaping (`''`).
- `createReply` produces the draft first; body/recipients are PATCHed onto it after.
- Graph small attachments (`contentBytes`) cap at ~3MB/file; bigger needs upload sessions (not implemented).
- Microsoft rotates refresh tokens — always persist the new one after refresh (`toStored` handles this).
