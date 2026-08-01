# Decisions — mariana-outlook-mcp

Append-only log: what we decided and why.

## 2026-07-31 — Initial build

- **PKCE public client, no client secret.** Microsoft desktop apps don't need one; removes an entire secret-management surface vs. the Google sibling. Only `MS_CLIENT_ID` (public by design) is configured.
- **`Mail.Send` scope never requested.** The drafts-only guarantee is enforced at Microsoft's permission layer — a compromised or buggy server process still cannot send. Stronger than the Google sibling, where gmail.compose could technically send.
- **Categories as the Gmail-labels analog; a folder for soft-delete.** Outlook has both concepts; categories behave like labels (many per message), while "To Be Deleted" as a folder matches native Outlook idiom and keeps messages out of the inbox view.
- **Reply drafts via `createReply` + PATCH.** Graph threads replies from a message ID, not a conversation ID — so the tool takes `reply_to_message_id` (differs from the sibling's `thread_id`; documented in the tool description).
- **Attendees not attached by default on `calendar_create`.** Graph auto-sends invitations whenever attendees exist on an event (no `sendUpdates: none` equivalent). To honor no-surprise-emails, attendee names go into the description unless `send_invites=true`.
- **Availability derived from `calendarView` showAs, not `getSchedule`.** Avoids needing the account's plaintext SMTP address (we store only hashes) and an extra permission.
- **Attachment cap 3MB/file.** Graph's simple `contentBytes` limit; upload sessions deferred until a real need appears.
- **Repo named `mariana-outlook-mcp` (not `app-` prefixed).** Deliberate REPO-STRUCTURE.md deviation so the pair groups with its established sibling `mariana-google-mcp` when browsing; flagged to Mariana at build time.
- **Public + MIT from day one**, with SETUP-PROMPT.md — the "paste one line into Claude Code" distribution pattern proven on the sibling earlier today.
- **Pre-publish code review (independent agent) found and we fixed 5 issues:** (1) batch tools now log partial successes so a mid-loop failure can't leave unlogged mutations; (2) per-account refresh mutex — Microsoft rotates refresh tokens, concurrent refreshes raced to invalid_grant; also `checkTokenHealth` now gates its live refresh on expiry; (3) `calendar_search` filters client-side (Graph's `$filter` on calendarView subject 500s when any event in range has an empty subject — documented bug); (4) `calendar_availability` computes default end-of-day in UTC (date-only inputs were silently shrinking the window by the host's UTC offset); (5) HTML-escaping on OAuth callback responses.
- **Security hardening (from commit-time review):** token files written 0600, tokens dir 0700 — ported back to the Google sibling the same evening.
