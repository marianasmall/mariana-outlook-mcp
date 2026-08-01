import * as http from "http";
import * as fs from "fs/promises";
import * as crypto from "crypto";
import open from "open";
import {
  ensureDirs,
  getTokenPath,
  hashEmail,
  addAccount,
  resolveAccount,
} from "./config.js";

// Mail.Send is DELIBERATELY absent: the token this server holds cannot send
// email, so "drafts only" is enforced by Microsoft's permission layer, not
// just by this code.
const SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.ReadWrite",
  "MailboxSettings.ReadWrite",
  "Calendars.ReadWrite",
  "Contacts.Read",
];

const GRAPH = "https://graph.microsoft.com/v1.0";

function authority(): string {
  return `https://login.microsoftonline.com/${process.env.MS_TENANT || "common"}`;
}

function getClientId(): string {
  const clientId = process.env.MS_CLIENT_ID;
  if (!clientId) {
    throw new Error(
      "MS_CLIENT_ID environment variable is required (Azure app registration's Application ID). No client secret is needed — this is a PKCE public client."
    );
  }
  return clientId;
}

interface StoredTokens {
  access_token: string;
  refresh_token: string;
  expires_at: number; // epoch ms
}

function b64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function escapeHtml(value: unknown): string {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

async function tokenRequest(params: Record<string, string>): Promise<any> {
  const res = await fetch(`${authority()}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params).toString(),
  });
  const data: any = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error_description || data.error || `Token request failed (${res.status})`);
  }
  return data;
}

function toStored(data: any, previousRefresh?: string): StoredTokens {
  return {
    access_token: data.access_token,
    // Microsoft rotates refresh tokens; fall back to the previous one if a
    // response omits it (e.g. some refresh responses).
    refresh_token: data.refresh_token || previousRefresh || "",
    expires_at: Date.now() + ((data.expires_in ?? 3600) - 120) * 1000,
  };
}

async function saveToken(emailHash: string, tokens: StoredTokens): Promise<void> {
  await ensureDirs();
  const tokenPath = getTokenPath(emailHash);
  // Owner-only: these are live credentials. mode applies on creation; chmod
  // covers files that already existed with looser permissions.
  await fs.writeFile(tokenPath, JSON.stringify(tokens, null, 2), { encoding: "utf-8", mode: 0o600 });
  await fs.chmod(tokenPath, 0o600);
}

async function loadToken(emailHash: string): Promise<StoredTokens | null> {
  try {
    const raw = await fs.readFile(getTokenPath(emailHash), "utf-8");
    return JSON.parse(raw) as StoredTokens;
  } catch {
    return null;
  }
}

async function fetchIdentityEmail(accessToken: string): Promise<string> {
  const res = await fetch(`${GRAPH}/me?$select=mail,userPrincipalName`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const data: any = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data?.error?.message || `Could not read identity (${res.status})`);
  }
  const email = data.mail || data.userPrincipalName;
  if (!email) throw new Error("Could not retrieve email from Microsoft account");
  return email;
}

/**
 * Start the OAuth consent flow WITHOUT blocking on completion (PKCE, no
 * client secret). Returns the consent URL immediately; the token is saved in
 * the background when the user finishes in the browser. The local callback
 * server times out after 15 minutes. `completion` resolves with the email.
 */
export async function startOAuthFlow(
  accountName: string
): Promise<{ authUrl: string; completion: Promise<string> }> {
  const clientId = getClientId();
  const codeVerifier = b64url(crypto.randomBytes(48));
  const codeChallenge = b64url(crypto.createHash("sha256").update(codeVerifier).digest());

  let resolveUrl!: (url: string) => void;
  let rejectUrl!: (err: Error) => void;
  const urlReady = new Promise<string>((res, rej) => {
    resolveUrl = res;
    rejectUrl = rej;
  });

  const completion = new Promise<string>((resolve, reject) => {
    const server = http.createServer();
    const timeout = setTimeout(() => {
      server.close();
      reject(new Error("OAuth flow timed out after 15 minutes without consent"));
    }, 15 * 60 * 1000);

    server.listen(0, async () => {
      const address = server.address();
      if (!address || typeof address === "string") {
        clearTimeout(timeout);
        const err = new Error("Failed to start local server");
        rejectUrl(err);
        reject(err);
        return;
      }
      const port = address.port;
      const redirectUri = `http://localhost:${port}`;

      const params = new URLSearchParams({
        client_id: clientId,
        response_type: "code",
        redirect_uri: redirectUri,
        response_mode: "query",
        scope: SCOPES.join(" "),
        code_challenge: codeChallenge,
        code_challenge_method: "S256",
        prompt: "select_account",
      });
      const authUrl = `${authority()}/oauth2/v2.0/authorize?${params.toString()}`;
      resolveUrl(authUrl);

      server.on("request", async (req, res) => {
        try {
          const url = new URL(req.url || "", `http://localhost:${port}`);
          const errParam = url.searchParams.get("error");
          if (errParam) {
            const desc = url.searchParams.get("error_description") || errParam;
            res.writeHead(400, { "Content-Type": "text/html" });
            res.end(`<h1>Authorization failed</h1><pre>${escapeHtml(desc)}</pre>`);
            clearTimeout(timeout);
            server.close();
            reject(new Error(desc));
            return;
          }
          const code = url.searchParams.get("code");
          if (!code) {
            // Ignore favicon and stray requests without failing the flow.
            res.writeHead(404);
            res.end();
            return;
          }

          const tokenData = await tokenRequest({
            client_id: clientId,
            grant_type: "authorization_code",
            code,
            redirect_uri: redirectUri,
            code_verifier: codeVerifier,
            scope: SCOPES.join(" "),
          });

          const email = await fetchIdentityEmail(tokenData.access_token);
          const emailHash = hashEmail(email);
          await saveToken(emailHash, toStored(tokenData));
          await addAccount(accountName, email);

          res.writeHead(200, { "Content-Type": "text/html" });
          res.end(
            `<h1>Authorization successful!</h1><p>Account "${escapeHtml(accountName)}" (${escapeHtml(email)}) connected. You can close this tab.</p>`
          );

          clearTimeout(timeout);
          server.close();
          resolve(email);
        } catch (err) {
          res.writeHead(500, { "Content-Type": "text/html" });
          res.end(`<h1>Error</h1><pre>${escapeHtml(err)}</pre>`);
          clearTimeout(timeout);
          server.close();
          reject(err instanceof Error ? err : new Error(String(err)));
        }
      });

      // Best effort: open the default browser. The URL is also returned to
      // the caller, so a missed/wrong-profile window is recoverable.
      try {
        await open(authUrl);
      } catch {
        /* URL returned to caller regardless */
      }
    });
  });

  const authUrl = await urlReady;
  return { authUrl, completion };
}

/** Blocking variant: waits for the user to finish consent. */
export async function runOAuthFlow(accountName: string): Promise<string> {
  const { completion } = await startOAuthFlow(accountName);
  return completion;
}

// Microsoft ROTATES refresh tokens on every redemption — two concurrent
// refreshes race to redeem the same token and the loser gets invalid_grant
// on a perfectly healthy account. Serialize per-account: concurrent callers
// await the in-flight refresh instead of starting their own.
const inflightRefresh = new Map<string, Promise<StoredTokens>>();

function refreshTokens(emailHash: string, stored: StoredTokens): Promise<StoredTokens> {
  const existing = inflightRefresh.get(emailHash);
  if (existing) return existing;

  const refresh = (async () => {
    const data = await tokenRequest({
      client_id: getClientId(),
      grant_type: "refresh_token",
      refresh_token: stored.refresh_token,
      scope: SCOPES.join(" "),
    });
    const next = toStored(data, stored.refresh_token);
    await saveToken(emailHash, next);
    return next;
  })();

  inflightRefresh.set(
    emailHash,
    refresh.finally(() => inflightRefresh.delete(emailHash))
  );
  return inflightRefresh.get(emailHash)!;
}

/**
 * Get a valid access token for an account, refreshing if expired.
 */
export async function getAccessToken(accountParam?: string): Promise<{
  accessToken: string;
  accountName: string;
}> {
  const resolved = await resolveAccount(accountParam);
  if (!resolved) {
    throw new Error(
      accountParam
        ? `Account "${accountParam}" not found. Run microsoft_auth to add it.`
        : "No accounts configured. Run microsoft_auth to set up your first account."
    );
  }

  let tokens = await loadToken(resolved.emailHash);
  if (!tokens || !tokens.refresh_token) {
    throw new Error(
      `No token found for account "${resolved.name}". Run microsoft_auth to re-authenticate.`
    );
  }

  if (Date.now() >= tokens.expires_at) {
    try {
      tokens = await refreshTokens(resolved.emailHash, tokens);
    } catch (err: any) {
      throw new Error(
        `Token refresh failed for account "${resolved.name}" (${err.message}). Run microsoft_auth to re-authenticate.`
      );
    }
  }

  return { accessToken: tokens.access_token, accountName: resolved.name };
}

/**
 * Check token health. An unexpired access token is reported valid as-is;
 * an expired one triggers a LIVE refresh against Microsoft, so a
 * revoked/expired refresh token reports "needs_reauth" instead of falsely
 * claiming "valid". Gating on expiry (rather than always refreshing) keeps
 * rotation churn and refresh collisions to a minimum.
 */
export async function checkTokenHealth(
  emailHash: string
): Promise<"valid" | "needs_reauth" | "missing"> {
  const tokens = await loadToken(emailHash);
  if (!tokens || !tokens.refresh_token) return "missing";
  if (Date.now() < tokens.expires_at) return "valid";
  try {
    await refreshTokens(emailHash, tokens);
    return "valid";
  } catch {
    return "needs_reauth";
  }
}
