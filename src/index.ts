#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { registerMailTools } from "./mail.js";
import { registerCalendarTools } from "./calendar.js";
import { registerContactsTools } from "./contacts.js";
import { startOAuthFlow, checkTokenHealth } from "./auth.js";
import { loadConfig, ensureDirs } from "./config.js";
import { logAction } from "./logger.js";

const server = new McpServer({
  name: "mariana-outlook-mcp",
  version: "1.0.0",
});

// --- microsoft_auth ---
server.tool(
  "microsoft_auth",
  "Authenticate a Microsoft account (M365 work/school or personal e.g. Hotmail). Returns the consent URL IMMEDIATELY (does not block): give it to the user, they approve in the browser as the right Microsoft account, and the token saves in the background. Confirm afterwards with microsoft_status. A browser tab also opens automatically as a convenience, but the URL in the response is authoritative if it opens with the wrong profile.",
  {
    account_name: z
      .string()
      .optional()
      .default("primary")
      .describe("Friendly name for this account (e.g. 'consulting', 'personal')"),
  },
  async ({ account_name }) => {
    try {
      const { authUrl, completion } = await startOAuthFlow(account_name);

      // Log completion (or failure) in the background; the tool returns now.
      completion
        .then((email) =>
          logAction({
            timestamp: new Date().toISOString(),
            tool: "microsoft_auth",
            account: account_name,
            summary: `Account authenticated: ${email}`,
          })
        )
        .catch((err) =>
          logAction({
            timestamp: new Date().toISOString(),
            tool: "microsoft_auth",
            account: account_name,
            summary: `Auth flow did not complete: ${err.message}`,
          })
        );

      return {
        content: [
          {
            type: "text" as const,
            text: `Consent flow started for account "${account_name}". Have the user open this URL and approve as the correct Microsoft account (a browser tab may have opened already):\n\n${authUrl}\n\nThe token saves automatically on completion (15-minute window). Verify with microsoft_status.`,
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [{ type: "text" as const, text: `Authentication failed: ${err.message}` }],
        isError: true,
      };
    }
  }
);

// --- microsoft_status ---
server.tool(
  "microsoft_status",
  "Check connection health for all configured Microsoft accounts with a LIVE credential check against Microsoft (not just local file state). tokenStatus 'needs_reauth' means run microsoft_auth for that account.",
  {},
  async () => {
    const config = await loadConfig();
    const accounts = Object.entries(config.accounts);

    if (accounts.length === 0) {
      return {
        content: [
          {
            type: "text" as const,
            text: "No accounts configured. Use microsoft_auth to set up your first account.",
          },
        ],
      };
    }

    const statuses = await Promise.all(
      accounts.map(async ([name, entry]) => {
        const tokenStatus = await checkTokenHealth(entry.emailHash);
        return {
          account: name,
          label: entry.label,
          tokenStatus,
          isDefault: name === config.defaultAccount,
          services: ["mail", "calendar", "contacts"],
        };
      })
    );

    return {
      content: [{ type: "text" as const, text: JSON.stringify(statuses, null, 2) }],
    };
  }
);

// Register all tool modules
registerMailTools(server);
registerCalendarTools(server);
registerContactsTools(server);

// Start server
async function main() {
  await ensureDirs();
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error("Fatal:", err);
  process.exit(1);
});
