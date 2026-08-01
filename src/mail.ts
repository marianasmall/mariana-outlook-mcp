import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { promises as fs } from "fs";
import * as path from "path";
import { getAccessToken } from "./auth.js";
import { graphFetch } from "./graph.js";
import { logAction } from "./logger.js";

// Graph's simple fileAttachment (contentBytes) path is capped at ~3MB per
// file; larger files need an upload session, which this server doesn't do yet.
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

const MESSAGE_FIELDS =
  "id,conversationId,subject,from,toRecipients,receivedDateTime,bodyPreview,categories,isRead";

function recipients(list?: string): Array<{ emailAddress: { address: string } }> | undefined {
  if (!list) return undefined;
  return list
    .split(",")
    .map((a) => a.trim())
    .filter(Boolean)
    .map((address) => ({ emailAddress: { address } }));
}

function formatAddress(r: any): string {
  const name = r?.emailAddress?.name || "";
  const addr = r?.emailAddress?.address || "";
  return name && name !== addr ? `${name} <${addr}>` : addr;
}

function formatSummary(msg: any) {
  return {
    id: msg.id,
    conversationId: msg.conversationId,
    from: formatAddress(msg.from),
    to: (msg.toRecipients || []).map(formatAddress).join(", "),
    subject: msg.subject || "",
    date: msg.receivedDateTime || "",
    snippet: msg.bodyPreview || "",
    categories: msg.categories || [],
    isRead: msg.isRead,
  };
}

async function findOrCreateFolder(accessToken: string, displayName: string): Promise<any> {
  const escaped = displayName.replace(/'/g, "''");
  const found = await graphFetch(
    accessToken,
    `/me/mailFolders?$filter=displayName eq '${encodeURIComponent(escaped)}'`
  );
  if (found.value && found.value.length > 0) return found.value[0];
  return graphFetch(accessToken, "/me/mailFolders", {
    method: "POST",
    body: { displayName },
  });
}

export function registerMailTools(server: McpServer): void {
  // --- outlook_search ---
  server.tool(
    "outlook_search",
    "Search Outlook messages. Uses KQL syntax (e.g. 'from:name subject:topic', or plain keywords).",
    {
      query: z.string().describe("Search query (KQL: from:, subject:, to:, or keywords)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
      max_results: z.coerce.number().optional().default(10).describe("Max results (default 10)"),
    },
    async ({ query, account, max_results }) => {
      const { accessToken } = await getAccessToken(account);
      const search = encodeURIComponent(`"${query.replace(/"/g, '\\"')}"`);
      const res = await graphFetch(
        accessToken,
        `/me/messages?$search=${search}&$top=${max_results}&$select=${MESSAGE_FIELDS}`
      );
      const summaries = (res.value || []).map(formatSummary);
      return { content: [{ type: "text" as const, text: JSON.stringify(summaries, null, 2) }] };
    }
  );

  // --- outlook_read ---
  server.tool(
    "outlook_read",
    "Read a specific Outlook message by ID. Returns full message content as plain text.",
    {
      message_id: z.string().describe("Outlook message ID"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ message_id, account }) => {
      const { accessToken } = await getAccessToken(account);
      const msg = await graphFetch(
        accessToken,
        `/me/messages/${message_id}?$select=${MESSAGE_FIELDS},ccRecipients,body,hasAttachments`,
        { headers: { Prefer: 'outlook.body-content-type="text"' } }
      );

      let attachments: string[] = [];
      if (msg.hasAttachments) {
        const att = await graphFetch(
          accessToken,
          `/me/messages/${message_id}/attachments?$select=name,size,contentType`
        );
        attachments = (att.value || []).map((a: any) => a.name);
      }

      const message = {
        id: msg.id,
        conversationId: msg.conversationId,
        from: formatAddress(msg.from),
        to: (msg.toRecipients || []).map(formatAddress).join(", "),
        cc: (msg.ccRecipients || []).map(formatAddress).join(", "),
        subject: msg.subject || "",
        date: msg.receivedDateTime || "",
        body: msg.body?.content || "",
        attachments,
        categories: msg.categories || [],
      };

      return { content: [{ type: "text" as const, text: JSON.stringify(message, null, 2) }] };
    }
  );

  // --- outlook_draft ---
  server.tool(
    "outlook_draft",
    "Create a draft email. Does NOT send — the draft appears in Outlook's Drafts folder for manual review and sending. (The server's Microsoft permissions exclude sending entirely.)",
    {
      to: z.string().describe("Recipient email address(es), comma-separated"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body (plain text)"),
      html: z
        .string()
        .optional()
        .describe(
          "Optional HTML body (tables, bold, links). PREFER this whenever the content has structure — figures, comparisons, per-person breakdowns — a formatted draft reads better than plain text. When provided it is used instead of the plain-text body."
        ),
      cc: z.string().optional().describe("CC recipients (comma-separated)"),
      bcc: z.string().optional().describe("BCC recipients (comma-separated)"),
      attachments: z
        .array(z.string())
        .optional()
        .describe("Absolute file paths to attach (3MB max per file)"),
      reply_to_message_id: z
        .string()
        .optional()
        .describe(
          "Message ID to draft a REPLY to (keeps the conversation thread and original recipients; get the ID from outlook_search/outlook_read). Omit for a new conversation."
        ),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ to, subject, body, html, cc, bcc, attachments, reply_to_message_id, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);

      const messageBody = {
        contentType: html ? "HTML" : "Text",
        content: html ?? body,
      };

      let draft: any;
      if (reply_to_message_id) {
        draft = await graphFetch(accessToken, `/me/messages/${reply_to_message_id}/createReply`, {
          method: "POST",
          body: {},
        });
        const patch: any = { body: messageBody };
        if (to) patch.toRecipients = recipients(to);
        if (cc) patch.ccRecipients = recipients(cc);
        if (bcc) patch.bccRecipients = recipients(bcc);
        draft = await graphFetch(accessToken, `/me/messages/${draft.id}`, {
          method: "PATCH",
          body: patch,
        });
      } else {
        draft = await graphFetch(accessToken, "/me/messages", {
          method: "POST",
          body: {
            subject,
            body: messageBody,
            toRecipients: recipients(to),
            ...(cc ? { ccRecipients: recipients(cc) } : {}),
            ...(bcc ? { bccRecipients: recipients(bcc) } : {}),
          },
        });
      }

      let attached = 0;
      const attachFailures: string[] = [];
      if (attachments && attachments.length > 0) {
        for (const filePath of attachments) {
          try {
            const data = await fs.readFile(filePath);
            if (data.length > MAX_ATTACHMENT_BYTES) {
              throw new Error(
                `exceeds ${MAX_ATTACHMENT_BYTES / 1024 / 1024}MB (Graph simple-attachment limit)`
              );
            }
            await graphFetch(accessToken, `/me/messages/${draft.id}/attachments`, {
              method: "POST",
              body: {
                "@odata.type": "#microsoft.graph.fileAttachment",
                name: path.basename(filePath),
                contentBytes: data.toString("base64"),
              },
            });
            attached++;
          } catch (err: any) {
            attachFailures.push(`${path.basename(filePath)}: ${err.message}`);
          }
        }
      }

      // The draft mutation already happened — log it even if attachments failed.
      await logAction({
        timestamp: new Date().toISOString(),
        tool: "outlook_draft",
        account: accountName,
        summary: `Draft created: '${subject}' to ${to}${attached ? ` (+${attached} attachment${attached > 1 ? "s" : ""})` : ""}${attachFailures.length ? ` (${attachFailures.length} attachment(s) FAILED)` : ""}${reply_to_message_id ? " (reply in thread)" : ""}`,
        ids: [draft.id || "unknown"],
      });

      const failureNote = attachFailures.length
        ? `\nWARNING — attachment(s) not added: ${attachFailures.join("; ")}`
        : "";
      return {
        content: [
          {
            type: "text" as const,
            text: `Draft created successfully (ID: ${draft.id}). Review and send from Outlook.${failureNote}`,
          },
        ],
        ...(attachFailures.length ? { isError: true } : {}),
      };
    }
  );

  // --- outlook_list_folders ---
  server.tool(
    "outlook_list_folders",
    "List Outlook mail folders for an account.",
    {
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ account }) => {
      const { accessToken } = await getAccessToken(account);
      const res = await graphFetch(
        accessToken,
        "/me/mailFolders?$top=100&$select=id,displayName,totalItemCount,unreadItemCount"
      );
      const folders = (res.value || []).map((f: any) => ({
        id: f.id,
        name: f.displayName,
        totalItems: f.totalItemCount,
        unreadItems: f.unreadItemCount,
      }));
      return { content: [{ type: "text" as const, text: JSON.stringify(folders, null, 2) }] };
    }
  );

  // --- outlook_list_categories ---
  server.tool(
    "outlook_list_categories",
    "List Outlook categories (the Gmail-labels equivalent) for an account.",
    {
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ account }) => {
      const { accessToken } = await getAccessToken(account);
      const res = await graphFetch(accessToken, "/me/outlook/masterCategories?$top=100");
      const categories = (res.value || []).map((c: any) => ({
        id: c.id,
        name: c.displayName,
        color: c.color,
      }));
      return { content: [{ type: "text" as const, text: JSON.stringify(categories, null, 2) }] };
    }
  );

  // --- outlook_create_category ---
  server.tool(
    "outlook_create_category",
    "Create a new Outlook category (the Gmail-labels equivalent).",
    {
      name: z.string().describe("Category name"),
      color_preset: z.coerce
        .number()
        .optional()
        .default(0)
        .describe("Color preset 0-24 (default 0)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ name, color_preset, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);
      const res = await graphFetch(accessToken, "/me/outlook/masterCategories", {
        method: "POST",
        body: { displayName: name, color: `preset${Math.min(Math.max(color_preset, 0), 24)}` },
      });

      await logAction({
        timestamp: new Date().toISOString(),
        tool: "outlook_create_category",
        account: accountName,
        summary: `Created category: "${name}" (ID: ${res.id})`,
      });

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ id: res.id, name: res.displayName, color: res.color }, null, 2),
          },
        ],
      };
    }
  );

  // --- outlook_apply_category ---
  server.tool(
    "outlook_apply_category",
    "Apply a category to one or more Outlook messages (by category NAME, not ID).",
    {
      message_ids: z.array(z.string()).describe("Array of Outlook message IDs"),
      category: z.string().describe("Category name (from outlook_list_categories)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ message_ids, category, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);

      const succeeded: string[] = [];
      const failed: string[] = [];
      for (const msgId of message_ids) {
        try {
          const current = await graphFetch(
            accessToken,
            `/me/messages/${msgId}?$select=categories`
          );
          const cats: string[] = current.categories || [];
          if (!cats.includes(category)) {
            await graphFetch(accessToken, `/me/messages/${msgId}`, {
              method: "PATCH",
              body: { categories: [...cats, category] },
            });
          }
          succeeded.push(msgId);
        } catch (err: any) {
          failed.push(`${msgId}: ${err.message}`);
        }
      }

      // Log whatever actually mutated, even when later items failed.
      if (succeeded.length > 0) {
        await logAction({
          timestamp: new Date().toISOString(),
          tool: "outlook_apply_category",
          account: accountName,
          summary: `Applied category "${category}" to ${succeeded.length} message(s)${failed.length ? ` (${failed.length} failed)` : ""}`,
          ids: succeeded,
        });
      }

      const failNote = failed.length ? `\nFAILED for: ${failed.join("; ")}` : "";
      return {
        content: [{ type: "text" as const, text: `Applied category to ${succeeded.length} message(s).${failNote}` }],
        ...(failed.length ? { isError: true } : {}),
      };
    }
  );

  // --- outlook_remove_category ---
  server.tool(
    "outlook_remove_category",
    "Remove a category from one or more Outlook messages (by category NAME).",
    {
      message_ids: z.array(z.string()).describe("Array of Outlook message IDs"),
      category: z.string().describe("Category name to remove"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ message_ids, category, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);

      const succeeded: string[] = [];
      const failed: string[] = [];
      for (const msgId of message_ids) {
        try {
          const current = await graphFetch(
            accessToken,
            `/me/messages/${msgId}?$select=categories`
          );
          const cats: string[] = current.categories || [];
          await graphFetch(accessToken, `/me/messages/${msgId}`, {
            method: "PATCH",
            body: { categories: cats.filter((c) => c !== category) },
          });
          succeeded.push(msgId);
        } catch (err: any) {
          failed.push(`${msgId}: ${err.message}`);
        }
      }

      if (succeeded.length > 0) {
        await logAction({
          timestamp: new Date().toISOString(),
          tool: "outlook_remove_category",
          account: accountName,
          summary: `Removed category "${category}" from ${succeeded.length} message(s)${failed.length ? ` (${failed.length} failed)` : ""}`,
          ids: succeeded,
        });
      }

      const failNote = failed.length ? `\nFAILED for: ${failed.join("; ")}` : "";
      return {
        content: [{ type: "text" as const, text: `Removed category from ${succeeded.length} message(s).${failNote}` }],
        ...(failed.length ? { isError: true } : {}),
      };
    }
  );

  // --- outlook_create_rule ---
  server.tool(
    "outlook_create_rule",
    "Create an Outlook inbox rule (the Gmail-filter equivalent). Matches incoming messages by criteria and applies actions automatically. At least one criterion and one action required.",
    {
      name: z.string().describe("Rule display name"),
      from_contains: z.array(z.string()).optional().describe("Match sender addresses/domains containing any of these"),
      subject_contains: z.array(z.string()).optional().describe("Match subject containing any of these"),
      body_contains: z.array(z.string()).optional().describe("Match body containing any of these"),
      move_to_folder_id: z.string().optional().describe("Folder ID to move matches to (from outlook_list_folders)"),
      assign_category: z.string().optional().describe("Category name to assign to matches"),
      mark_read: z.boolean().optional().describe("Mark matching messages as read"),
      forward_to: z.string().optional().describe("Email address to forward matches to (NOTE: forwarding sends mail automatically — use deliberately)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ name, from_contains, subject_contains, body_contains, move_to_folder_id, assign_category, mark_read, forward_to, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);

      const conditions: Record<string, unknown> = {};
      if (from_contains?.length) conditions.senderContains = from_contains;
      if (subject_contains?.length) conditions.subjectContains = subject_contains;
      if (body_contains?.length) conditions.bodyContains = body_contains;
      if (Object.keys(conditions).length === 0) {
        return {
          content: [{ type: "text" as const, text: "Error: At least one criterion (from_contains, subject_contains, body_contains) is required." }],
          isError: true,
        };
      }

      const actions: Record<string, unknown> = { stopProcessingRules: false };
      if (move_to_folder_id) actions.moveToFolder = move_to_folder_id;
      if (assign_category) actions.assignCategories = [assign_category];
      if (mark_read) actions.markAsRead = true;
      if (forward_to) actions.forwardTo = [{ emailAddress: { address: forward_to } }];
      if (Object.keys(actions).length === 1) {
        return {
          content: [{ type: "text" as const, text: "Error: At least one action (move_to_folder_id, assign_category, mark_read, forward_to) is required." }],
          isError: true,
        };
      }

      const existing = await graphFetch(accessToken, "/me/mailFolders/inbox/messageRules?$top=100");
      const sequence = (existing.value?.length || 0) + 1;

      const res = await graphFetch(accessToken, "/me/mailFolders/inbox/messageRules", {
        method: "POST",
        body: { displayName: name, sequence, isEnabled: true, conditions, actions },
      });

      await logAction({
        timestamp: new Date().toISOString(),
        tool: "outlook_create_rule",
        account: accountName,
        summary: `Created inbox rule: "${name}" (ID: ${res.id})`,
      });

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ id: res.id, name: res.displayName, sequence: res.sequence }, null, 2),
          },
        ],
      };
    }
  );

  // --- outlook_move_to_delete ---
  server.tool(
    "outlook_move_to_delete",
    "Soft-delete: move messages to a 'To Be Deleted' folder. Does NOT delete or trash messages. NOTE: moved messages get NEW message IDs.",
    {
      message_ids: z.array(z.string()).describe("Array of Outlook message IDs to move"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ message_ids, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);
      const folder = await findOrCreateFolder(accessToken, "To Be Deleted");

      const succeeded: string[] = [];
      const failed: string[] = [];
      for (const msgId of message_ids) {
        try {
          await graphFetch(accessToken, `/me/messages/${msgId}/move`, {
            method: "POST",
            body: { destinationId: folder.id },
          });
          succeeded.push(msgId);
        } catch (err: any) {
          failed.push(`${msgId}: ${err.message}`);
        }
      }

      // Log the messages that really moved, even when later items failed —
      // the audit trail must reflect actual mutations.
      if (succeeded.length > 0) {
        await logAction({
          timestamp: new Date().toISOString(),
          tool: "outlook_move_to_delete",
          account: accountName,
          summary: `Moved ${succeeded.length} message(s) to "To Be Deleted"${failed.length ? ` (${failed.length} failed)` : ""}`,
          ids: succeeded,
        });
      }

      const failNote = failed.length ? `\nFAILED for: ${failed.join("; ")}` : "";
      return {
        content: [
          {
            type: "text" as const,
            text: `Moved ${succeeded.length} message(s) to "To Be Deleted" folder. Review in Outlook before permanently deleting.${failNote}`,
          },
        ],
        ...(failed.length ? { isError: true } : {}),
      };
    }
  );
}
