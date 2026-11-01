import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { Composio } from "@composio/core";
import { execFile } from "node:child_process";
import { renderCustom, markHtml, esc, clip } from "./widgetKit.ts";

const server = new McpServer({ name: "microsoft-teams", version: "1.0.0" });
const ACCENT = "#5059c9"; // Microsoft Teams purple from Fluent design system
const COMPOSIO_API_HOST = "backend.composio.dev";

// Helper to enforce an explicit timeout on promises
async function withTimeout<T>(promise: Promise<T>, ms = 8000): Promise<T> {
  let timer: NodeJS.Timeout;
  const timeoutPromise = new Promise<never>((_, reject) => {
    timer = setTimeout(() => {
      reject(new Error(`Operation timed out after ${ms}ms`));
    }, ms);
  });
  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    clearTimeout(timer!);
  }
}

// Open redirect URL in the default browser using macOS `open`
function openBrowser(url: string): void {
  try {
    execFile("open", [url], (err) => {
      if (err) console.error("Failed to open browser with macOS open:", err);
    });
  } catch (err) {
    console.error("Failed to execute open command:", err);
  }
}

function isAuthError(errText: string): boolean {
  const lower = errText.toLowerCase();
  return (
    lower.includes("401") ||
    lower.includes("unauthorized") ||
    lower.includes("invalid token") ||
    lower.includes("invalid_token") ||
    lower.includes("revoked") ||
    lower.includes("expired") ||
    lower.includes("re-authenticate") ||
    lower.includes("reauthenticate")
  );
}

function isPermissionError(errText: string): boolean {
  const lower = errText.toLowerCase();
  return (
    lower.includes("403") ||
    lower.includes("forbidden") ||
    lower.includes("admin approval") ||
    lower.includes("permission") ||
    lower.includes("accessdenied") ||
    lower.includes("authorization_requestdenied")
  );
}

async function handleReauth(session: any): Promise<void> {
  try {
    const authRes = await withTimeout(session.authorize("microsoft_teams"), 8000);
    const redirectUrl = authRes?.redirectUrl || authRes?.url || authRes?.data?.redirectUrl;
    if (redirectUrl) {
      openBrowser(redirectUrl);
    }
  } catch (err) {
    console.error("Failed to call session.authorize('microsoft_teams'):", err);
  }
}

// The session reports the Teams connection for this user and auth config.
async function ensureTeamsAuth(session: any): Promise<void> {
  const res = await withTimeout(session.toolkits({ toolkits: ["microsoft_teams"] }), 8000);
  const connection = res?.items?.find((t: any) => t?.slug === "microsoft_teams")?.connection;
  if (connection?.isActive) return;

  await handleReauth(session);
  const status = String(connection?.connectedAccount?.status || "").toUpperCase();
  if (["EXPIRED", "REVOKED", "FAILED", "INACTIVE"].includes(status)) {
    throw new Error("Your Teams connection expired or was revoked. Reconnect in your browser, then try again.");
  }
  throw new Error("Connect Microsoft Teams in your browser, then try again.");
}

function getComposioClient(): { composio: Composio; authConfigId: string; userId: string } {
  const apiKey = process.env.COMPOSIO_API_KEY;
  const authConfigId = process.env.COMPOSIO_AUTH_CONFIG_ID;
  const userId = process.env.VOICEOS_USER_ID;

  if (!apiKey || !authConfigId || !userId) {
    throw new Error(
      "Missing required Composio configuration. Please set COMPOSIO_API_KEY, COMPOSIO_AUTH_CONFIG_ID, and VOICEOS_USER_ID in settings."
    );
  }

  const composio = new Composio({
    apiKey,
    baseUrl: `https://${COMPOSIO_API_HOST}`,
  } as any);
  return { composio, authConfigId, userId };
}

async function getTeamsSession(): Promise<{ session: any }> {
  const { composio, authConfigId, userId } = getComposioClient();
  const session = await withTimeout(
    composio.create(userId, {
      toolkits: ["microsoft_teams"],
      authConfigs: { microsoft_teams: authConfigId },
    }),
    8000
  );

  await ensureTeamsAuth(session);
  return { session };
}

async function executeTeamsAction(
  session: any,
  slug: string,
  params: Record<string, any>
): Promise<any> {
  // Thrown errors and failed results are mapped in one place, so a reconnect
  // error is never re-caught and re-mapped (which would request a second link).
  let errText: string | null = null;
  let res: any;
  try {
    res = await withTimeout(
      session.execute(slug, params, { signal: AbortSignal.timeout(8000) } as any),
      8000
    );
    if (res && (res.successful === false || res.error)) {
      errText = typeof res.error === "string" ? res.error : JSON.stringify(res.error || res);
    }
  } catch (err: any) {
    errText = err?.message || String(err || "");
  }
  if (errText === null) return res;

  if (isAuthError(errText)) {
    await handleReauth(session);
    throw new Error("Your Teams connection expired or was revoked. Reconnect in your browser, then try again.");
  }
  if (slug === "MICROSOFT_TEAMS_CREATE_CHANNEL" && isPermissionError(errText)) {
    throw new Error("Creating channels needs admin approval in your organization's Microsoft 365.");
  }
  throw new Error(errText);
}

// LIST_USER_CHAT_MESSAGES rejects "me" and needs the signed-in user's GUID,
// so look it up once from the profile and cache it for this server process.
let cachedUserGuid: string | null = null;

async function getMyUserGuid(session: any): Promise<string> {
  if (cachedUserGuid) return cachedUserGuid;
  const res = await executeTeamsAction(session, "MICROSOFT_TEAMS_GET_MY_PROFILE", {
    user_id: "me",
    select: "id",
  });
  const profile = res?.data?.response_data || res?.response_data || res?.data || {};
  const id = profile?.id || profile?.data?.id;
  if (!id) {
    throw new Error("Couldn't look up your Microsoft account ID. Try again in a moment.");
  }
  cachedUserGuid = String(id);
  return cachedUserGuid;
}

function cleanContent(htmlOrText: string): string {
  if (!htmlOrText) return "";
  return htmlOrText.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
}

function formatMessageTime(timeIsoOrStr: string): string {
  if (!timeIsoOrStr) return "";
  try {
    const d = new Date(timeIsoOrStr);
    if (isNaN(d.getTime())) return timeIsoOrStr;
    const now = new Date();
    const isToday = d.toDateString() === now.toDateString();
    if (isToday) {
      return d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
    }
    const sameYear = d.getFullYear() === now.getFullYear();
    return d.toLocaleDateString(
      "en-US",
      sameYear ? { month: "short", day: "numeric" } : { month: "short", day: "numeric", year: "numeric" }
    );
  } catch {
    return timeIsoOrStr;
  }
}

function formatMeetingTime(startIso: string, endIso: string): string {
  try {
    const s = new Date(startIso);
    const e = new Date(endIso);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) {
      return `${startIso} – ${endIso}`;
    }
    const dateStr = s.toLocaleDateString("en-US", { month: "short", day: "numeric" });
    const sTime = s.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
    const eTime = e.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
    return `${dateStr} · ${sTime} – ${eTime}`;
  } catch {
    return `${startIso} – ${endIso}`;
  }
}

// ---------------------------------------------------------------------------
// Tool 1: get_all_chats (Read)
// ---------------------------------------------------------------------------
const CHATS_CSS = `
.tc{font:13px/1.4 "Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;padding:10px 14px 12px}
.tc-hd{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.tc-title{font-size:12.5px;font-weight:600;color:var(--ink-1)}
.tc-count{margin-left:auto;font-size:10.5px;color:var(--ink-4);font-variant-numeric:tabular-nums}
.tc-rw{display:flex;align-items:center;gap:9px;height:38px;padding:0 8px;border-radius:6px;border-top:1px solid var(--line)}
.tc-rw:first-of-type{border-top:0}
.tc-badge{width:16px;height:16px;border-radius:50%;display:grid;place-items:center;flex:none}
.tc-badge-purple{background:var(--accent);color:var(--fill-1)}
html[data-k-theme=dark] .tc-badge-purple{color:#ffffff}
html[data-k-theme=light] .tc-badge-purple{color:#ffffff}
.tc-badge-neutral{background:var(--fill-2);color:var(--ink-3)}
.tc-topic{flex:1;min-width:0;max-width:220px;font-size:13px;font-weight:500;color:var(--ink-1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tc-time{margin-left:auto;flex:none;font-size:10.5px;color:var(--ink-4);font-variant-numeric:tabular-nums}
.tc-type{margin-left:8px;flex:none;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:var(--ink-4);opacity:.45}
.tc-empty{padding:24px 12px;text-align:center;font-size:12px;color:var(--ink-3)}
`;

server.registerTool(
  "get_all_chats",
  {
    title: "Get all chats",
    description: "Fetch all Microsoft Teams chats for the user. 1:1 chats have no topic, so to find a person's chat, check sender names with get_chat_messages. Use when the user asks for their Teams chats, recent conversations, or to list chat channels.",
    inputSchema: {
      limit: z.number().min(1).max(20).optional().describe("Maximum number of chats to return (default 10)"),
    } as any,
  },
  async (args: any) => {
    const { session } = await getTeamsSession();
    const res = await executeTeamsAction(session, "MICROSOFT_TEAMS_CHATS_GET_ALL_CHATS", {
      user_id: "me",
      top: 50,
    });

    const rawList = res?.data?.response_data || res?.response_data || res?.data?.value || res?.value || res?.data || [];
    const items = Array.isArray(rawList) ? rawList : (rawList?.chats || []);
    const limit = Math.min(Math.max(args?.limit ?? 10, 1), 20);
    // Match the Teams app: drop chats the user has hidden, newest activity first.
    const chats = items
      .filter((c: any) => !(c?.viewpoint?.isHidden ?? c?.viewpoint?.is_hidden))
      .map((c: any) => ({
        id: String(c?.id || ""),
        chatType: String(c?.chatType || c?.chat_type || "oneOnOne"),
        topic: c?.topic ? String(c.topic) : null,
        lastUpdated: String(c?.lastUpdatedDateTime || c?.last_updated_date_time || ""),
      }))
      .sort((a: any, b: any) => (Date.parse(b.lastUpdated) || 0) - (Date.parse(a.lastUpdated) || 0))
      .slice(0, limit);

    const rowsHtml = chats.length === 0
      ? `<div class="tc-empty">No active chats found</div>`
      : chats.map((c: any) => {
          const isGroupOrMeeting = c.chatType === "group" || c.chatType === "meeting";
          const badgeClass = isGroupOrMeeting ? "tc-badge-purple" : "tc-badge-neutral";
          const typeLabel = c.chatType === "oneOnOne" ? "1:1" : c.chatType === "meeting" ? "Meeting" : "Group";
          const title = c.topic || (c.chatType === "oneOnOne" ? "Direct conversation" : `Chat ${clip(c.id, 12)}`);
          return `
            <div class="tc-rw">
              <span class="tc-badge ${badgeClass}">
                <svg width="10" height="10" viewBox="0 0 16 16" fill="currentColor">
                  ${isGroupOrMeeting
                    ? '<path d="M11 6a3 3 0 1 1-6 0 3 3 0 0 1 6 0zM0 14a4 4 0 0 1 8 0H0zm9-1a3 3 0 0 1 6 0h-6z"/>'
                    : '<circle cx="8" cy="6" r="3"/><path d="M2 14a6 6 0 0 1 12 0H2z"/>'}
                </svg>
              </span>
              <span class="tc-topic">${esc(clip(title, 36))}</span>
              <span class="tc-time">${esc(formatMessageTime(c.lastUpdated))}</span>
              <span class="tc-type">${esc(typeLabel)}</span>
            </div>
          `;
        }).join("");

    const cardHeight = chats.length === 0 ? 96 : Math.min(410, 48 + chats.length * 38);
    const card = renderCustom({
      accent: ACCENT,
      css: CHATS_CSS,
      body: `
        <div class="tc">
          <div class="tc-hd">
            ${markHtml()}
            <span class="tc-title">Teams Chats</span>
            <span class="tc-count">${chats.length} active</span>
          </div>
          ${rowsHtml}
        </div>
      `,
      height: cardHeight,
    });

    const payload = {
      chats,
      _voiceos_glance: { blocks: [{ type: "widget", html: card.html, height: card.height }] },
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(payload) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool 2: get_chat_messages (Read)
// ---------------------------------------------------------------------------
const MESSAGES_CSS = `
.tm{font:13px/1.4 "Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;padding:10px 14px 12px}
.tm-hd{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.tm-title{font-size:12.5px;font-weight:600;color:var(--ink-1)}
.tm-count{margin-left:auto;font-size:10.5px;color:var(--ink-4);font-variant-numeric:tabular-nums}
.tm-rw{padding:6px 8px;border-top:1px solid var(--line);min-height:44px;border-radius:6px;box-sizing:border-box}
.tm-rw:first-of-type{border-top:0}
.tm-meta{display:flex;align-items:center;justify-content:space-between;gap:8px}
.tm-sender{font-size:11px;font-weight:600;color:var(--accent);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;flex:1}
.tm-time{font-size:10px;color:var(--ink-4);font-variant-numeric:tabular-nums;flex:none;opacity:.42}
.tm-body{font-size:12.5px;color:var(--ink-2);line-height:1.35;margin-top:2px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;word-break:break-word}
.tm-empty{padding:22px 12px;text-align:center;font-size:11px;color:var(--ink-4)}
`;

server.registerTool(
  "get_chat_messages",
  {
    title: "Get chat messages",
    description: "Fetch recent messages from a Microsoft Teams chat thread. Use when the user asks to read messages from a chat, see what someone sent in Teams, or check conversation history.",
    inputSchema: {
      chat_id: z.string().describe("Teams chat identifier"),
      top: z.number().min(1).max(50).optional().describe("Number of recent messages to retrieve (default 10)"),
    } as any,
  },
  async (args: any) => {
    const { session } = await getTeamsSession();
    const top = Math.min(Math.max(args?.top ?? 10, 1), 50);
    // Over-fetch so system events (call started, member added) can be dropped
    // and the user still gets `top` real messages.
    const res = await executeTeamsAction(session, "MICROSOFT_TEAMS_LIST_USER_CHAT_MESSAGES", {
      user_id: await getMyUserGuid(session),
      chat_id: args.chat_id,
      top: Math.min(top + 15, 50),
    });

    const rawList = res?.data?.response_data || res?.response_data || res?.data?.value || res?.value || res?.data || [];
    const items = Array.isArray(rawList) ? rawList : (rawList?.messages || []);
    const isSystemEvent = (m: any): boolean => {
      const type = m?.messageType || m?.message_type;
      return type ? type !== "message" : !m?.from;
    };
    const messages = items.filter((m: any) => !isSystemEvent(m)).slice(0, top).map((m: any) => {
      const sender = m?.from?.user?.displayName || m?.from?.application?.displayName || m?.sender?.displayName || m?.sender_name || "Unknown";
      const rawContent = m?.body?.content || m?.content || "";
      const text = cleanContent(rawContent);
      const timeStr = m?.createdDateTime || m?.created_at || m?.time || "";
      return {
        id: String(m?.id || ""),
        sender: String(sender),
        text: String(text) || "(attachment or image)",
        time: formatMessageTime(timeStr),
      };
    });

    const rowsHtml = messages.length === 0
      ? `<div class="tm-empty">No messages in this chat</div>`
      : messages.map((m: any) => `
          <div class="tm-rw">
            <div class="tm-meta">
              <span class="tm-sender">${esc(m.sender)}</span>
              <span class="tm-time">${esc(m.time)}</span>
            </div>
            <div class="tm-body">${esc(clip(m.text, 220))}</div>
          </div>
        `).join("");

    const cardHeight = messages.length === 0 ? 96 : Math.min(410, 48 + messages.length * 52);
    const card = renderCustom({
      accent: ACCENT,
      css: MESSAGES_CSS,
      body: `
        <div class="tm">
          <div class="tm-hd">
            ${markHtml()}
            <span class="tm-title">Chat Messages</span>
            <span class="tm-count">${messages.length} messages</span>
          </div>
          ${rowsHtml}
        </div>
      `,
      height: cardHeight,
    });

    const payload = {
      messages,
      _voiceos_glance: { blocks: [{ type: "widget", html: card.html, height: card.height }] },
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(payload) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool 3: send_chat_message (Act)
// ---------------------------------------------------------------------------
const SEND_RESULT_CSS = `
.sm-card{font:13px/1.4 "Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;padding:12px 14px}
.sm-hd{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.sm-title{font-size:13px;font-weight:600;color:var(--ink-1)}
.sm-badge{margin-left:auto;font-size:10px;font-weight:600;letter-spacing:.02em;text-transform:uppercase;color:var(--accent);background:color-mix(in srgb,var(--accent) 18%,transparent);padding:2px 7px;border-radius:999px}
.sm-to{display:flex;align-items:center;gap:6px;font-size:11.5px;margin-bottom:6px}
.sm-lbl{color:var(--ink-4);font-weight:500}
.sm-val{color:var(--ink-1);font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sm-preview{font-size:12.5px;color:var(--ink-2);line-height:1.35;background:var(--fill-1);padding:8px 10px;border-radius:6px;border:1px solid var(--line);word-break:break-word}
`;

server.registerTool(
  "send_chat_message",
  {
    title: "Send chat message",
    description: "Send a direct message or group chat message in Microsoft Teams. Use when the user asks to send, text, or post a message to a person or chat in Teams.",
    inputSchema: {
      chat_id: z.string().describe("Target Microsoft Teams chat identifier"),
      content: z.string().describe("Message text content to send"),
      recipient_name: z.string().optional().describe("Recipient or chat display name for UI display only (not sent to Composio)"),
    } as any,
  },
  async (args: any) => {
    const { session } = await getTeamsSession();
    // Do not send recipient_name to Composio
    await executeTeamsAction(session, "MICROSOFT_TEAMS_TEAMS_POST_CHAT_MESSAGE", {
      chat_id: args.chat_id,
      content: args.content,
    });

    const recipient = args.recipient_name || `Chat ${clip(args.chat_id, 14)}`;
    const card = renderCustom({
      accent: ACCENT,
      css: SEND_RESULT_CSS,
      body: `
        <div class="sm-card">
          <div class="sm-hd">
            ${markHtml()}
            <span class="sm-title">Message Sent</span>
            <span class="sm-badge">Sent</span>
          </div>
          <div class="sm-to">
            <span class="sm-lbl">To</span>
            <span class="sm-val">${esc(recipient)}</span>
          </div>
          <div class="sm-preview">${esc(clip(args.content, 180))}</div>
        </div>
      `,
      height: 135,
    });

    const payload = {
      success: true,
      chat_id: args.chat_id,
      recipient_name: args.recipient_name,
      content: args.content,
      _voiceos_glance: { blocks: [{ type: "widget", html: card.html, height: card.height }] },
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(payload) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool 4: reply_to_message (Act)
// ---------------------------------------------------------------------------
const REPLY_RESULT_CSS = `
.rm-card{font:13px/1.4 "Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;padding:12px 14px}
.rm-hd{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.rm-title{font-size:13px;font-weight:600;color:var(--ink-1)}
.rm-badge{margin-left:auto;font-size:10px;font-weight:600;letter-spacing:.02em;text-transform:uppercase;color:var(--accent);background:color-mix(in srgb,var(--accent) 18%,transparent);padding:2px 7px;border-radius:999px}
.rm-meta{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--ink-4);margin-bottom:6px}
.rm-preview{font-size:12.5px;color:var(--ink-2);line-height:1.35;background:var(--fill-1);padding:8px 10px;border-radius:6px;border:1px solid var(--line);word-break:break-word}
`;

type ReplyTarget = { team_id: string; channel_id: string; message_id: string };

const COPY_LINK_HINT = "In Teams, use More options, then Copy link on the channel post.";

// A channel post's "Copy link" URL carries all three IDs:
// /l/message/{channelId}/{messageId}?groupId={teamId}&parentMessageId={rootId}
function parseMessageLink(link: string): ReplyTarget {
  let url: URL;
  try {
    url = new URL(link.trim());
  } catch {
    throw new Error(`That isn't a valid Teams link. ${COPY_LINK_HINT}`);
  }
  const match = url.pathname.match(/\/l\/message\/([^/]+)\/([^/]+)/);
  const teamId = url.searchParams.get("groupId");
  if (!match || !teamId) {
    throw new Error(
      "That link isn't to a team channel post. To reply in a 1:1 or group chat, send a message to that chat instead."
    );
  }
  // A link copied from a reply inside a thread points at that reply; Graph wants the root post.
  const messageId = url.searchParams.get("parentMessageId") || decodeURIComponent(match[2]);
  return { team_id: teamId, channel_id: decodeURIComponent(match[1]), message_id: messageId };
}

function replyTarget(args: any): ReplyTarget {
  if (args?.message_link) return parseMessageLink(String(args.message_link));
  if (args?.team_id && args?.channel_id && args?.message_id) {
    return { team_id: args.team_id, channel_id: args.channel_id, message_id: args.message_id };
  }
  throw new Error(`To reply to a channel post, I need its link. ${COPY_LINK_HINT}`);
}

server.registerTool(
  "reply_to_message",
  {
    title: "Reply to message",
    description: "Reply to a post in a Microsoft Teams team channel. Use when the user asks to reply to a channel post or thread in a team. Pass the post's link (from Copy link on the post in Teams) as message_link; the team, channel and message IDs are only a fallback. Only for team channels: to reply in a 1:1 or group chat, use send_chat_message in that chat.",
    inputSchema: {
      message_link: z.string().optional().describe("Link to the channel post, from Copy link on the post in Teams. Contains the team, channel and message IDs."),
      team_id: z.string().optional().describe("Team ID. Not needed when message_link is given"),
      channel_id: z.string().optional().describe("Channel ID within the team. Not needed when message_link is given"),
      message_id: z.string().optional().describe("Root post ID to reply to. Not needed when message_link is given"),
      content: z.string().describe("Reply message content"),
    } as any,
  },
  async (args: any) => {
    const target = replyTarget(args);
    const { session } = await getTeamsSession();
    await executeTeamsAction(session, "MICROSOFT_TEAMS_POST_MESSAGE_REPLY", {
      ...target,
      content: args.content,
    });

    const card = renderCustom({
      accent: ACCENT,
      css: REPLY_RESULT_CSS,
      body: `
        <div class="rm-card">
          <div class="rm-hd">
            ${markHtml()}
            <span class="rm-title">Reply Sent</span>
            <span class="rm-badge">Replied</span>
          </div>
          <div class="rm-meta">
            <span>Thread: ${esc(clip(target.message_id, 16))}</span>
          </div>
          <div class="rm-preview">${esc(clip(args.content, 180))}</div>
        </div>
      `,
      height: 125,
    });

    const payload = {
      success: true,
      ...target,
      content: args.content,
      _voiceos_glance: { blocks: [{ type: "widget", html: card.html, height: card.height }] },
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(payload) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool 5: create_online_meeting (Act)
// ---------------------------------------------------------------------------
const MEETING_RESULT_CSS = `
.mm-card{font:13px/1.4 "Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;padding:12px 14px}
.mm-hd{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.mm-title{font-size:13px;font-weight:600;color:var(--ink-1)}
.mm-badge{margin-left:auto;font-size:10px;font-weight:600;letter-spacing:.02em;text-transform:uppercase;color:var(--accent);background:color-mix(in srgb,var(--accent) 18%,transparent);padding:2px 7px;border-radius:999px}
.mm-subj{font-size:14px;font-weight:600;color:var(--ink-1);margin-bottom:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mm-time{font-size:11.5px;color:var(--ink-3);font-variant-numeric:tabular-nums;margin-bottom:10px}
.mm-join{display:inline-flex;align-items:center;gap:6px;background:var(--accent);color:var(--fill-1);font-size:12px;font-weight:600;text-decoration:none;padding:6px 12px;border-radius:6px}
html[data-k-theme=dark] .mm-join{color:#ffffff}
html[data-k-theme=light] .mm-join{color:#ffffff}
`;

server.registerTool(
  "create_online_meeting",
  {
    title: "Create online meeting",
    description: "Schedule a new Microsoft Teams online meeting with start and end times. Use when the user asks to schedule a Teams meeting, set up a Teams call, or create an online calendar meeting.",
    inputSchema: {
      subject: z.string().describe("Meeting topic or subject"),
      start_date_time: z.string().describe("Meeting start time in ISO 8601 with local offset (e.g. 2025-03-20T15:00:00-04:00)"),
      end_date_time: z.string().describe("Meeting end time in ISO 8601 with local offset (e.g. 2025-03-20T15:30:00-04:00)"),
    } as any,
  },
  async (args: any) => {
    const { session } = await getTeamsSession();
    const res = await executeTeamsAction(session, "MICROSOFT_TEAMS_CREATE_MEETING", {
      subject: args.subject,
      start_date_time: args.start_date_time,
      end_date_time: args.end_date_time,
    });

    const data = res?.data?.response_data || res?.response_data || res?.data || {};
    const joinWebUrl = data?.joinWebUrl || data?.join_web_url || data?.webUrl || "";
    const meetingId = data?.id || "created";
    const subject = data?.subject || args.subject;
    const timeFormatted = formatMeetingTime(args.start_date_time, args.end_date_time);
    const joinLinkHtml = joinWebUrl
      ? `<a class="mm-join" href="${esc(joinWebUrl)}">Join Teams meeting</a>`
      : `<span class="mm-join">Meeting scheduled</span>`;

    const card = renderCustom({
      accent: ACCENT,
      css: MEETING_RESULT_CSS,
      body: `
        <div class="mm-card">
          <div class="mm-hd">
            ${markHtml()}
            <span class="mm-title">Meeting Created</span>
            <span class="mm-badge">Created</span>
          </div>
          <div class="mm-subj">${esc(subject)}</div>
          <div class="mm-time">${esc(timeFormatted)}</div>
          ${joinLinkHtml}
        </div>
      `,
      height: 155,
    });

    const payload = {
      success: true,
      id: meetingId,
      subject,
      joinWebUrl,
      start_date_time: args.start_date_time,
      end_date_time: args.end_date_time,
      _voiceos_glance: { blocks: [{ type: "widget", html: card.html, height: card.height }] },
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(payload) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool 6: create_channel (Act)
// ---------------------------------------------------------------------------
const CHANNEL_RESULT_CSS = `
.cm-card{font:13px/1.4 "Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;padding:12px 14px}
.cm-hd{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.cm-title{font-size:13px;font-weight:600;color:var(--ink-1)}
.cm-badge{margin-left:auto;font-size:10px;font-weight:600;letter-spacing:.02em;text-transform:uppercase;color:var(--accent);background:color-mix(in srgb,var(--accent) 18%,transparent);padding:2px 7px;border-radius:999px}
.cm-name{display:flex;align-items:center;gap:4px;font-size:14px;font-weight:600;color:var(--ink-1);margin-bottom:4px}
.cm-hash{color:var(--accent);font-weight:700}
.cm-team{font-size:11px;color:var(--ink-4)}
`;

server.registerTool(
  "create_channel",
  {
    title: "Create channel",
    description: "Create a new channel inside a Microsoft Teams team. Use when the user asks to create a channel, add a channel to a team, or set up a new Teams topic channel.",
    inputSchema: {
      team_id: z.string().describe("Target Microsoft Teams team ID"),
      name: z.string().describe("Channel name"),
    } as any,
  },
  async (args: any) => {
    const { session } = await getTeamsSession();
    const res = await executeTeamsAction(session, "MICROSOFT_TEAMS_CREATE_CHANNEL", {
      team_id: args.team_id,
      name: args.name,
    });

    const data = res?.data?.response_data || res?.response_data || res?.data || {};
    const channelId = data?.id || "created";
    const channelName = data?.displayName || data?.name || args.name;

    const card = renderCustom({
      accent: ACCENT,
      css: CHANNEL_RESULT_CSS,
      body: `
        <div class="cm-card">
          <div class="cm-hd">
            ${markHtml()}
            <span class="cm-title">Channel Created</span>
            <span class="cm-badge">Active</span>
          </div>
          <div class="cm-name">
            <span class="cm-hash">#</span>
            <span>${esc(channelName)}</span>
          </div>
          <div class="cm-team">Team ID: ${esc(clip(args.team_id, 24))}</div>
        </div>
      `,
      height: 120,
    });

    const payload = {
      success: true,
      id: channelId,
      name: channelName,
      team_id: args.team_id,
      _voiceos_glance: { blocks: [{ type: "widget", html: card.html, height: card.height }] },
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(payload) }] };
  }
);

await server.connect(new StdioServerTransport());
