// Offline tests: runs the real server over MCP against tests/mock-composio.ts. No keys or network needed.
import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
const MOCK = join(ROOT, "tests", "mock-composio.ts");
const manifest = JSON.parse(readFileSync(join(ROOT, "voiceos.integration.json"), "utf8"));
const source = readFileSync(join(ROOT, "server.ts"), "utf8");

const APPROVED_TOOLS = [
  "get_all_chats",
  "get_chat_messages",
  "send_chat_message",
  "reply_to_message",
  "create_online_meeting",
  "create_channel",
];
const WRITE_TOOLS = ["send_chat_message", "reply_to_message", "create_online_meeting", "create_channel"];
const APPROVED_SLUGS = [
  "MICROSOFT_TEAMS_CHATS_GET_ALL_CHATS",
  "MICROSOFT_TEAMS_CREATE_CHANNEL",
  "MICROSOFT_TEAMS_CREATE_MEETING",
  "MICROSOFT_TEAMS_GET_MY_PROFILE",
  "MICROSOFT_TEAMS_LIST_USER_CHAT_MESSAGES",
  "MICROSOFT_TEAMS_POST_MESSAGE_REPLY",
  "MICROSOFT_TEAMS_TEAMS_POST_CHAT_MESSAGE",
];
const SETUP_ENV = { COMPOSIO_API_KEY: "test", COMPOSIO_AUTH_CONFIG_ID: "ac_test", VOICEOS_USER_ID: "tester" };

type Server = { client: Client; log: string[] };

// PATH points nowhere so the server's `open` call fails instead of opening a real browser.
async function startServer(env: Record<string, string>): Promise<Server> {
  const log: string[] = [];
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: ["--preload", MOCK, join(ROOT, "server.ts")],
    stderr: "pipe",
    env: { PATH: "/nonexistent", HOME: mkdtempSync(join(tmpdir(), "teams-test-")), ...env },
  });
  transport.stderr?.on("data", (chunk: Buffer) => log.push(...String(chunk).split("\n").filter(Boolean)));
  const client = new Client({ name: "test", version: "1.0.0" });
  await client.connect(transport);
  return { client, log };
}

const mockCalls = (server: Server) =>
  server.log.filter((line) => line.startsWith("MOCKCALL ")).map((line) => JSON.parse(line.slice(9)));
const executed = (server: Server) =>
  mockCalls(server).filter((c) => c.path.endsWith("/execute")).map((c) => c.body as { tool_slug: string; arguments: any });
const lastArguments = (server: Server, slug: string) => executed(server).filter((c) => c.tool_slug === slug).at(-1)?.arguments;
const linkRequests = (server: Server) => mockCalls(server).filter((c) => c.path.endsWith("/link"));

async function run(server: Server, name: string, args: Record<string, unknown>) {
  const result: any = await server.client.callTool({ name, arguments: args });
  const text: string = result.content[0].text;
  let body: any = {};
  try {
    body = JSON.parse(text);
  } catch {
    body = { error: text };
  }
  const block = body._voiceos_glance?.blocks?.[0];
  return { isError: Boolean(result.isError), text, body, block, html: String(block?.html ?? "") };
}

describe("manifest", () => {
  test("follows VoiceOS's limits", () => {
    expect(manifest.id).toBe("com.voiceos.generated.microsoft-teams");
    expect(manifest.version).toMatch(/^\d+\.\d+\.\d+$/);
    expect(manifest.summary.length).toBeLessThanOrEqual(140);
  });

  test("declares exactly the six approved tools", () => {
    expect(manifest.tools.map((t: any) => t.name)).toEqual(APPROVED_TOOLS);
  });

  test("allows only the Composio domain", () => {
    expect(manifest.permissions).toEqual([{ kind: "network", domains: ["backend.composio.dev"] }]);
  });

  test("puts a confirmation card on every write and on no read", () => {
    for (const tool of manifest.tools) {
      expect(Boolean(tool.confirmation)).toBe(WRITE_TOOLS.includes(tool.name));
    }
  });

  test("binds every confirmation field to a real tool input", () => {
    for (const tool of manifest.tools.filter((t: any) => t.confirmation)) {
      const bound = [...String(tool.confirmation.root.html).matchAll(/data-voiceos-key="([a-z_]+)"/g)].map((m) => m[1]);
      expect(bound.length).toBeGreaterThan(0);
      for (const key of bound) expect(Object.keys(tool.inputSchema.properties)).toContain(key);
    }
  });

  test("declares every environment variable the server reads", () => {
    const read = [...new Set([...source.matchAll(/process\.env\.([A-Z_]+)/g)].map((m) => m[1]))].sort();
    const declared = manifest.preferences.map((p: any) => p.name).sort();
    expect(read).toEqual(declared);
    for (const pref of manifest.preferences) expect(pref.required).toBe(true);
  });

  test("calls only approved Composio tools", () => {
    const slugs = [...new Set([...source.matchAll(/"(MICROSOFT_TEAMS_[A-Z_]+)"/g)].map((m) => m[1]))].sort();
    expect(slugs).toEqual(APPROVED_SLUGS);
  });
});

describe("connected account", () => {
  let server: Server;

  beforeAll(async () => {
    server = await startServer({ ...SETUP_ENV, MOCK_CONNECTION: "active" });
  });

  afterAll(async () => {
    await server.client.close();
  });

  test("the server exposes the same six tools and inputs as the manifest", async () => {
    const { tools } = await server.client.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual([...APPROVED_TOOLS].sort());
    for (const tool of tools) {
      const declared = manifest.tools.find((t: any) => t.name === tool.name);
      expect(Object.keys(tool.inputSchema.properties ?? {}).sort()).toEqual(Object.keys(declared.inputSchema.properties).sort());
    }
  });

  test("get_all_chats skips hidden chats, lists the rest newest first, and scopes the session to Teams", async () => {
    const r = await run(server, "get_all_chats", {});
    expect(r.isError).toBe(false);
    expect(r.body.chats.map((c: any) => c.id)).toEqual(["19:newest", "19:middle", "19:old"]);
    expect(lastArguments(server, "MICROSOFT_TEAMS_CHATS_GET_ALL_CHATS")).toEqual({ user_id: "me", top: 50 });
    expect(r.block.type).toBe("widget");
    expect(r.html).toContain("Teams Chats");

    const create = mockCalls(server).find((c) => c.path === "/tool_router/session");
    expect(create.body.user_id).toBe("tester");
    expect(create.body.auth_configs).toEqual({ microsoft_teams: "ac_test" });
    expect(create.body.toolkits).toEqual({ enable: ["microsoft_teams"] });
  });

  test("get_all_chats respects the limit", async () => {
    const r = await run(server, "get_all_chats", { limit: 1 });
    expect(r.body.chats.map((c: any) => c.id)).toEqual(["19:newest"]);
  });

  test("get_chat_messages passes the signed-in user's ID and skips system events", async () => {
    const r = await run(server, "get_chat_messages", { chat_id: "19:newest", top: 3 });
    expect(r.isError).toBe(false);
    expect(lastArguments(server, "MICROSOFT_TEAMS_LIST_USER_CHAT_MESSAGES")).toEqual({
      user_id: "8ea0e38b-efb3-4757-924a-5f94061cf8c2",
      chat_id: "19:newest",
      top: 18,
    });
    expect(r.body.messages.map((m: any) => m.id)).toEqual(["m1", "m3", "m4"]);
    expect(r.body.messages[0]).toMatchObject({ sender: "Priya", text: "See you at 3" });
    expect(r.body.messages[1].text).toBe("(attachment or image)");
    expect(r.body.messages[2].time).toBe("Dec 1, 2025");
  });

  test("the user's ID is looked up once and reused", async () => {
    await run(server, "get_chat_messages", { chat_id: "19:middle", top: 1 });
    const lookups = executed(server).filter((c) => c.tool_slug === "MICROSOFT_TEAMS_GET_MY_PROFILE");
    expect(lookups).toHaveLength(1);
    expect(lookups[0].arguments).toEqual({ user_id: "me", select: "id" });
  });

  test("send_chat_message sends only the chat and text, not the display name", async () => {
    const r = await run(server, "send_chat_message", { chat_id: "19:newest", content: "On my way", recipient_name: "Priya" });
    expect(r.isError).toBe(false);
    expect(lastArguments(server, "MICROSOFT_TEAMS_TEAMS_POST_CHAT_MESSAGE")).toEqual({ chat_id: "19:newest", content: "On my way" });
    expect(r.body).toMatchObject({ success: true, recipient_name: "Priya", content: "On my way" });
    expect(r.html).toContain("Message Sent");
    expect(r.html).toContain("Priya");
  });

  test("reply_to_message passes the team, channel and message IDs", async () => {
    const args = { team_id: "team-1", channel_id: "19:channel@thread.tacv2", message_id: "1616990032035", content: "Thanks" };
    const r = await run(server, "reply_to_message", args);
    expect(r.isError).toBe(false);
    expect(lastArguments(server, "MICROSOFT_TEAMS_POST_MESSAGE_REPLY")).toEqual(args);
    expect(r.html).toContain("Reply Sent");
  });

  test("reply_to_message reads the team, channel and post IDs from a post's link", async () => {
    const link =
      "https://teams.microsoft.com/l/message/19%3A4a95f7d8db4c4e7fae857bcebe0623e6%40thread.tacv2/1616990032035" +
      "?groupId=fbe2bf47-16c8-47cf-b4a5-4b9b187c508b&tenantId=2432b57b-0abd-43db-aa7b-16eadd115d34&createdTime=1616990032035";
    const r = await run(server, "reply_to_message", { message_link: link, content: "Sounds good" });
    expect(r.isError).toBe(false);
    expect(lastArguments(server, "MICROSOFT_TEAMS_POST_MESSAGE_REPLY")).toEqual({
      team_id: "fbe2bf47-16c8-47cf-b4a5-4b9b187c508b",
      channel_id: "19:4a95f7d8db4c4e7fae857bcebe0623e6@thread.tacv2",
      message_id: "1616990032035",
      content: "Sounds good",
    });
  });

  test("reply_to_message replies to the root post when the link points at a reply in the thread", async () => {
    const link =
      "https://teams.microsoft.com/l/message/19%3A4a95f7d8db4c4e7fae857bcebe0623e6%40thread.tacv2/1616990171266" +
      "?groupId=fbe2bf47-16c8-47cf-b4a5-4b9b187c508b&parentMessageId=1616990032035&createdTime=1616990171266";
    await run(server, "reply_to_message", { message_link: link, content: "Same here" });
    expect(lastArguments(server, "MICROSOFT_TEAMS_POST_MESSAGE_REPLY").message_id).toBe("1616990032035");
  });

  test("reply_to_message refuses chat links and missing targets without calling Composio", async () => {
    const before = executed(server).length;

    const chatLink = "https://teams.microsoft.com/l/message/19:abc@thread.v2/1616990032035?context=%7B%22contextType%22%3A%22chat%22%7D";
    const chat = await run(server, "reply_to_message", { message_link: chatLink, content: "hi" });
    expect(chat.isError).toBe(true);
    expect(chat.text).toContain("isn't to a team channel post");

    const notALink = await run(server, "reply_to_message", { message_link: "general channel", content: "hi" });
    expect(notALink.isError).toBe(true);
    expect(notALink.text).toContain("isn't a valid Teams link");

    const nothing = await run(server, "reply_to_message", { content: "hi" });
    expect(nothing.isError).toBe(true);
    expect(nothing.text).toContain("Copy link");

    expect(executed(server).length).toBe(before);
  });

  test("create_online_meeting passes the times as given and returns the join link", async () => {
    const args = { subject: "Sync", start_date_time: "2026-11-02T15:00:00-06:00", end_date_time: "2026-11-02T15:30:00-06:00" };
    const r = await run(server, "create_online_meeting", args);
    expect(r.isError).toBe(false);
    expect(lastArguments(server, "MICROSOFT_TEAMS_CREATE_MEETING")).toEqual(args);
    expect(r.body.joinWebUrl).toBe("https://teams.microsoft.com/l/meetup-join/test");
    expect(r.html).toContain("Meeting Created");
  });

  test("create_channel explains that it needs admin approval", async () => {
    const r = await run(server, "create_channel", { team_id: "team-1", name: "voiceos-test" });
    expect(r.isError).toBe(true);
    expect(r.text).toBe("Creating channels needs admin approval in your organization's Microsoft 365.");
  });

  test("a revoked token during a call asks the user to reconnect", async () => {
    const before = linkRequests(server).length;
    const r = await run(server, "send_chat_message", { chat_id: "19:revoked", content: "hi" });
    expect(r.isError).toBe(true);
    expect(r.text).toBe("Your Teams connection expired or was revoked. Reconnect in your browser, then try again.");
    expect(linkRequests(server).length).toBe(before + 1);
  });

  test("only approved Composio tools were called", () => {
    for (const call of executed(server)) expect(APPROVED_SLUGS).toContain(call.tool_slug);
  });

  // The SDK prints its "Upgrade available" notice to stdout, which is the MCP channel.
  test("never runs the SDK's npm version check", () => {
    expect(mockCalls(server).filter((c) => c.host === "registry.npmjs.org")).toEqual([]);
  });
});

describe("no connected account", () => {
  let server: Server;

  beforeAll(async () => {
    server = await startServer({ ...SETUP_ENV, MOCK_CONNECTION: "none" });
  });

  afterAll(async () => {
    await server.client.close();
  });

  test("asks the user to connect, starts sign-in, and runs nothing", async () => {
    const r = await run(server, "get_all_chats", {});
    expect(r.isError).toBe(true);
    expect(r.text).toBe("Connect Microsoft Teams in your browser, then try again.");
    expect(linkRequests(server)).toHaveLength(1);
    expect(linkRequests(server)[0].body).toEqual({ toolkit: "microsoft_teams" });
    expect(executed(server)).toHaveLength(0);
  });
});

describe("expired connection", () => {
  let server: Server;

  beforeAll(async () => {
    server = await startServer({ ...SETUP_ENV, MOCK_CONNECTION: "expired" });
  });

  afterAll(async () => {
    await server.client.close();
  });

  test("asks the user to reconnect and runs nothing", async () => {
    const r = await run(server, "send_chat_message", { chat_id: "19:newest", content: "hi" });
    expect(r.isError).toBe(true);
    expect(r.text).toBe("Your Teams connection expired or was revoked. Reconnect in your browser, then try again.");
    expect(linkRequests(server)).toHaveLength(1);
    expect(executed(server)).toHaveLength(0);
  });
});

describe("missing setup", () => {
  let server: Server;

  beforeAll(async () => {
    server = await startServer({});
  });

  afterAll(async () => {
    await server.client.close();
  });

  test("says which settings are missing", async () => {
    const r = await run(server, "get_all_chats", {});
    expect(r.isError).toBe(true);
    expect(r.text).toContain("COMPOSIO_API_KEY, COMPOSIO_AUTH_CONFIG_ID, and VOICEOS_USER_ID");
    expect(mockCalls(server)).toHaveLength(0);
  });
});
