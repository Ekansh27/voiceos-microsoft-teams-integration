// Preloaded into the server for offline tests: replaces fetch with a fake Composio and
// logs each call to stderr. MOCK_CONNECTION picks the Teams connection state the fake
// reports: "active" (default), "none" or "expired".
const CONNECTION = process.env.MOCK_CONNECTION ?? "active";
const SESSION_ID = "trs_test";
const USER_GUID = "8ea0e38b-efb3-4757-924a-5f94061cf8c2";

const respond = (data: unknown, status = 200) =>
  new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json" } });
const completed = (data: Record<string, unknown>) =>
  respond({ data, error: null, log_id: "log_test", result_type: "completed" });
const failed = (error: string) =>
  respond({ data: {}, error, log_id: "log_test", result_type: "failed" });

const chats = [
  { id: "19:old", chatType: "group", topic: "Old project", lastUpdatedDateTime: "2025-03-01T10:00:00Z", viewpoint: { isHidden: false } },
  { id: "19:hidden", chatType: "group", topic: "Hidden chat", lastUpdatedDateTime: "2026-10-31T10:00:00Z", viewpoint: { isHidden: true } },
  { id: "19:newest", chatType: "meeting", topic: "Standup", lastUpdatedDateTime: "2026-10-31T09:00:00Z", viewpoint: { isHidden: false } },
  { id: "19:middle", chatType: "oneOnOne", topic: null, lastUpdatedDateTime: "2026-06-19T12:00:00Z", viewpoint: { isHidden: false } },
];

const messages = [
  { id: "m1", messageType: "message", from: { user: { displayName: "Priya" } }, body: { content: "<p>See you at 3</p>" }, createdDateTime: "2026-10-31T09:00:00Z" },
  { id: "m2", messageType: "systemEventMessage", from: null, body: { content: "<systemEventMessage/>" }, createdDateTime: "2026-10-31T08:59:00Z" },
  { id: "m3", messageType: "message", from: { user: { displayName: "Sam" } }, body: { content: "" }, createdDateTime: "2026-10-30T15:00:00Z" },
  { id: "m4", messageType: "message", from: { user: { displayName: "Priya" } }, body: { content: "Older message" }, createdDateTime: "2025-12-01T15:00:00Z" },
];

function execute(slug: string, args: Record<string, any>): Response {
  switch (slug) {
    case "MICROSOFT_TEAMS_CHATS_GET_ALL_CHATS":
      return completed({ value: chats });
    case "MICROSOFT_TEAMS_GET_MY_PROFILE":
      return completed({ id: USER_GUID, displayName: "Test User" });
    case "MICROSOFT_TEAMS_LIST_USER_CHAT_MESSAGES":
      if (args.user_id !== USER_GUID) return failed("user ID must be a GUID");
      return completed({ value: messages.slice(0, args.top ?? 10) });
    case "MICROSOFT_TEAMS_TEAMS_POST_CHAT_MESSAGE":
      if (args.chat_id === "19:revoked") return failed("401 Unauthorized: token has been revoked");
      return completed({ id: "m-new", body: { content: args.content } });
    case "MICROSOFT_TEAMS_POST_MESSAGE_REPLY":
      return completed({ id: "r-new", replyToId: args.message_id });
    case "MICROSOFT_TEAMS_CREATE_MEETING":
      return completed({ id: "meet-1", subject: args.subject, joinWebUrl: "https://teams.microsoft.com/l/meetup-join/test" });
    case "MICROSOFT_TEAMS_CREATE_CHANNEL":
      return failed("403 Forbidden: Missing scope permissions on the request. API requires one of 'Channel.Create'.");
    default:
      return failed(`Unexpected tool: ${slug}`);
  }
}

function toolkits(): Response {
  const connected_account =
    CONNECTION === "none"
      ? undefined
      : {
          id: "ca_test",
          status: CONNECTION === "expired" ? "EXPIRED" : "ACTIVE",
          auth_config: { id: "ac_test", auth_scheme: "OAUTH2", is_composio_managed: true },
        };
  return respond({
    items: [{ slug: "microsoft_teams", name: "Microsoft Teams", is_no_auth: false, connected_account }],
    next_cursor: null,
    total_pages: 1,
  });
}

globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
  const request = input instanceof Request ? input : new Request(input, init);
  const url = new URL(request.url);
  const path = url.pathname.replace(/^\/api\/v3(\.1)?/, "");
  const text = await request.clone().text();
  const body = text ? JSON.parse(text) : undefined;
  console.error(`MOCKCALL ${JSON.stringify({ method: request.method, host: url.host, path, query: url.search, body })}`);

  // The SDK's own version check and telemetry: answer them so tests never reach the network.
  if (url.host !== "backend.composio.dev") return respond({});

  if (request.method === "POST" && path === "/tool_router/session") {
    return respond({
      session_id: SESSION_ID,
      mcp: { type: "http", url: "https://backend.composio.dev/tool_router/mcp" },
      config: { user_id: body?.user_id, preload: { tools: [] }, workbench: { enable: false } },
      config_version: 1,
      warnings: [],
    });
  }
  if (path === `/tool_router/session/${SESSION_ID}/toolkits`) return toolkits();
  if (path === `/tool_router/session/${SESSION_ID}/link`) {
    return respond({ redirect_url: "https://connect.composio.dev/link/test", connected_account_id: "ca_new", link_token: "lt_test" });
  }
  if (path === `/tool_router/session/${SESSION_ID}/execute`) return execute(body?.tool_slug, body?.arguments ?? {});
  return respond({ error: { message: `Unexpected request: ${request.method} ${path}` } }, 404);
}) as typeof fetch;
