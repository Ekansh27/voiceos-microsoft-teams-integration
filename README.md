# Microsoft Teams for VoiceOS

A VoiceOS integration (UIUC-20) that lets the VoiceOS agent read Teams chats, send messages and replies, and create online meetings and channels. It is a local MCP server that runs every action through Composio's `microsoft_teams` toolkit, so the code never handles Microsoft tokens.

Built with the VoiceOS Integration Studio, then fixed and tidied by hand. `AGENTS.md` describes the VoiceOS integration contract.

## Tools

Exactly the six tools approved in the ticket. The two reads run immediately; the four writes show a confirmation card first, where the user can edit the arguments.

| Tool | Type | Composio tool |
| --- | --- | --- |
| `get_all_chats` | Read | `MICROSOFT_TEAMS_CHATS_GET_ALL_CHATS` |
| `get_chat_messages` | Read | `MICROSOFT_TEAMS_LIST_USER_CHAT_MESSAGES` |
| `send_chat_message` | Write, confirms first | `MICROSOFT_TEAMS_TEAMS_POST_CHAT_MESSAGE` |
| `reply_to_message` | Write, confirms first | `MICROSOFT_TEAMS_POST_MESSAGE_REPLY` |
| `create_online_meeting` | Write, confirms first | `MICROSOFT_TEAMS_CREATE_MEETING` |
| `create_channel` | Write, confirms first | `MICROSOFT_TEAMS_CREATE_CHANNEL` |

`get_chat_messages` also calls `MICROSOFT_TEAMS_GET_MY_PROFILE` internally to look up the signed-in user's ID, because the message-listing tool rejects `"me"`. It is not exposed to the agent.

### Account tools

Integration Studio has no native account linking, so these three follow Todoist's. They are account-linking tools, not Teams actions, so they don't count toward the six above. All three can be called from the account card as well as by voice.

| Tool | Type | What it does |
| --- | --- | --- |
| `connect_account` | Confirms first, runs in the background | Opens the Microsoft sign-in in the browser and waits up to 2 minutes for it to finish. If Teams is already connected, it only checks that |
| `connection_status` | Read | Shows the account card: connected, expired, revoked or waiting for sign-in, with Connect or Disconnect and Refresh buttons. Never starts a sign-in |
| `disconnect_account` | Confirms first | Revokes access where Composio supports it, then deletes this user's Teams connections for this auth config. Chats and messages aren't touched |

## Setup

1. In the Composio dashboard (Platform view), create an auth config for **Microsoft Teams** with **OAuth2** and **Composio managed** credentials, and set these 7 scopes:
   `Chat.ReadBasic`, `Chat.Read`, `ChatMessage.Send`, `ChannelMessage.Send`, `OnlineMeetings.ReadWrite`, `offline_access`, `User.Read`
2. Create a Composio API key with **Session management** (read and write) and **Session tool execution** (write), plus read access to Tools, Toolkits and Auth configs, and read and write access to Connected accounts (disconnecting needs write).
3. Install the folder in VoiceOS and fill in the setup fields:

| Field | Value |
| --- | --- |
| `COMPOSIO_API_KEY` | Your Composio API key (secret) |
| `COMPOSIO_AUTH_CONFIG_ID` | The auth config's ID (`ac_...`) |
| `VOICEOS_USER_ID` | A stable ID for the user, such as a netid. Not `default`, not an email |

Say "connect me to Teams", or say "show my Teams connection" and use the Connect button. If you skip this, the first Teams tool you use opens the Microsoft sign-in as a fallback. Sign in, then ask again.

## Sign-in and errors

- Before each call, the server asks the Composio session whether Teams is connected. If it isn't, or a call fails with an auth error, it opens the sign-in and the agent says to reconnect. Composio refreshes tokens on its own using `offline_access`.
- Errors are thrown with short, spoken-friendly messages. Composio calls time out after 8 seconds.
- Network access is limited to `backend.composio.dev`.

## Testing

`bun install && bun test` runs 50 offline tests. They start the real server over MCP with `tests/mock-composio.ts` preloaded, which replaces `fetch` with a fake Composio, so no keys or network are needed. They cover the manifest (the six tools plus the account tools, confirmations only on writes, card fields bound to real inputs, declared env vars, the network allowlist), every tool's arguments and results (including reading IDs from a post's link), the not-connected, expired, revoked-token and missing-setup paths, and connecting and disconnecting (finished, unfinished and failed sign-in, duplicate connections, a key without write access).

Tested live on Nov 1, 2026 with a UIUC Microsoft 365 account:

- Get all chats, Get chat messages, Send chat message (including editing the text and declining), Create online meeting (correct local time, working join link)
- Reply to message, from a channel post's copied link
- Create channel returns the expected admin-approval error
- Deleting the Composio connection gives the reconnect message, and one sign-in recovers

Tested live on Nov 4, 2026, on version 1.1.0:

- The Schedule card shows a readable time instead of raw timestamps
- The account card, and Disconnect up to the clear error for a key without write access
- Reply to message, after rewording its description so the agent uses it instead of VoiceOS's built-in draft

A successful disconnect and the Connect button's sign-in haven't been tested live, because they need a key with write access to Connected accounts. So far they've only run against the fake Composio.

## Known limitations

- **Create channel needs admin approval.** Microsoft always requires an org admin to approve `Channel.Create`, and including it blocked sign-in for everyone, so it is left out of the scopes. The tool returns a clear "needs admin approval" error until an admin approves it.
- **Reply needs the post's link.** The user copies it with Copy link on the post in Teams, and the server reads the team, channel and post IDs from it. Finding the post by name or author would need `ChannelMessage.Read.All`, which always requires admin approval, and integrations don't get screen context from VoiceOS. Create channel still needs a team ID.
- **Scheduled meetings don't appear on the calendar.** Create online meeting makes the Teams meeting and a working join link, but no calendar entry shows up. This will be fixed separately.
- **Some old meeting chats still appear in Get all chats.** Chats hidden by the user are skipped, but Teams also hides inactive meeting chats in its own UI without marking them hidden in Microsoft's data. They show up below the recent chats, since the list is sorted newest first.
- **Personal Microsoft accounts don't work.** Microsoft's Teams API only supports licensed work or school accounts.
- **Don't use the Studio's refine on this folder.** It regenerates `server.ts` and would undo the hand fixes to the sign-in check and message lookup.
