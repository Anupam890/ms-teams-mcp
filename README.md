# @anupam365/ms-teams-mcp

MCP server for Microsoft Teams via the Graph API. Exposes **73 tools** for managing teams, channels, chats, messages, meetings, planner, calendar, apps, tabs, scheduling, search, and authentication.

## Prerequisites

- **Node.js 18+**
- **Microsoft 365 subscription** with Teams enabled
- **Azure AD app registration** with appropriate Graph API permissions

## Setup

### 1. Create an Azure AD App

1. Go to [Azure Portal](https://portal.azure.com) → **App Registrations** → **New registration**
2. Name it (e.g., `Teams MCP`), leave redirect URI blank, click **Register**
3. Copy the **Application (client) ID** and **Directory (tenant) ID**

### 2. Add API Permissions

**API Permissions** → **Add a permission** → **Microsoft Graph**

#### Application permissions (for client-credentials auth):

| Permission | Type | Required for |
|---|---|---|
| `Channel.Create` | Application | Creating channels |
| `Channel.Delete.All` | Application | Deleting channels |
| `Channel.ReadBasic.All` | Application | Reading channel metadata |
| `ChannelMessage.Read.All` | Application | Reading channel messages |
| `ChannelMessage.UpdatePolicyViolation.All` | Application | Message policy |
| `Calendars.ReadWrite` | Application | Creating calendar events |
| `Group.Create` | Application | Creating teams |
| `Group.Read.All` | Application | Listing teams |
| `Group.ReadWrite.All` | Application | Updating teams |
| `Member.Read.Hidden` | Application | Reading hidden memberships |
| `OnlineMeetings.ReadWrite.All` | Application | Creating online meetings |
| `Team.Create` | Application | Creating teams |
| `Team.ReadBasic.All` | Application | Reading team metadata |
| `User.Read.All` | Application | Reading user info |
| `User.ReadBasic.All` | Application | Basic user info |

Click **Grant admin consent** → **Yes**

#### Delegated permissions (for device-code auth):

- `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Send`, `ChannelMessage.Read.All`, `Member.Read.All`, `OnlineMeetings.ReadWrite.All`, `Calendars.ReadWrite`, `User.Read.All`

### 3. Create a Client Secret

**Certificates & secrets** → **Client secrets** → **New client secret** → Copy the **Value**

## Configuration

### Via environment variables

```env
TEAMS_TENANT_ID=your-tenant-id
TEAMS_CLIENT_ID=your-client-id
TEAMS_CLIENT_SECRET=your-client-secret
TEAMS_AUTH_MODE=client-credentials
```

For device-code auth (interactive):

```env
TEAMS_AUTH_MODE=device-code
# No TEAMS_CLIENT_SECRET needed
```

For direct token (bypasses OAuth flows):

```env
TEAMS_AUTH_TOKEN=eyJ...
# TEAMS_TENANT_ID and TEAMS_CLIENT_ID not required
```

Read-only mode (reduced scopes):

```env
TEAMS_READ_ONLY=true
```

### MCP Client Config (opencode.json)

```json
{
  "mcpServers": {
    "microsoft-teams": {
      "command": "npx",
      "args": ["-y", "@anupam365/ms-teams-mcp"],
      "env": {
        "TEAMS_TENANT_ID": "...",
        "TEAMS_CLIENT_ID": "...",
        "TEAMS_CLIENT_SECRET": "...",
        "TEAMS_AUTH_MODE": "client-credentials"
      }
    }
  }
}
```

## Run

```bash
npx -y @anupam365/ms-teams-mcp
```

## Tools

### Teams (8)

| Tool | Description |
|---|---|
| `list_teams` | List all teams |
| `get_team` | Get team details |
| `create_team` | Create a new team |
| `update_team` | Update team metadata |
| `delete_team` | Delete a team |
| `archive_team` | Archive a team |
| `unarchive_team` | Restore an archived team |
| `clone_team` | Clone an existing team |

### Channels (7)

| Tool | Description |
|---|---|
| `list_channels` | List channels in a team |
| `get_channel` | Get channel details |
| `create_channel` | Create a new channel |
| `update_channel` | Update channel settings |
| `delete_channel` | Delete a channel |
| `archive_channel` | Archive a channel |
| `unarchive_channel` | Restore an archived channel |

### Channel Members (3)

| Tool | Description |
|---|---|
| `list_channel_members` | List channel members |
| `add_channel_member` | Add a user to a channel |
| `remove_channel_member` | Remove a user from a channel |

### Channel Messages (5)

| Tool | Description |
|---|---|
| `send_channel_message` | Send a message (supports format, importance, @mentions) |
| `list_channel_messages` | List recent channel messages |
| `reply_to_channel_message` | Reply to a channel message |
| `update_channel_message` | Edit an existing channel message |
| `delete_channel_message` | Delete a channel message |

### Chats (8)

| Tool | Description |
|---|---|
| `list_chats` | List 1-on-1 and group chats |
| `get_chat` | Get chat details |
| `create_chat` | Create a new 1-on-1 or group chat |
| `send_chat_message` | Send a chat message (supports format, importance) |
| `list_chat_messages` | List chat messages |
| `list_chat_members` | List chat members |
| `add_chat_member` | Add a user to a group chat |
| `list_chat_tabs` | List tabs in a chat |

### Chat Messages (2)

| Tool | Description |
|---|---|
| `update_chat_message` | Edit an existing chat message |
| `delete_chat_message` | Delete a chat message |

### Apps & Tabs (7)

| Tool | Description |
|---|---|
| `list_available_apps` | List tenant app catalog |
| `list_team_installed_apps` | List apps installed in a team |
| `install_app_in_team` | Install an app in a team |
| `uninstall_app_from_team` | Remove an app from a team |
| `list_channel_tabs` | List tabs in a channel |
| `add_channel_tab` | Add a tab to a channel |
| `remove_channel_tab` | Remove a tab from a channel |

### Planner (4)

| Tool | Description |
|---|---|
| `list_team_plans` | List plans in a team |
| `list_plan_buckets` | List buckets in a plan |
| `list_plan_tasks` | List tasks in a plan |
| `create_planner_task` | Create a new task |

### Calendar & Meetings (7)

| Tool | Description |
|---|---|
| `list_team_events` | List team calendar events |
| `create_event` | Create a team calendar event |
| `create_user_event` | Create event on a user's calendar |
| `get_event` | Get event details |
| `create_online_meeting` | Create a Teams meeting |
| `get_online_meeting` | Get meeting details |
| `list_online_meetings` | List upcoming meetings |

### Presence & Settings (4)

| Tool | Description |
|---|---|
| `get_user_presence` | Get user presence status |
| `get_team_settings` | Get team settings |
| `update_team_member_settings` | Update team settings |
| `list_team_members` | List all team members |

### Tags (3)

| Tool | Description |
|---|---|
| `list_team_tags` | List tags in a team |
| `create_tag` | Create a new tag |
| `delete_tag` | Delete a tag |

### Scheduling (6)

| Tool | Description |
|---|---|
| `list_scheduling_groups` | List scheduling groups |
| `list_shifts` | List shifts |
| `create_shift` | Create a shift |
| `list_time_off_reasons` | List time-off reasons |
| `list_time_off` | List time-off entries |
| `list_time_off_requests` | List time-off requests |

### Subscriptions (3)

| Tool | Description |
|---|---|
| `create_subscription` | Create a webhook subscription |
| `list_subscriptions` | List active subscriptions |
| `delete_subscription` | Delete a subscription |

### Search & Users (3)

| Tool | Description |
|---|---|
| `search_users` | Search users in the organization |
| `search_messages` | Search across all Teams messages |
| `get_my_mentions` | Get messages where you are @mentioned |

### Files (1)

| Tool | Description |
|---|---|
| `list_channel_files` | List files in a channel's folder |

### Auth Management (2)

| Tool | Description |
|---|---|
| `check_auth` | Check authentication status and mode |
| `logout` | Clear cached tokens and log out |

## Auth Modes

| Mode | Description | Best for |
|---|---|---|
| `token` | Direct AUTH_TOKEN env var | Testing, CI |
| `client-credentials` | App-only, no user interaction | Automated/headless use |
| `device-code` | Interactive browser login with persistent token cache | Personal/local use |

### Token Persistence

In device-code mode, the auth record is saved to `~/.microsoft-teams-mcp/auth-record.json`. On subsequent starts, the token is silently refreshed without re-prompting. If the refresh fails, the device-code prompt appears again.

### Read-Only Mode

Set `TEAMS_READ_ONLY=true` to restrict Graph scopes to `User.Read.All` and `Group.Read.All` only. Write operations will fail.

## Development

```bash
git clone <repo>
cd ms-teams-mcp
npm install
npm run dev   # watch mode
npm run build # compile
```

## License

MIT
