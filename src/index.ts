import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { AuthManager, AuthConfig, getScopes } from "./auth.js";
import { TeamsClient } from "./teams.js";

function getConfig(): AuthConfig {
  const tenantId = process.env.TEAMS_TENANT_ID;
  const clientId = process.env.TEAMS_CLIENT_ID;
  const clientSecret = process.env.TEAMS_CLIENT_SECRET;
  const authToken = process.env.TEAMS_AUTH_TOKEN;
  const readOnly = process.env.TEAMS_READ_ONLY === "true";
  const authMode = (authToken ? "token" : process.env.TEAMS_AUTH_MODE ?? "device-code") as AuthConfig["authMode"];

  if (authMode !== "token") {
    if (!tenantId) throw new Error("TEAMS_TENANT_ID is required");
    if (!clientId) throw new Error("TEAMS_CLIENT_ID is required");
  }

  return { tenantId, clientId, clientSecret, authToken, authMode, readOnly };
}

const config = getConfig();
const auth = new AuthManager(config);
let teams!: TeamsClient;

const server = new Server(
  { name: "microsoft-teams-mcp", version: "1.0.0" },
  { capabilities: { tools: {} } }
);

const makeTool = (name: string, description: string, properties: Record<string, any> = {}, required: string[] = []) => ({
  name,
  description,
  inputSchema: { type: "object", properties, required },
});

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    makeTool("list_teams", "List all Microsoft Teams the authenticated user belongs to"),
    makeTool("get_team", "Get details of a specific team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("create_team", "Create a new Microsoft Team", {
      display_name: { type: "string", description: "Display name for the new team" },
      description: { type: "string", description: "Optional description" },
      visibility: { type: "string", enum: ["private", "public"], description: "Team visibility (default private)" },
    }, ["display_name"]),
    makeTool("update_team", "Update a team's settings or metadata", {
      team_id: { type: "string", description: "The ID of the team" },
      display_name: { type: "string", description: "New display name" },
      description: { type: "string", description: "New description" },
      visibility: { type: "string", enum: ["private", "public"], description: "New visibility" },
    }, ["team_id"]),
    makeTool("delete_team", "Delete a Microsoft Team permanently", {
      team_id: { type: "string", description: "The ID of the team to delete" },
    }, ["team_id"]),
    makeTool("archive_team", "Archive a team (makes it read-only)", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("unarchive_team", "Restore an archived team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("clone_team", "Clone an existing team", {
      team_id: { type: "string", description: "The ID of the source team" },
      display_name: { type: "string", description: "Display name for the cloned team" },
      description: { type: "string", description: "Optional description" },
      visibility: { type: "string", enum: ["private", "public"], description: "Visibility for cloned team" },
    }, ["team_id", "display_name"]),

    makeTool("list_channels", "List all channels in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("get_channel", "Get details of a specific channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
    }, ["team_id", "channel_id"]),
    makeTool("create_channel", "Create a new channel in a team", {
      team_id: { type: "string", description: "The ID of the team" },
      display_name: { type: "string", description: "Display name for the new channel" },
      description: { type: "string", description: "Optional description" },
    }, ["team_id", "display_name"]),
    makeTool("update_channel", "Update a channel's name or description", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      display_name: { type: "string", description: "New display name" },
      description: { type: "string", description: "New description" },
    }, ["team_id", "channel_id"]),
    makeTool("delete_channel", "Delete a channel permanently", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
    }, ["team_id", "channel_id"]),
    makeTool("archive_channel", "Archive a channel (read-only)", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
    }, ["team_id", "channel_id"]),
    makeTool("unarchive_channel", "Restore an archived channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
    }, ["team_id", "channel_id"]),

    makeTool("list_channel_members", "List all members of a channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
    }, ["team_id", "channel_id"]),
    makeTool("add_channel_member", "Add a user to a channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      user_id: { type: "string", description: "The user's object ID (from Entra ID)" },
      role: { type: "string", enum: ["member", "owner"], description: "Role (default member)" },
    }, ["team_id", "channel_id", "user_id"]),
    makeTool("remove_channel_member", "Remove a user from a channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      membership_id: { type: "string", description: "The membership ID (from list_channel_members)" },
    }, ["team_id", "channel_id", "membership_id"]),

    makeTool("send_channel_message", "Send a message to a channel with optional format, importance, and mentions", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      content: { type: "string", description: "Message content" },
      format: { type: "string", enum: ["text", "html"], description: "Content format (text or html)" },
      importance: { type: "string", enum: ["normal", "high", "urgent"], description: "Message importance level" },
      mentions: {
        type: "array",
        items: {
          type: "object",
          properties: {
            mention: { type: "string" },
            userId: { type: "string" },
          },
        },
        description: "Optional array of @mentions",
      },
    }, ["team_id", "channel_id", "content"]),
    makeTool("list_channel_messages", "List recent messages in a channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      top: { type: "number", description: "Number of messages to fetch (default 20)" },
    }, ["team_id", "channel_id"]),
    makeTool("reply_to_channel_message", "Reply to an existing channel message", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      parent_message_id: { type: "string", description: "The ID of the message to reply to" },
      content: { type: "string", description: "Reply content" },
    }, ["team_id", "channel_id", "parent_message_id", "content"]),

    makeTool("list_chats", "List all 1-on-1 and group chats"),
    makeTool("get_chat", "Get details of a specific chat", {
      chat_id: { type: "string", description: "The ID of the chat" },
    }, ["chat_id"]),
    makeTool("send_chat_message", "Send a message to a 1-on-1 or group chat with optional format and importance", {
      chat_id: { type: "string", description: "The ID of the chat" },
      content: { type: "string", description: "Message content" },
      format: { type: "string", enum: ["text", "html"], description: "Content format (text or html)" },
      importance: { type: "string", enum: ["normal", "high", "urgent"], description: "Message importance level" },
    }, ["chat_id", "content"]),
    makeTool("list_chat_messages", "List recent messages in a chat", {
      chat_id: { type: "string", description: "The ID of the chat" },
      top: { type: "number", description: "Number of messages to fetch (default 20)" },
    }, ["chat_id"]),
    makeTool("list_chat_members", "List members of a 1-on-1 or group chat", {
      chat_id: { type: "string", description: "The ID of the chat" },
    }, ["chat_id"]),
    makeTool("add_chat_member", "Add a user to a group chat", {
      chat_id: { type: "string", description: "The ID of the chat" },
      user_id: { type: "string", description: "The user's object ID (from Entra ID)" },
    }, ["chat_id", "user_id"]),
    makeTool("list_chat_tabs", "List tabs in a chat", {
      chat_id: { type: "string", description: "The ID of the chat" },
    }, ["chat_id"]),

    makeTool("list_available_apps", "List all Teams apps available in the tenant catalog"),
    makeTool("list_team_installed_apps", "List apps installed in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("install_app_in_team", "Install a Teams app in a team", {
      team_id: { type: "string", description: "The ID of the team" },
      app_catalog_id: { type: "string", description: "The app's catalog ID (from list_available_apps)" },
    }, ["team_id", "app_catalog_id"]),
    makeTool("uninstall_app_from_team", "Remove an installed app from a team", {
      team_id: { type: "string", description: "The ID of the team" },
      app_installation_id: { type: "string", description: "The app installation ID (from list_team_installed_apps)" },
    }, ["team_id", "app_installation_id"]),

    makeTool("list_channel_tabs", "List tabs in a channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
    }, ["team_id", "channel_id"]),
    makeTool("add_channel_tab", "Add a tab to a channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      display_name: { type: "string", description: "Display name for the tab" },
      teams_app_id: { type: "string", description: "The Teams app ID (from list_available_apps)" },
      entity_url: { type: "string", description: "Optional URL for the tab's content" },
    }, ["team_id", "channel_id", "display_name", "teams_app_id"]),
    makeTool("remove_channel_tab", "Remove a tab from a channel", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      tab_id: { type: "string", description: "The ID of the tab" },
    }, ["team_id", "channel_id", "tab_id"]),

    makeTool("list_team_plans", "List all Planner plans in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("list_plan_buckets", "List all buckets in a Planner plan", {
      plan_id: { type: "string", description: "The ID of the plan" },
    }, ["plan_id"]),
    makeTool("list_plan_tasks", "List tasks in a Planner plan, optionally filtered by bucket", {
      plan_id: { type: "string", description: "The ID of the plan" },
      bucket_id: { type: "string", description: "Optional bucket ID to filter tasks" },
    }, ["plan_id"]),
    makeTool("create_planner_task", "Create a new task in a Planner plan", {
      plan_id: { type: "string", description: "The ID of the plan" },
      bucket_id: { type: "string", description: "The ID of the bucket" },
      title: { type: "string", description: "Task title" },
      due_date_time: { type: "string", description: "Optional due date/time (ISO 8601)" },
    }, ["plan_id", "bucket_id", "title"]),

    makeTool("list_team_events", "List calendar events in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("create_event", "Create a calendar event in a team", {
      team_id: { type: "string", description: "The ID of the team" },
      subject: { type: "string", description: "Event subject/title" },
      start_date_time: { type: "string", description: "Start time (ISO 8601, e.g. 2026-06-23T14:00:00Z)" },
      end_date_time: { type: "string", description: "End time (ISO 8601)" },
      description: { type: "string", description: "Optional event description" },
    }, ["team_id", "subject", "start_date_time", "end_date_time"]),
    makeTool("create_user_event", "Create a calendar event on a user's personal calendar", {
      user_id: { type: "string", description: "The user's object ID or UPN" },
      subject: { type: "string", description: "Event subject/title" },
      start_date_time: { type: "string", description: "Start time (ISO 8601, e.g. 2026-06-23T14:00:00Z)" },
      end_date_time: { type: "string", description: "End time (ISO 8601)" },
      description: { type: "string", description: "Optional event description" },
    }, ["user_id", "subject", "start_date_time", "end_date_time"]),
    makeTool("get_event", "Get a specific calendar event", {
      team_id: { type: "string", description: "The ID of the team" },
      event_id: { type: "string", description: "The ID of the event" },
    }, ["team_id", "event_id"]),

    makeTool("create_online_meeting", "Create an online meeting", {
      subject: { type: "string", description: "Meeting subject" },
      start_date_time: { type: "string", description: "Start time (ISO 8601)" },
      end_date_time: { type: "string", description: "End time (ISO 8601)" },
      participants: {
        type: "array",
        items: { type: "string" },
        description: "Optional array of attendee UPNs",
      },
      organizer_id: {
        type: "string",
        description: "Organizer's UPN or object ID (required for client-credentials auth)",
      },
    }, ["subject", "start_date_time", "end_date_time"]),
    makeTool("get_online_meeting", "Get details of an online meeting", {
      meeting_id: { type: "string", description: "The ID of the meeting" },
    }, ["meeting_id"]),
    makeTool("list_online_meetings", "List upcoming online meetings"),

    makeTool("get_user_presence", "Get a user's presence status (Available, Busy, etc.)", {
      user_id: { type: "string", description: "The user's object ID or UPN" },
    }, ["user_id"]),

    makeTool("get_team_settings", "Get a team's member, messaging, and fun settings", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("update_team_member_settings", "Update a team's member/messaging/fun settings", {
      team_id: { type: "string", description: "The ID of the team" },
      member_settings: {
        type: "object",
        description: "Member settings (e.g. allowCreateUpdateChannels, allowDeleteChannels)",
        properties: {
          allowCreateUpdateChannels: { type: "boolean" },
          allowDeleteChannels: { type: "boolean" },
          allowAddRemoveApps: { type: "boolean" },
          allowCreateUpdateRemoveTabs: { type: "boolean" },
          allowCreateUpdateRemoveConnectors: { type: "boolean" },
        },
      },
      messaging_settings: {
        type: "object",
        description: "Messaging settings (e.g. allowUserEditMessages)",
        properties: {
          allowUserEditMessages: { type: "boolean" },
          allowUserDeleteMessages: { type: "boolean" },
          allowOwnerDeleteMessages: { type: "boolean" },
          allowTeamMentions: { type: "boolean" },
          allowChannelMentions: { type: "boolean" },
        },
      },
    }, ["team_id"]),

    makeTool("list_team_tags", "List tags in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("create_tag", "Create a tag in a team", {
      team_id: { type: "string", description: "The ID of the team" },
      display_name: { type: "string", description: "Tag display name" },
    }, ["team_id", "display_name"]),
    makeTool("delete_tag", "Delete a tag from a team", {
      team_id: { type: "string", description: "The ID of the team" },
      tag_id: { type: "string", description: "The ID of the tag" },
    }, ["team_id", "tag_id"]),

    makeTool("list_scheduling_groups", "List scheduling groups in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("list_shifts", "List shifts in a team, optionally by scheduling group", {
      team_id: { type: "string", description: "The ID of the team" },
      group_id: { type: "string", description: "Optional scheduling group ID to filter" },
    }, ["team_id"]),
    makeTool("create_shift", "Create a shift for a user", {
      team_id: { type: "string", description: "The ID of the team" },
      user_id: { type: "string", description: "The user's object ID" },
      start_date_time: { type: "string", description: "Shift start time (ISO 8601)" },
      end_date_time: { type: "string", description: "Shift end time (ISO 8601)" },
      scheduling_group_id: { type: "string", description: "Optional scheduling group ID" },
    }, ["team_id", "user_id", "start_date_time", "end_date_time"]),
    makeTool("list_time_off_reasons", "List time-off reasons set up in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("list_time_off", "List time-off entries in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),
    makeTool("list_time_off_requests", "List time-off requests in a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),

    makeTool("create_subscription", "Create a webhook subscription for change notifications", {
      resource: { type: "string", description: "Resource URL to monitor (e.g. /teams/{id}/channels)" },
      change_type: { type: "string", description: "Type of change (created, updated, deleted)" },
      notification_url: { type: "string", description: "URL where notifications are sent" },
      expiration_date_time: { type: "string", description: "Optional expiration (ISO 8601, max 3 days)" },
    }, ["resource", "change_type", "notification_url"]),
    makeTool("list_subscriptions", "List all active webhook subscriptions"),
    makeTool("delete_subscription", "Delete a webhook subscription", {
      subscription_id: { type: "string", description: "The ID of the subscription" },
    }, ["subscription_id"]),

    makeTool("list_channel_files", "List files in a channel's files folder", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
    }, ["team_id", "channel_id"]),
    makeTool("list_team_members", "List all members of a team", {
      team_id: { type: "string", description: "The ID of the team" },
    }, ["team_id"]),

    makeTool("search_users", "Search for users in the organization", {
      query: { type: "string", description: "Search query (matches displayName or email)" },
    }, ["query"]),
    makeTool("create_chat", "Create a new 1-on-1 or group chat", {
      chat_type: { type: "string", enum: ["oneOnOne", "group"], description: "Type of chat" },
      user_ids: {
        type: "array", items: { type: "string" },
        description: "Array of user object IDs to add to the chat",
      },
      topic: { type: "string", description: "Optional topic for group chats" },
    }, ["chat_type", "user_ids"]),
    makeTool("update_channel_message", "Edit an existing channel message", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      message_id: { type: "string", description: "The ID of the message to update" },
      content: { type: "string", description: "New message content" },
    }, ["team_id", "channel_id", "message_id", "content"]),
    makeTool("delete_channel_message", "Delete a channel message", {
      team_id: { type: "string", description: "The ID of the team" },
      channel_id: { type: "string", description: "The ID of the channel" },
      message_id: { type: "string", description: "The ID of the message to delete" },
    }, ["team_id", "channel_id", "message_id"]),
    makeTool("update_chat_message", "Edit an existing chat message", {
      chat_id: { type: "string", description: "The ID of the chat" },
      message_id: { type: "string", description: "The ID of the message to update" },
      content: { type: "string", description: "New message content" },
    }, ["chat_id", "message_id", "content"]),
    makeTool("delete_chat_message", "Delete a chat message", {
      chat_id: { type: "string", description: "The ID of the chat" },
      message_id: { type: "string", description: "The ID of the message to delete" },
    }, ["chat_id", "message_id"]),
    makeTool("search_messages", "Search across all Teams messages using Microsoft Search", {
      query: { type: "string", description: "Search query string" },
    }, ["query"]),
    makeTool("get_my_mentions", "Get messages where the authenticated user is @mentioned"),
    makeTool("check_auth", "Check the current authentication status and show auth mode"),
    makeTool("logout", "Log out and clear cached authentication tokens"),
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    const a = args as Record<string, any>;

    switch (name) {
      case "list_teams": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTeams(), null, 2) }] };
      }
      case "get_team": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getTeam(a.team_id), null, 2) }] };
      }
      case "create_team": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createTeam(a.display_name, a.description, a.visibility), null, 2) }] };
      }
      case "update_team": {
        const patch: Record<string, any> = {};
        if (a.display_name) patch.displayName = a.display_name;
        if (a.description) patch.description = a.description;
        if (a.visibility) patch.visibility = a.visibility;
        return { content: [{ type: "text", text: JSON.stringify(await teams.updateTeam(a.team_id, patch), null, 2) }] };
      }
      case "delete_team": {
        await teams.deleteTeam(a.team_id);
        return { content: [{ type: "text", text: "Team deleted successfully" }] };
      }
      case "archive_team": {
        await teams.archiveTeam(a.team_id);
        return { content: [{ type: "text", text: "Team archived successfully" }] };
      }
      case "unarchive_team": {
        await teams.unarchiveTeam(a.team_id);
        return { content: [{ type: "text", text: "Team unarchived successfully" }] };
      }
      case "clone_team": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.cloneTeam(a.team_id, a.display_name, a.description, a.visibility), null, 2) }] };
      }

      case "list_channels": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChannels(a.team_id), null, 2) }] };
      }
      case "get_channel": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getChannel(a.team_id, a.channel_id), null, 2) }] };
      }
      case "create_channel": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createChannel(a.team_id, a.display_name, a.description), null, 2) }] };
      }
      case "update_channel": {
        const cp: Record<string, any> = {};
        if (a.display_name) cp.displayName = a.display_name;
        if (a.description) cp.description = a.description;
        return { content: [{ type: "text", text: JSON.stringify(await teams.updateChannel(a.team_id, a.channel_id, cp), null, 2) }] };
      }
      case "delete_channel": {
        await teams.deleteChannel(a.team_id, a.channel_id);
        return { content: [{ type: "text", text: "Channel deleted successfully" }] };
      }
      case "archive_channel": {
        await teams.archiveChannel(a.team_id, a.channel_id);
        return { content: [{ type: "text", text: "Channel archived successfully" }] };
      }
      case "unarchive_channel": {
        await teams.unarchiveChannel(a.team_id, a.channel_id);
        return { content: [{ type: "text", text: "Channel unarchived successfully" }] };
      }

      case "list_channel_members": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChannelMembers(a.team_id, a.channel_id), null, 2) }] };
      }
      case "add_channel_member": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.addChannelMember(a.team_id, a.channel_id, a.user_id, a.role), null, 2) }] };
      }
      case "remove_channel_member": {
        await teams.removeChannelMember(a.team_id, a.channel_id, a.membership_id);
        return { content: [{ type: "text", text: "Member removed successfully" }] };
      }

      case "send_channel_message": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.sendChannelMessage(a.team_id, a.channel_id, a.content, a.format, a.importance, a.mentions), null, 2) }] };
      }
      case "list_channel_messages": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChannelMessages(a.team_id, a.channel_id, a.top), null, 2) }] };
      }
      case "reply_to_channel_message": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.replyToChannelMessage(a.team_id, a.channel_id, a.parent_message_id, a.content), null, 2) }] };
      }

      case "list_chats": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChats(), null, 2) }] };
      }
      case "get_chat": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getChat(a.chat_id), null, 2) }] };
      }
      case "send_chat_message": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.sendChatMessage(a.chat_id, a.content, a.format, a.importance), null, 2) }] };
      }
      case "list_chat_messages": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChatMessages(a.chat_id, a.top), null, 2) }] };
      }
      case "list_chat_members": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChatMembers(a.chat_id), null, 2) }] };
      }
      case "add_chat_member": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.addChatMember(a.chat_id, a.user_id), null, 2) }] };
      }
      case "list_chat_tabs": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChatTabs(a.chat_id), null, 2) }] };
      }

      case "list_available_apps": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listAvailableApps(), null, 2) }] };
      }
      case "list_team_installed_apps": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTeamInstalledApps(a.team_id), null, 2) }] };
      }
      case "install_app_in_team": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.installAppInTeam(a.team_id, a.app_catalog_id), null, 2) }] };
      }
      case "uninstall_app_from_team": {
        await teams.uninstallAppFromTeam(a.team_id, a.app_installation_id);
        return { content: [{ type: "text", text: "App uninstalled successfully" }] };
      }

      case "list_channel_tabs": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChannelTabs(a.team_id, a.channel_id), null, 2) }] };
      }
      case "add_channel_tab": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.addChannelTab(a.team_id, a.channel_id, a.display_name, a.teams_app_id, a.entity_url), null, 2) }] };
      }
      case "remove_channel_tab": {
        await teams.removeChannelTab(a.team_id, a.channel_id, a.tab_id);
        return { content: [{ type: "text", text: "Tab removed successfully" }] };
      }

      case "list_team_plans": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTeamPlans(a.team_id), null, 2) }] };
      }
      case "list_plan_buckets": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listPlanBuckets(a.plan_id), null, 2) }] };
      }
      case "list_plan_tasks": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listPlanTasks(a.plan_id, a.bucket_id), null, 2) }] };
      }
      case "create_planner_task": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createPlannerTask(a.plan_id, a.bucket_id, a.title, a.due_date_time), null, 2) }] };
      }

      case "list_team_events": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTeamEvents(a.team_id), null, 2) }] };
      }
      case "create_event": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createEvent(a.team_id, a.subject, a.start_date_time, a.end_date_time, a.description), null, 2) }] };
      }
      case "create_user_event": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createUserEvent(a.user_id, a.subject, a.start_date_time, a.end_date_time, a.description), null, 2) }] };
      }
      case "get_event": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getEvent(a.team_id, a.event_id), null, 2) }] };
      }

      case "create_online_meeting": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createOnlineMeeting(a.subject, a.start_date_time, a.end_date_time, a.participants, a.organizer_id), null, 2) }] };
      }
      case "get_online_meeting": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getOnlineMeeting(a.meeting_id), null, 2) }] };
      }
      case "list_online_meetings": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listOnlineMeetings(), null, 2) }] };
      }

      case "get_user_presence": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getUserPresence(a.user_id), null, 2) }] };
      }

      case "get_team_settings": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getTeamSettings(a.team_id), null, 2) }] };
      }
      case "update_team_member_settings": {
        const sp: Record<string, any> = {};
        if (a.member_settings) sp.memberSettings = a.member_settings;
        if (a.messaging_settings) sp.messagingSettings = a.messaging_settings;
        return { content: [{ type: "text", text: JSON.stringify(await teams.updateTeamMemberSettings(a.team_id, sp), null, 2) }] };
      }

      case "list_team_tags": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTeamTags(a.team_id), null, 2) }] };
      }
      case "create_tag": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createTag(a.team_id, a.display_name), null, 2) }] };
      }
      case "delete_tag": {
        await teams.deleteTag(a.team_id, a.tag_id);
        return { content: [{ type: "text", text: "Tag deleted successfully" }] };
      }

      case "list_scheduling_groups": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listSchedulingGroups(a.team_id), null, 2) }] };
      }
      case "list_shifts": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listShifts(a.team_id, a.group_id), null, 2) }] };
      }
      case "create_shift": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createShift(a.team_id, a.user_id, a.start_date_time, a.end_date_time, a.scheduling_group_id), null, 2) }] };
      }
      case "list_time_off_reasons": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTimeOffReasons(a.team_id), null, 2) }] };
      }
      case "list_time_off": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTimeOff(a.team_id), null, 2) }] };
      }
      case "list_time_off_requests": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTimeOffRequests(a.team_id), null, 2) }] };
      }

      case "create_subscription": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createSubscription(a.resource, a.change_type, a.notification_url, a.expiration_date_time), null, 2) }] };
      }
      case "list_subscriptions": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listSubscriptions(), null, 2) }] };
      }
      case "delete_subscription": {
        await teams.deleteSubscription(a.subscription_id);
        return { content: [{ type: "text", text: "Subscription deleted successfully" }] };
      }

      case "list_channel_files": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listChannelFilesRoot(a.team_id, a.channel_id), null, 2) }] };
      }
      case "list_team_members": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.listTeamMembers(a.team_id), null, 2) }] };
      }

      case "search_users": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.searchUsers(a.query), null, 2) }] };
      }
      case "create_chat": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.createChat(a.chat_type, a.user_ids, a.topic), null, 2) }] };
      }
      case "update_channel_message": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.updateChannelMessage(a.team_id, a.channel_id, a.message_id, a.content), null, 2) }] };
      }
      case "delete_channel_message": {
        await teams.deleteChannelMessage(a.team_id, a.channel_id, a.message_id);
        return { content: [{ type: "text", text: "Channel message deleted successfully" }] };
      }
      case "update_chat_message": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.updateChatMessage(a.chat_id, a.message_id, a.content), null, 2) }] };
      }
      case "delete_chat_message": {
        await teams.deleteChatMessage(a.chat_id, a.message_id);
        return { content: [{ type: "text", text: "Chat message deleted successfully" }] };
      }
      case "search_messages": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.searchMessages(a.query), null, 2) }] };
      }
      case "get_my_mentions": {
        return { content: [{ type: "text", text: JSON.stringify(await teams.getMyMentions(), null, 2) }] };
      }

      case "check_auth": {
        const ok = await auth.isAuthenticated();
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              authenticated: ok,
              authMode: auth.getAuthMode(),
              readOnly: auth.isReadOnly(),
            }, null, 2),
          }],
        };
      }
      case "logout": {
        await auth.logout();
        return { content: [{ type: "text", text: "Logged out successfully. Cached tokens cleared." }] };
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return { content: [{ type: "text", text: `Error: ${message}` }], isError: true };
  }
});

async function main() {
  await auth.ensureAuthenticated();
  const credential = auth.getCredential();
  const scopes = getScopes(auth.isReadOnly());
  teams = new TeamsClient(credential, auth.getAuthMode(), scopes);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Microsoft Teams MCP server running on stdio");
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
