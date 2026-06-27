import { TokenCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import "isomorphic-fetch";

function escapeOData(str: string): string {
  return str.replace(/'/g, "''");
}

function createClient(credential: TokenCredential, scopes: string[]): Client {
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes,
  });
  return Client.initWithMiddleware({ authProvider });
}

export interface TeamSummary {
  id: string;
  displayName: string;
  description?: string;
  visibility?: string;
}

export interface ChannelSummary {
  id: string;
  displayName: string;
  description?: string;
  membershipType?: string;
}

export interface MessageSummary {
  id: string;
  subject: string;
  body: string;
  from: string;
  createdDateTime: string;
}

export interface MemberSummary {
  id: string;
  displayName: string;
  email: string;
  userPrincipalName: string;
  userId?: string;
}

export interface UserSummary {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
}

export interface ChatMessageResult {
  id: string;
  chatId?: string;
  channelId?: string;
  teamId?: string;
  subject: string;
  body: string;
  from: string;
  createdDateTime: string;
  importance?: string;
  mentions?: string[];
}

export interface ChatSummary {
  id: string;
  topic?: string;
  chatType: string;
  createdDateTime: string;
}

export interface AppSummary {
  id: string;
  displayName: string;
  distributionMethod: string;
}

export interface TabSummary {
  id: string;
  displayName: string;
  teamsAppId?: string;
  entityUrl?: string;
}

export interface PlanSummary {
  id: string;
  title: string;
  createdDateTime: string;
}

export interface BucketSummary {
  id: string;
  name: string;
}

export interface TaskSummary {
  id: string;
  title: string;
  dueDateTime?: string;
  priority?: number;
}

export interface EventSummary {
  id: string;
  subject: string;
  start: string;
  end: string;
}

export interface MeetingSummary {
  id: string;
  subject: string;
  startDateTime: string;
  endDateTime: string;
  joinUrl: string;
}

export interface TagSummary {
  id: string;
  displayName: string;
}

export interface ShiftSummary {
  id: string;
  userId: string;
  startDateTime: string;
  endDateTime: string;
}

export interface SubscriptionSummary {
  id: string;
  resource: string;
  changeType: string;
  expirationDateTime: string;
}

export interface FileItem {
  id: string;
  name: string;
  size: number;
  webUrl: string;
  folder?: boolean;
}

export class TeamsClient {
  private client: Client;
  private credential: TokenCredential;
  private authMode: string;
  private scopes: string[];

  constructor(credential: TokenCredential, authMode: string = "device-code", scopes?: string[]) {
    this.credential = credential;
    this.scopes = scopes ?? ["https://graph.microsoft.com/.default"];
    this.client = createClient(credential, this.scopes);
    this.authMode = authMode;
  }

  private async getToken(): Promise<string> {
    const token = await this.credential.getToken(this.scopes);
    if (!token) throw new Error("Failed to acquire token");
    return token.token;
  }

  private async graphPost(
    url: string,
    body: Record<string, any>
  ): Promise<any> {
    const token = await this.getToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });
    return response.json();
  }

  private async graphPatch(
    url: string,
    body: Record<string, any>
  ): Promise<any> {
    const token = await this.getToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      const err = await response.text();
      throw new Error(`PATCH ${url} failed: ${response.status} ${err}`);
    }
    return response.json().catch(() => ({}));
  }

  private async graphDelete(url: string): Promise<void> {
    const token = await this.getToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok && response.status !== 204) {
      const err = await response.text();
      throw new Error(`DELETE ${url} failed: ${response.status} ${err}`);
    }
  }

  async listTeams(): Promise<TeamSummary[]> {
    let res: any;
    if (this.authMode === "client-credentials") {
      res = await this.client
        .api("/groups")
        .filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
        .select(["id", "displayName", "description", "visibility"])
        .get();
    } else {
      res = await this.client
        .api("/me/joinedTeams")
        .select(["id", "displayName", "description", "visibility"])
        .get();
    }
    return (
      res.value?.map((t: any) => ({
        id: t.id,
        displayName: t.displayName,
        description: t.description,
        visibility: t.visibility,
      })) ?? []
    );
  }

  async getTeam(teamId: string): Promise<TeamSummary> {
    const t = await this.client
      .api(`/teams/${teamId}`)
      .select(["id", "displayName", "description", "visibility"])
      .get();
    return {
      id: t.id,
      displayName: t.displayName,
      description: t.description,
      visibility: t.visibility,
    };
  }

  async createTeam(
    displayName: string,
    description?: string,
    visibility: "private" | "public" = "private"
  ): Promise<any> {
    const body: Record<string, any> = {
      displayName,
      visibility,
      template: "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
    };
    if (description) body.description = description;

    const token = await this.getToken();
    const response = await fetch("https://graph.microsoft.com/v1.0/teams", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    const result = await response.json();
    result._status = response.status;
    return result;
  }

  async updateTeam(
    teamId: string,
    patch: Record<string, any>
  ): Promise<any> {
    return this.graphPatch(`/teams/${teamId}`, patch);
  }

  async deleteTeam(teamId: string): Promise<void> {
    const t = await this.client.api(`/groups/${teamId}`).get();
    await this.graphDelete(`/groups/${teamId}`);
  }

  async archiveTeam(teamId: string): Promise<void> {
    await this.client.api(`/teams/${teamId}/archive`).post({});
  }

  async unarchiveTeam(teamId: string): Promise<void> {
    await this.client.api(`/teams/${teamId}/unarchive`).post({});
  }

  async cloneTeam(
    teamId: string,
    displayName: string,
    description?: string,
    visibility?: "private" | "public",
    partsToClone?: string
  ): Promise<any> {
    const body: Record<string, any> = { displayName };
    if (description) body.description = description;
    if (visibility) body.visibility = visibility;
    body.partsToClone = partsToClone ?? "apps,tabs,settings,channels,members";

    const token = await this.getToken();
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/clone`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      }
    );

    return {
      _status: response.status,
      _location: response.headers.get("Location"),
    };
  }

  async listChannels(teamId: string): Promise<ChannelSummary[]> {
    const res = await this.client
      .api(`/teams/${teamId}/channels`)
      .select(["id", "displayName", "description", "membershipType"])
      .get();
    return (
      res.value?.map((c: any) => ({
        id: c.id,
        displayName: c.displayName,
        description: c.description,
        membershipType: c.membershipType,
      })) ?? []
    );
  }

  async getChannel(teamId: string, channelId: string): Promise<ChannelSummary> {
    const c = await this.client
      .api(`/teams/${teamId}/channels/${channelId}`)
      .select(["id", "displayName", "description", "membershipType"])
      .get();
    return {
      id: c.id,
      displayName: c.displayName,
      description: c.description,
      membershipType: c.membershipType,
    };
  }

  async createChannel(
    teamId: string,
    displayName: string,
    description?: string
  ): Promise<ChannelSummary> {
    const body: Record<string, string> = { displayName };
    if (description) body.description = description;

    const res = await this.client
      .api(`/teams/${teamId}/channels`)
      .post(body);
    return {
      id: res.id,
      displayName: res.displayName,
      description: res.description,
    };
  }

  async updateChannel(
    teamId: string,
    channelId: string,
    patch: Record<string, any>
  ): Promise<any> {
    return this.graphPatch(`/teams/${teamId}/channels/${channelId}`, patch);
  }

  async deleteChannel(teamId: string, channelId: string): Promise<void> {
    await this.graphDelete(`/teams/${teamId}/channels/${channelId}`);
  }

  async archiveChannel(teamId: string, channelId: string): Promise<void> {
    await this.client
      .api(`/teams/${teamId}/channels/${channelId}/archive`)
      .post({});
  }

  async unarchiveChannel(teamId: string, channelId: string): Promise<void> {
    await this.client
      .api(`/teams/${teamId}/channels/${channelId}/unarchive`)
      .post({});
  }

  async listChannelMembers(
    teamId: string,
    channelId: string
  ): Promise<MemberSummary[]> {
    const res = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/members`)
      .get();
    return (
      res.value?.map((m: any) => ({
        id: m.id,
        displayName: m.displayName,
        email: m.email ?? "",
        userPrincipalName: m.userPrincipalName ?? "",
        userId: m.userId,
      })) ?? []
    );
  }

  async addChannelMember(
    teamId: string,
    channelId: string,
    userId: string,
    role: string = "member"
  ): Promise<any> {
    return this.client.api(`/teams/${teamId}/channels/${channelId}/members`).post(
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: [role],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`,
      }
    );
  }

  async removeChannelMember(
    teamId: string,
    channelId: string,
    membershipId: string
  ): Promise<void> {
    await this.graphDelete(
      `/teams/${teamId}/channels/${channelId}/members/${membershipId}`
    );
  }

  async sendChannelMessage(
    teamId: string,
    channelId: string,
    content: string,
    format?: string,
    importance?: string,
    mentions?: Array<{ mention: string; userId: string }>
  ): Promise<{ id: string }> {
    const body: Record<string, any> = {
      body: { content },
    };
    if (format === "html") {
      body.body.contentType = "html";
    }
    if (importance && ["normal", "high", "urgent"].includes(importance)) {
      body.importance = importance;
    }
    if (mentions?.length) {
      body.mentions = mentions.map((m, i) => ({
        id: i,
        mentionText: m.mention,
        mentioned: {
          "@odata.type": "#microsoft.graph.user",
          id: m.userId,
        },
      }));
    }
    const res = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/messages`)
      .post(body);
    return { id: res.id };
  }

  async listChannelMessages(
    teamId: string,
    channelId: string,
    top: number = 20
  ): Promise<MessageSummary[]> {
    const res = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/messages`)
      .top(top)
      .orderby("createdDateTime desc")
      .get();
    return (
      res.value?.map((m: any) => ({
        id: m.id,
        subject: m.subject ?? "",
        body: m.body?.content ?? "",
        from:
          m.from?.user?.displayName ??
          m.from?.application?.displayName ??
          "Unknown",
        createdDateTime: m.createdDateTime,
      })) ?? []
    );
  }

  async replyToChannelMessage(
    teamId: string,
    channelId: string,
    parentMessageId: string,
    content: string
  ): Promise<{ id: string }> {
    const res = await this.client
      .api(
        `/teams/${teamId}/channels/${channelId}/messages/${parentMessageId}/replies`
      )
      .post({ body: { content } });
    return { id: res.id };
  }

  async listChats(): Promise<ChatSummary[]> {
    const res = await this.client
      .api("/chats")
      .select(["id", "topic", "chatType", "createdDateTime"])
      .get();
    return (
      res.value?.map((c: any) => ({
        id: c.id,
        topic: c.topic,
        chatType: c.chatType,
        createdDateTime: c.createdDateTime,
      })) ?? []
    );
  }

  async getChat(chatId: string): Promise<ChatSummary> {
    const c = await this.client
      .api(`/chats/${chatId}`)
      .select(["id", "topic", "chatType", "createdDateTime"])
      .get();
    return {
      id: c.id,
      topic: c.topic,
      chatType: c.chatType,
      createdDateTime: c.createdDateTime,
    };
  }

  async sendChatMessage(
    chatId: string,
    content: string,
    format?: string,
    importance?: string
  ): Promise<{ id: string }> {
    const body: Record<string, any> = {
      body: { content },
    };
    if (format === "html") {
      body.body.contentType = "html";
    }
    if (importance && ["normal", "high", "urgent"].includes(importance)) {
      body.importance = importance;
    }
    const res = await this.client
      .api(`/chats/${chatId}/messages`)
      .post(body);
    return { id: res.id };
  }

  async listChatMessages(
    chatId: string,
    top: number = 20
  ): Promise<MessageSummary[]> {
    const res = await this.client
      .api(`/chats/${chatId}/messages`)
      .top(top)
      .orderby("createdDateTime desc")
      .get();
    return (
      res.value?.map((m: any) => ({
        id: m.id,
        subject: m.subject ?? "",
        body: m.body?.content ?? "",
        from:
          m.from?.user?.displayName ??
          m.from?.application?.displayName ??
          "Unknown",
        createdDateTime: m.createdDateTime,
      })) ?? []
    );
  }

  async listChatMembers(chatId: string): Promise<MemberSummary[]> {
    const res = await this.client.api(`/chats/${chatId}/members`).get();
    return (
      res.value?.map((m: any) => ({
        id: m.id,
        displayName: m.displayName,
        email: m.email ?? "",
        userPrincipalName: m.userPrincipalName ?? "",
        userId: m.userId,
      })) ?? []
    );
  }

  async addChatMember(chatId: string, userId: string): Promise<any> {
    return this.client.api(`/chats/${chatId}/members`).post({
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`,
    });
  }

  async listChatTabs(chatId: string): Promise<TabSummary[]> {
    const res = await this.client.api(`/chats/${chatId}/tabs`).get();
    return (
      res.value?.map((t: any) => ({
        id: t.id,
        displayName: t.displayName,
        teamsAppId: t.teamsApp?.id,
      })) ?? []
    );
  }

  async listAvailableApps(): Promise<AppSummary[]> {
    const res = await this.client.api("/appCatalogs/teamsApps").get();
    return (
      res.value?.map((a: any) => ({
        id: a.id,
        displayName: a.displayName,
        distributionMethod: a.distributionMethod,
      })) ?? []
    );
  }

  async listTeamInstalledApps(teamId: string): Promise<any[]> {
    const res = await this.client
      .api(`/teams/${teamId}/installedApps`)
      .expand("teamsApp")
      .get();
    return res.value ?? [];
  }

  async installAppInTeam(
    teamId: string,
    appCatalogId: string
  ): Promise<any> {
    return this.client.api(`/teams/${teamId}/installedApps`).post({
      "teamsApp@odata.bind": `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${appCatalogId}`,
    });
  }

  async uninstallAppFromTeam(
    teamId: string,
    appInstallationId: string
  ): Promise<void> {
    await this.graphDelete(
      `/teams/${teamId}/installedApps/${appInstallationId}`
    );
  }

  async listChannelTabs(
    teamId: string,
    channelId: string
  ): Promise<TabSummary[]> {
    const res = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/tabs`)
      .expand("teamsApp")
      .get();
    return (
      res.value?.map((t: any) => ({
        id: t.id,
        displayName: t.displayName,
        teamsAppId: t.teamsApp?.id,
      })) ?? []
    );
  }

  async addChannelTab(
    teamId: string,
    channelId: string,
    displayName: string,
    teamsAppId: string,
    entityUrl?: string
  ): Promise<any> {
    const body: Record<string, any> = {
      displayName,
      "teamsApp@odata.bind": `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${teamsAppId}`,
    };
    if (entityUrl) {
      body.configuration = { entityUrl };
    }
    return this.client
      .api(`/teams/${teamId}/channels/${channelId}/tabs`)
      .post(body);
  }

  async removeChannelTab(
    teamId: string,
    channelId: string,
    tabId: string
  ): Promise<void> {
    await this.graphDelete(
      `/teams/${teamId}/channels/${channelId}/tabs/${tabId}`
    );
  }

  async listTeamPlans(teamId: string): Promise<PlanSummary[]> {
    const res = await this.client
      .api(`/groups/${teamId}/planner/plans`)
      .get();
    return (
      res.value?.map((p: any) => ({
        id: p.id,
        title: p.title,
        createdDateTime: p.createdDateTime,
      })) ?? []
    );
  }

  async listPlanBuckets(planId: string): Promise<BucketSummary[]> {
    const res = await this.client
      .api(`/planner/plans/${planId}/buckets`)
      .get();
    return (
      res.value?.map((b: any) => ({
        id: b.id,
        name: b.name,
      })) ?? []
    );
  }

  async listPlanTasks(
    planId: string,
    bucketId?: string
  ): Promise<TaskSummary[]> {
    let req = this.client.api(`/planner/plans/${planId}/tasks`);
    if (bucketId) req = req.filter(`bucketId eq '${bucketId}'`);
    const res = await req.get();
    return (
      res.value?.map((t: any) => ({
        id: t.id,
        title: t.title,
        dueDateTime: t.dueDateTime,
        priority: t.priority,
      })) ?? []
    );
  }

  async createPlannerTask(
    planId: string,
    bucketId: string,
    title: string,
    dueDateTime?: string
  ): Promise<any> {
    const body: Record<string, any> = {
      planId,
      bucketId,
      title,
    };
    if (dueDateTime) body.dueDateTime = dueDateTime;

    const token = await this.getToken();
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/planner/tasks",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      }
    );
    return response.json();
  }

  async listTeamEvents(teamId: string): Promise<EventSummary[]> {
    const res = await this.client
      .api(`/groups/${teamId}/events`)
      .select(["id", "subject", "start", "end"])
      .get();
    return (
      res.value?.map((e: any) => ({
        id: e.id,
        subject: e.subject,
        start: e.start?.dateTime ?? "",
        end: e.end?.dateTime ?? "",
      })) ?? []
    );
  }

  async createEvent(
    teamId: string,
    subject: string,
    startDateTime: string,
    endDateTime: string,
    description?: string
  ): Promise<any> {
    const body: Record<string, any> = {
      subject,
      start: { dateTime: startDateTime, timeZone: "UTC" },
      end: { dateTime: endDateTime, timeZone: "UTC" },
    };
    if (description) body.body = { content: description, contentType: "text" };
    return this.client.api(`/groups/${teamId}/events`).post(body);
  }

  async createUserEvent(
    userId: string,
    subject: string,
    startDateTime: string,
    endDateTime: string,
    description?: string
  ): Promise<any> {
    const body: Record<string, any> = {
      subject,
      start: { dateTime: startDateTime, timeZone: "UTC" },
      end: { dateTime: endDateTime, timeZone: "UTC" },
    };
    if (description) body.body = { content: description, contentType: "text" };
    const token = await this.getToken();
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${userId}/calendar/events`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      }
    );
    return response.json();
  }

  async getEvent(
    teamId: string,
    eventId: string
  ): Promise<EventSummary> {
    const e = await this.client
      .api(`/groups/${teamId}/events/${eventId}`)
      .select(["id", "subject", "start", "end"])
      .get();
    return {
      id: e.id,
      subject: e.subject,
      start: e.start?.dateTime ?? "",
      end: e.end?.dateTime ?? "",
    };
  }

  async createOnlineMeeting(
    subject: string,
    startDateTime: string,
    endDateTime: string,
    participants?: string[],
    organizerId?: string
  ): Promise<MeetingSummary> {
    const body: Record<string, any> = {
      subject,
      startDateTime,
      endDateTime,
    };
    if (participants?.length) {
      body.participants = {
        attendees: participants.map((p) => ({
          upn: p,
          role: "attendee",
        })),
      };
    }

    const token = await this.getToken();
    const userId = organizerId ?? "me";
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      }
    );
    const m = await response.json();
    return {
      id: m.id,
      subject: m.subject,
      startDateTime: m.startDateTime,
      endDateTime: m.endDateTime,
      joinUrl: m.joinUrl ?? m.joinWebUrl ?? "",
    };
  }

  async getOnlineMeeting(meetingId: string): Promise<MeetingSummary> {
    const m = await this.client
      .api(`/users/me/onlineMeetings/${meetingId}`)
      .get();
    return {
      id: m.id,
      subject: m.subject,
      startDateTime: m.startDateTime,
      endDateTime: m.endDateTime,
      joinUrl: m.joinUrl ?? m.joinWebUrl ?? "",
    };
  }

  async listOnlineMeetings(): Promise<MeetingSummary[]> {
    const res = await this.client
      .api("/users/me/onlineMeetings")
      .filter("startDateTime ge now()")
      .get();
    return (
      res.value?.map((m: any) => ({
        id: m.id,
        subject: m.subject,
        startDateTime: m.startDateTime,
        endDateTime: m.endDateTime,
        joinUrl: m.joinUrl ?? m.joinWebUrl ?? "",
      })) ?? []
    );
  }

  async getUserPresence(userId: string): Promise<any> {
    return this.client.api(`/users/${userId}/presence`).get();
  }

  async getTeamSettings(teamId: string): Promise<any> {
    return this.client
      .api(`/teams/${teamId}`)
      .select([
        "memberSettings",
        "messagingSettings",
        "funSettings",
        "discoverySettings",
      ])
      .get();
  }

  async updateTeamMemberSettings(
    teamId: string,
    settings: Record<string, any>
  ): Promise<any> {
    return this.graphPatch(`/teams/${teamId}`, settings);
  }

  async listTeamTags(teamId: string): Promise<TagSummary[]> {
    const res = await this.client
      .api(`/teams/${teamId}/tags`)
      .get();
    return (
      res.value?.map((t: any) => ({
        id: t.id,
        displayName: t.displayName,
      })) ?? []
    );
  }

  async createTag(
    teamId: string,
    displayName: string
  ): Promise<TagSummary> {
    const res = await this.client
      .api(`/teams/${teamId}/tags`)
      .post({ displayName });
    return { id: res.id, displayName: res.displayName };
  }

  async deleteTag(teamId: string, tagId: string): Promise<void> {
    await this.graphDelete(`/teams/${teamId}/tags/${tagId}`);
  }

  async listSchedulingGroups(teamId: string): Promise<any[]> {
    const res = await this.client
      .api(`/teams/${teamId}/schedule/schedulingGroups`)
      .get();
    return res.value ?? [];
  }

  async listShifts(
    teamId: string,
    groupId?: string
  ): Promise<ShiftSummary[]> {
    let req = this.client.api(`/teams/${teamId}/schedule/shifts`);
    if (groupId) req = req.filter(`schedulingGroupId eq '${groupId}'`);
    const res = await req.get();
    return (
      res.value?.map((s: any) => ({
        id: s.id,
        userId: s.userId,
        startDateTime: s.startDateTime,
        endDateTime: s.endDateTime,
      })) ?? []
    );
  }

  async createShift(
    teamId: string,
    userId: string,
    startDateTime: string,
    endDateTime: string,
    schedulingGroupId?: string
  ): Promise<any> {
    const body: Record<string, any> = {
      userId,
      startDateTime,
      endDateTime,
      draftShift: {
        startDateTime,
        endDateTime,
        userId,
      },
    };
    if (schedulingGroupId) body.schedulingGroupId = schedulingGroupId;

    const token = await this.getToken();
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/schedule/shifts`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      }
    );
    return response.json();
  }

  async listTimeOffReasons(teamId: string): Promise<any[]> {
    const res = await this.client
      .api(`/teams/${teamId}/schedule/timeOffReasons`)
      .get();
    return res.value ?? [];
  }

  async listTimeOff(teamId: string): Promise<any[]> {
    const res = await this.client
      .api(`/teams/${teamId}/schedule/timeOff`)
      .get();
    return res.value ?? [];
  }

  async listTimeOffRequests(teamId: string): Promise<any[]> {
    const res = await this.client
      .api(`/teams/${teamId}/schedule/timeOffRequests`)
      .get();
    return res.value ?? [];
  }

  async createSubscription(
    resource: string,
    changeType: string,
    notificationUrl: string,
    expirationDateTime?: string
  ): Promise<SubscriptionSummary> {
    const body: Record<string, string> = {
      resource,
      changeType,
      notificationUrl,
    };
    if (expirationDateTime) {
      body.expirationDateTime = expirationDateTime;
    }

    const token = await this.getToken();
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/subscriptions",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      }
    );
    const s = await response.json();
    return {
      id: s.id,
      resource: s.resource,
      changeType: s.changeType,
      expirationDateTime: s.expirationDateTime,
    };
  }

  async listSubscriptions(): Promise<SubscriptionSummary[]> {
    const res = await this.client.api("/subscriptions").get();
    return (
      res.value?.map((s: any) => ({
        id: s.id,
        resource: s.resource,
        changeType: s.changeType,
        expirationDateTime: s.expirationDateTime,
      })) ?? []
    );
  }

  async deleteSubscription(subscriptionId: string): Promise<void> {
    await this.graphDelete(`/subscriptions/${subscriptionId}`);
  }

  async listChannelFilesRoot(
    teamId: string,
    channelId: string
  ): Promise<FileItem[]> {
    const drive = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/filesFolder`)
      .get();
    const res = await this.client
      .api(`/groups/${teamId}/drive/items/${drive.id}/children`)
      .select(["id", "name", "size", "webUrl", "folder"])
      .get();
    return (
      res.value?.map((f: any) => ({
        id: f.id,
        name: f.name,
        size: f.size ?? 0,
        webUrl: f.webUrl,
        folder: !!f.folder,
      })) ?? []
    );
  }

  async listTeamMembers(teamId: string): Promise<MemberSummary[]> {
    const res = await this.client.api(`/teams/${teamId}/members`).get();
    return (
      res.value?.map((m: any) => ({
        id: m.id,
        displayName: m.displayName,
        email: m.email ?? "",
        userPrincipalName: m.userPrincipalName ?? "",
        userId: m.userId,
      })) ?? []
    );
  }

  async searchUsers(query: string): Promise<UserSummary[]> {
    const res = await this.client
      .api("/users")
      .search(`"displayName:${escapeOData(query)}" OR "mail:${escapeOData(query)}"`)
      .select(["id", "displayName", "mail", "userPrincipalName", "jobTitle"])
      .header("ConsistencyLevel", "eventual")
      .top(25)
      .get();
    return (
      res.value?.map((u: any) => ({
        id: u.id,
        displayName: u.displayName,
        mail: u.mail ?? "",
        userPrincipalName: u.userPrincipalName,
        jobTitle: u.jobTitle,
      })) ?? []
    );
  }

  async createChat(
    chatType: string,
    userIds: string[],
    topic?: string
  ): Promise<ChatSummary> {
    const body: Record<string, any> = {
      chatType,
      members: userIds.map((uid) => ({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${uid}')`,
      })),
    };
    if (topic && chatType === "group") body.topic = topic;
    const res = await this.client.api("/chats").post(body);
    return {
      id: res.id,
      topic: res.topic,
      chatType: res.chatType,
      createdDateTime: res.createdDateTime,
    };
  }

  async updateChannelMessage(
    teamId: string,
    channelId: string,
    messageId: string,
    content: string
  ): Promise<any> {
    return this.graphPatch(
      `/teams/${teamId}/channels/${channelId}/messages/${messageId}`,
      { body: { content } }
    );
  }

  async deleteChannelMessage(
    teamId: string,
    channelId: string,
    messageId: string
  ): Promise<void> {
    await this.graphDelete(
      `/teams/${teamId}/channels/${channelId}/messages/${messageId}`
    );
  }

  async updateChatMessage(
    chatId: string,
    messageId: string,
    content: string
  ): Promise<any> {
    return this.graphPatch(
      `/chats/${chatId}/messages/${messageId}`,
      { body: { content } }
    );
  }

  async deleteChatMessage(
    chatId: string,
    messageId: string
  ): Promise<void> {
    await this.graphDelete(`/chats/${chatId}/messages/${messageId}`);
  }

  async searchMessages(query: string): Promise<ChatMessageResult[]> {
    const token = await this.getToken();
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/search/query",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          requests: [
            {
              entityTypes: ["chatMessage"],
              query: { queryString: query },
              from: 0,
              size: 25,
            },
          ],
        }),
      }
    );
    const data = await response.json();
    const hits = data.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    return hits.map((h: any) => {
      const r = h.resource;
      return {
        id: r.id,
        chatId: r.chatId,
        channelId: r.channelIdentity?.channelId,
        teamId: r.channelIdentity?.teamId,
        subject: r.subject ?? "",
        body: r.body?.content ?? "",
        from: r.from?.user?.displayName ?? "Unknown",
        createdDateTime: r.createdDateTime,
        importance: r.importance,
      };
    });
  }

  async getMyMentions(): Promise<ChatMessageResult[]> {
    const token = await this.getToken();
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/search/query",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          requests: [
            {
              entityTypes: ["chatMessage"],
              query: { queryString: "mentions:me" },
              from: 0,
              size: 25,
            },
          ],
        }),
      }
    );
    const data = await response.json();
    const hits = data.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    return hits.map((h: any) => {
      const r = h.resource;
      return {
        id: r.id,
        chatId: r.chatId,
        channelId: r.channelIdentity?.channelId,
        teamId: r.channelIdentity?.teamId,
        subject: r.subject ?? "",
        body: r.body?.content ?? "",
        from: r.from?.user?.displayName ?? "Unknown",
        createdDateTime: r.createdDateTime,
      };
    });
  }
}
