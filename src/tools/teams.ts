import { Client } from "@microsoft/microsoft-graph-client";

export async function listTeams(client: Client, params: { top?: number }) {
  const { top = 20 } = params;
  const result = await client.api("/me/joinedTeams").top(top).get();
  return result.value.map((t: any) => ({
    id: t.id,
    displayName: t.displayName,
    description: t.description,
    isArchived: t.isArchived,
  }));
}

export async function listChannels(client: Client, params: { teamId: string }) {
  const { teamId } = params;
  const result = await client.api(`/teams/${teamId}/channels`).get();
  return result.value.map((c: any) => ({
    id: c.id,
    displayName: c.displayName,
    description: c.description,
    membershipType: c.membershipType,
    webUrl: c.webUrl,
  }));
}

export async function listChannelMessages(client: Client, params: {
  teamId: string;
  channelId: string;
  top?: number;
}) {
  const { teamId, channelId, top = 20 } = params;
  const result = await client
    .api(`/teams/${teamId}/channels/${channelId}/messages`)
    .top(top)
    .get();
  return result.value.map((m: any) => ({
    id: m.id,
    createdDateTime: m.createdDateTime,
    from: m.from?.user?.displayName ?? m.from?.application?.displayName ?? null,
    subject: m.subject,
    body: m.body?.content,
    bodyType: m.body?.contentType,
    importance: m.importance,
    webUrl: m.webUrl,
  }));
}

export async function sendChannelMessage(client: Client, params: {
  teamId: string;
  channelId: string;
  content: string;
  contentType?: "text" | "html";
  subject?: string;
}) {
  const { teamId, channelId, content, contentType = "text", subject } = params;
  const body: Record<string, unknown> = {
    body: { contentType, content },
  };
  if (subject) body.subject = subject;
  const result = await client
    .api(`/teams/${teamId}/channels/${channelId}/messages`)
    .post(body);
  return { id: result.id, webUrl: result.webUrl };
}

export async function listChats(client: Client, params: { top?: number }) {
  const { top = 20 } = params;
  const result = await client
    .api("/me/chats")
    .expand("members")
    .top(top)
    .get();
  return result.value.map((c: any) => ({
    id: c.id,
    chatType: c.chatType,
    topic: c.topic ?? null,
    lastUpdatedDateTime: c.lastUpdatedDateTime,
    members: c.members?.map((m: any) => m.displayName) ?? [],
  }));
}

export async function listChatMessages(client: Client, params: { chatId: string; top?: number }) {
  const { chatId, top = 20 } = params;
  const result = await client.api(`/me/chats/${chatId}/messages`).top(top).get();
  return result.value.map((m: any) => ({
    id: m.id,
    createdDateTime: m.createdDateTime,
    from: m.from?.user?.displayName ?? null,
    body: m.body?.content,
    bodyType: m.body?.contentType,
  }));
}

export async function sendChatMessage(client: Client, params: {
  chatId: string;
  content: string;
  contentType?: "text" | "html";
}) {
  const { chatId, content, contentType = "text" } = params;
  const result = await client.api(`/me/chats/${chatId}/messages`).post({
    body: { contentType, content },
  });
  return { id: result.id };
}
