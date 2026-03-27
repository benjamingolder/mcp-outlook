import { Client } from "@microsoft/microsoft-graph-client";

export async function getMyPresence(client: Client) {
  const p = await client.api("/me/presence").get();
  return {
    id: p.id,
    availability: p.availability,
    activity: p.activity,
    statusMessage: p.statusMessage?.message?.content ?? null,
  };
}

export async function getUserPresence(client: Client, params: { userId: string }) {
  const { userId } = params;
  const p = await client.api(`/users/${userId}/presence`).get();
  return {
    id: p.id,
    availability: p.availability,
    activity: p.activity,
    statusMessage: p.statusMessage?.message?.content ?? null,
  };
}

export async function getPresenceForUsers(client: Client, params: { userIds: string[] }) {
  const { userIds } = params;
  const result = await client.api("/communications/getPresencesByUserId").post({ ids: userIds });
  return result.value.map((p: any) => ({
    id: p.id,
    availability: p.availability,
    activity: p.activity,
  }));
}

export async function setMyPresence(client: Client, params: {
  availability: "Available" | "Busy" | "DoNotDisturb" | "BeRightBack" | "Away" | "Offline";
  activity: string;
  expirationDuration?: string;
}) {
  const { availability, activity, expirationDuration = "PT1H" } = params;
  await client.api("/me/presence/setPresence").post({
    sessionId: "mcp-outlook",
    availability,
    activity,
    expirationDuration,
  });
  return { success: true, message: "Präsenz gesetzt." };
}
