import { getGraphClient } from "../graph.js";

export async function getMyPresence() {
  const client = getGraphClient();
  const p = await client.api("/me/presence").get();
  return {
    id: p.id,
    availability: p.availability,
    activity: p.activity,
    statusMessage: p.statusMessage?.message?.content ?? null,
  };
}

export async function getUserPresence(params: { userId: string }) {
  const { userId } = params;
  const client = getGraphClient();
  const p = await client.api(`/users/${userId}/presence`).get();
  return {
    id: p.id,
    availability: p.availability,
    activity: p.activity,
    statusMessage: p.statusMessage?.message?.content ?? null,
  };
}

export async function getPresenceForUsers(params: { userIds: string[] }) {
  const { userIds } = params;
  const client = getGraphClient();
  const result = await client.api("/communications/getPresencesByUserId").post({ ids: userIds });
  return result.value.map((p: any) => ({
    id: p.id,
    availability: p.availability,
    activity: p.activity,
  }));
}

export async function setMyPresence(params: {
  availability: "Available" | "Busy" | "DoNotDisturb" | "BeRightBack" | "Away" | "Offline";
  activity: string;
  expirationDuration?: string;
}) {
  const { availability, activity, expirationDuration = "PT1H" } = params;
  const client = getGraphClient();
  await client.api("/me/presence/setPresence").post({
    sessionId: "mcp-outlook",
    availability,
    activity,
    expirationDuration,
  });
  return { success: true, message: "Präsenz gesetzt." };
}
