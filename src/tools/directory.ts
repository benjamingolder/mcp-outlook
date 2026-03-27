import { Client } from "@microsoft/microsoft-graph-client";

export async function listUsers(client: Client, params: { top?: number; filter?: string; search?: string }) {
  const { top = 20, filter, search } = params;

  let req = client
    .api("/users")
    .top(top)
    .select("id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,mobilePhone,accountEnabled");

  if (filter) req = req.filter(filter);
  if (search) req = req.header("ConsistencyLevel", "eventual").search(`"displayName:${search}"`);

  const result = await req.get();
  return result.value.map((u: any) => ({
    id: u.id,
    displayName: u.displayName,
    userPrincipalName: u.userPrincipalName,
    mail: u.mail,
    jobTitle: u.jobTitle,
    department: u.department,
    officeLocation: u.officeLocation,
    mobilePhone: u.mobilePhone,
    accountEnabled: u.accountEnabled,
  }));
}

export async function getUser(client: Client, params: { userId: string }) {
  const { userId } = params;
  const u = await client.api(`/users/${userId}`).get();
  return {
    id: u.id,
    displayName: u.displayName,
    userPrincipalName: u.userPrincipalName,
    mail: u.mail,
    jobTitle: u.jobTitle,
    department: u.department,
    officeLocation: u.officeLocation,
    mobilePhone: u.mobilePhone,
    businessPhones: u.businessPhones,
    accountEnabled: u.accountEnabled,
    createdDateTime: u.createdDateTime,
  };
}

export async function listGroups(client: Client, params: { top?: number; filter?: string; search?: string }) {
  const { top = 20, filter, search } = params;

  let req = client
    .api("/groups")
    .top(top)
    .select("id,displayName,description,mail,groupTypes,membershipRule,visibility");

  if (filter) req = req.filter(filter);
  if (search) req = req.header("ConsistencyLevel", "eventual").search(`"displayName:${search}"`);

  const result = await req.get();
  return result.value.map((g: any) => ({
    id: g.id,
    displayName: g.displayName,
    description: g.description,
    mail: g.mail,
    groupTypes: g.groupTypes,
    visibility: g.visibility,
  }));
}

export async function listGroupMembers(client: Client, params: { groupId: string; top?: number }) {
  const { groupId, top = 50 } = params;
  const result = await client.api(`/groups/${groupId}/members`).top(top).get();
  return result.value.map((m: any) => ({
    id: m.id,
    displayName: m.displayName,
    userPrincipalName: m.userPrincipalName ?? null,
    mail: m.mail ?? null,
    type: m["@odata.type"]?.replace("#microsoft.graph.", "") ?? null,
  }));
}

export async function addGroupMember(client: Client, params: { groupId: string; userId: string }) {
  const { groupId, userId } = params;
  await client.api(`/groups/${groupId}/members/$ref`).post({
    "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
  });
  return { success: true, message: "Mitglied hinzugefügt." };
}

export async function removeGroupMember(client: Client, params: { groupId: string; userId: string }) {
  const { groupId, userId } = params;
  await client.api(`/groups/${groupId}/members/${userId}/$ref`).delete();
  return { success: true, message: "Mitglied entfernt." };
}
