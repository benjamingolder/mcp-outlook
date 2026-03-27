import { Client } from "@microsoft/microsoft-graph-client";

export async function listSharepointSites(client: Client, params: { search?: string; top?: number }) {
  const { search = "", top = 10 } = params;

  const result = await client
    .api("/sites")
    .query({ search })
    .top(top)
    .get();

  return result.value.map((s: any) => ({
    id: s.id,
    name: s.name,
    displayName: s.displayName,
    webUrl: s.webUrl,
    description: s.description ?? null,
  }));
}

export async function getSharepointSite(client: Client, params: { siteId: string }) {
  const { siteId } = params;

  const s = await client.api(`/sites/${siteId}`).get();
  return {
    id: s.id,
    name: s.name,
    displayName: s.displayName,
    webUrl: s.webUrl,
    description: s.description ?? null,
    createdDateTime: s.createdDateTime,
    lastModifiedDateTime: s.lastModifiedDateTime,
  };
}

export async function listSharepointFiles(client: Client, params: {
  siteId: string;
  driveId?: string;
  folderId?: string;
  top?: number;
}) {
  const { siteId, driveId, folderId, top = 20 } = params;

  let path: string;
  if (driveId && folderId) {
    path = `/sites/${siteId}/drives/${driveId}/items/${folderId}/children`;
  } else if (driveId) {
    path = `/sites/${siteId}/drives/${driveId}/root/children`;
  } else {
    path = `/sites/${siteId}/drive/root/children`;
  }

  const result = await client.api(path).top(top).get();
  return result.value.map((item: any) => ({
    id: item.id,
    name: item.name,
    type: item.folder ? "folder" : "file",
    size: item.size ?? null,
    lastModifiedDateTime: item.lastModifiedDateTime,
    webUrl: item.webUrl,
    mimeType: item.file?.mimeType ?? null,
  }));
}

export async function listSharepointLists(client: Client, params: { siteId: string; top?: number }) {
  const { siteId, top = 20 } = params;

  const result = await client.api(`/sites/${siteId}/lists`).top(top).get();
  return result.value.map((l: any) => ({
    id: l.id,
    name: l.name,
    displayName: l.displayName,
    webUrl: l.webUrl,
    listType: l.list?.template ?? null,
    createdDateTime: l.createdDateTime,
  }));
}

export async function getSharepointList(client: Client, params: { siteId: string; listId: string }) {
  const { siteId, listId } = params;

  const l = await client.api(`/sites/${siteId}/lists/${listId}`).get();
  return {
    id: l.id,
    name: l.name,
    displayName: l.displayName,
    webUrl: l.webUrl,
    listType: l.list?.template ?? null,
    createdDateTime: l.createdDateTime,
    lastModifiedDateTime: l.lastModifiedDateTime,
    description: l.description ?? null,
  };
}

export async function updateSharepointList(client: Client, params: {
  siteId: string;
  listId: string;
  displayName?: string;
  description?: string;
}) {
  const { siteId, listId, displayName, description } = params;

  const body: Record<string, unknown> = {};
  if (displayName) body.displayName = displayName;
  if (description !== undefined) body.description = description;

  await client.api(`/sites/${siteId}/lists/${listId}`).patch(body);
  return { success: true, message: "Liste aktualisiert." };
}

export async function deleteSharepointList(client: Client, params: { siteId: string; listId: string }) {
  const { siteId, listId } = params;

  await client.api(`/sites/${siteId}/lists/${listId}`).delete();
  return { success: true, message: "Liste gelöscht." };
}

export async function createSharepointList(client: Client, params: {
  siteId: string;
  displayName: string;
  description?: string;
  columns: { name: string; type: "text" | "number" | "boolean" | "dateTime" | "choice"; choices?: string[] }[];
}) {
  const { siteId, displayName, description, columns } = params;

  const columnDefs = columns.map((col) => {
    const def: Record<string, unknown> = { name: col.name };
    if (col.type === "text") def.text = {};
    else if (col.type === "number") def.number = {};
    else if (col.type === "boolean") def.boolean = {};
    else if (col.type === "dateTime") def.dateTime = {};
    else if (col.type === "choice") def.choice = { choices: col.choices ?? [] };
    return def;
  });

  const result = await client.api(`/sites/${siteId}/lists`).post({
    displayName,
    ...(description && { description }),
    list: { template: "genericList" },
    columns: columnDefs,
  });

  return {
    id: result.id,
    displayName: result.displayName,
    webUrl: result.webUrl,
  };
}

export async function listSharepointListItems(client: Client, params: {
  siteId: string;
  listId: string;
  top?: number;
  filter?: string;
}) {
  const { siteId, listId, top = 20, filter } = params;

  let req = client
    .api(`/sites/${siteId}/lists/${listId}/items`)
    .expand("fields")
    .top(top);

  if (filter) req = req.filter(filter);

  const result = await req.get();
  return result.value.map((item: any) => ({
    id: item.id,
    createdDateTime: item.createdDateTime,
    lastModifiedDateTime: item.lastModifiedDateTime,
    webUrl: item.webUrl,
    fields: item.fields,
  }));
}

export async function getSharepointListItem(client: Client, params: {
  siteId: string;
  listId: string;
  itemId: string;
}) {
  const { siteId, listId, itemId } = params;

  const item = await client
    .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
    .expand("fields")
    .get();

  return {
    id: item.id,
    createdDateTime: item.createdDateTime,
    lastModifiedDateTime: item.lastModifiedDateTime,
    webUrl: item.webUrl,
    fields: item.fields,
  };
}

export async function createSharepointListItem(client: Client, params: {
  siteId: string;
  listId: string;
  fields: Record<string, unknown>;
}) {
  const { siteId, listId, fields } = params;

  const result = await client
    .api(`/sites/${siteId}/lists/${listId}/items`)
    .post({ fields });

  return { id: result.id, webUrl: result.webUrl, fields: result.fields };
}

export async function updateSharepointListItem(client: Client, params: {
  siteId: string;
  listId: string;
  itemId: string;
  fields: Record<string, unknown>;
}) {
  const { siteId, listId, itemId, fields } = params;

  await client
    .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`)
    .patch(fields);

  return { success: true, message: "Eintrag aktualisiert." };
}

export async function deleteSharepointListItem(client: Client, params: {
  siteId: string;
  listId: string;
  itemId: string;
}) {
  const { siteId, listId, itemId } = params;

  await client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}`).delete();
  return { success: true, message: "Eintrag gelöscht." };
}

export async function searchSharepoint(client: Client, params: { query: string; top?: number }) {
  const { query, top = 10 } = params;

  const result = await client.api("/search/query").post({
    requests: [
      {
        entityTypes: ["driveItem", "listItem"],
        query: { queryString: query },
        from: 0,
        size: top,
      },
    ],
  });

  const hits = result.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
  return hits.map((h: any) => ({
    id: h.resource?.id,
    name: h.resource?.name,
    webUrl: h.resource?.webUrl,
    lastModifiedDateTime: h.resource?.lastModifiedDateTime,
    summary: h.summary ?? null,
  }));
}

export async function createSharepointFolder(client: Client, params: {
  siteId: string;
  driveId?: string;
  parentId?: string;
  folderName: string;
}) {
  const { siteId, driveId, parentId, folderName } = params;

  let path: string;
  if (driveId && parentId) {
    path = `/sites/${siteId}/drives/${driveId}/items/${parentId}/children`;
  } else if (driveId) {
    path = `/sites/${siteId}/drives/${driveId}/root/children`;
  } else if (parentId) {
    path = `/sites/${siteId}/drive/items/${parentId}/children`;
  } else {
    path = `/sites/${siteId}/drive/root/children`;
  }

  const result = await client.api(path).post({
    name: folderName,
    folder: {},
    "@microsoft.graph.conflictBehavior": "rename",
  });

  return {
    id: result.id,
    name: result.name,
    webUrl: result.webUrl,
    createdDateTime: result.createdDateTime,
  };
}

export async function uploadSharepointFile(client: Client, params: {
  siteId: string;
  driveId?: string;
  parentId?: string;
  fileName: string;
  content: string;
  mimeType?: string;
}) {
  const { siteId, driveId, parentId, fileName, content, mimeType = "text/plain" } = params;

  let path: string;
  const encodedName = encodeURIComponent(fileName);
  if (driveId && parentId) {
    path = `/sites/${siteId}/drives/${driveId}/items/${parentId}:/${encodedName}:/content`;
  } else if (driveId) {
    path = `/sites/${siteId}/drives/${driveId}/root:/${encodedName}:/content`;
  } else if (parentId) {
    path = `/sites/${siteId}/drive/items/${parentId}:/${encodedName}:/content`;
  } else {
    path = `/sites/${siteId}/drive/root:/${encodedName}:/content`;
  }

  const buffer = Buffer.from(content, "utf-8");
  const result = await client.api(path).header("Content-Type", mimeType).put(buffer);

  return {
    id: result.id,
    name: result.name,
    size: result.size,
    webUrl: result.webUrl,
    createdDateTime: result.createdDateTime,
  };
}

export async function deleteSharepointFile(client: Client, params: {
  siteId: string;
  driveId?: string;
  itemId: string;
}) {
  const { siteId, driveId, itemId } = params;

  const path = driveId
    ? `/sites/${siteId}/drives/${driveId}/items/${itemId}`
    : `/sites/${siteId}/drive/items/${itemId}`;

  await client.api(path).delete();
  return { success: true, message: "Datei/Ordner gelöscht." };
}

export async function moveSharepointFile(client: Client, params: {
  siteId: string;
  driveId?: string;
  itemId: string;
  destinationParentId: string;
  newName?: string;
}) {
  const { siteId, driveId, itemId, destinationParentId, newName } = params;

  const path = driveId
    ? `/sites/${siteId}/drives/${driveId}/items/${itemId}`
    : `/sites/${siteId}/drive/items/${itemId}`;

  const body: Record<string, unknown> = {
    parentReference: { id: destinationParentId },
  };
  if (newName) body.name = newName;

  const result = await client.api(path).patch(body);
  return {
    id: result.id,
    name: result.name,
    webUrl: result.webUrl,
    parentPath: result.parentReference?.path ?? null,
  };
}
