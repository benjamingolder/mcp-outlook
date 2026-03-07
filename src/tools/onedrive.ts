import { getGraphClient } from "../graph.js";

export async function listOneDriveFiles(params: { folderId?: string; top?: number }) {
  const { folderId, top = 20 } = params;
  const client = getGraphClient();

  const path = folderId
    ? `/me/drive/items/${folderId}/children`
    : `/me/drive/root/children`;

  const result = await client.api(path).top(top).get();
  return result.value.map((item: any) => ({
    id: item.id,
    name: item.name,
    type: item.folder ? "folder" : "file",
    size: item.size ?? null,
    lastModifiedDateTime: item.lastModifiedDateTime,
    webUrl: item.webUrl,
    mimeType: item.file?.mimeType ?? null,
    parentPath: item.parentReference?.path ?? null,
  }));
}

export async function searchOneDrive(params: { query: string; top?: number }) {
  const { query, top = 20 } = params;
  const client = getGraphClient();

  const result = await client
    .api(`/me/drive/root/search(q='${encodeURIComponent(query)}')`)
    .top(top)
    .get();

  return result.value.map((item: any) => ({
    id: item.id,
    name: item.name,
    type: item.folder ? "folder" : "file",
    size: item.size ?? null,
    lastModifiedDateTime: item.lastModifiedDateTime,
    webUrl: item.webUrl,
    mimeType: item.file?.mimeType ?? null,
    parentPath: item.parentReference?.path ?? null,
  }));
}

export async function getOneDriveFileInfo(params: { fileId: string }) {
  const { fileId } = params;
  const client = getGraphClient();

  const item = await client.api(`/me/drive/items/${fileId}`).get();
  return {
    id: item.id,
    name: item.name,
    type: item.folder ? "folder" : "file",
    size: item.size ?? null,
    lastModifiedDateTime: item.lastModifiedDateTime,
    createdDateTime: item.createdDateTime,
    webUrl: item.webUrl,
    downloadUrl: item["@microsoft.graph.downloadUrl"] ?? null,
    mimeType: item.file?.mimeType ?? null,
    parentPath: item.parentReference?.path ?? null,
  };
}

export async function createOneDriveFolder(params: { parentId?: string; folderName: string }) {
  const { parentId, folderName } = params;
  const client = getGraphClient();

  const path = parentId
    ? `/me/drive/items/${parentId}/children`
    : `/me/drive/root/children`;

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

export async function uploadOneDriveFile(params: {
  parentId?: string;
  fileName: string;
  content: string;
  mimeType?: string;
}) {
  const { parentId, fileName, content, mimeType = "text/plain" } = params;
  const client = getGraphClient();

  const path = parentId
    ? `/me/drive/items/${parentId}:/${encodeURIComponent(fileName)}:/content`
    : `/me/drive/root:/${encodeURIComponent(fileName)}:/content`;

  const buffer = Buffer.from(content, "utf-8");

  const result = await client
    .api(path)
    .header("Content-Type", mimeType)
    .put(buffer);

  return {
    id: result.id,
    name: result.name,
    size: result.size,
    webUrl: result.webUrl,
    createdDateTime: result.createdDateTime,
  };
}

export async function deleteOneDriveItem(params: { itemId: string }) {
  const { itemId } = params;
  const client = getGraphClient();

  await client.api(`/me/drive/items/${itemId}`).delete();
  return { success: true, message: "Element gelöscht." };
}

export async function moveOneDriveItem(params: {
  itemId: string;
  destinationParentId: string;
  newName?: string;
}) {
  const { itemId, destinationParentId, newName } = params;
  const client = getGraphClient();

  const body: Record<string, unknown> = {
    parentReference: { id: destinationParentId },
  };
  if (newName) body.name = newName;

  const result = await client.api(`/me/drive/items/${itemId}`).patch(body);
  return {
    id: result.id,
    name: result.name,
    webUrl: result.webUrl,
    parentPath: result.parentReference?.path ?? null,
  };
}

export async function renameOneDriveItem(params: { itemId: string; newName: string }) {
  const { itemId, newName } = params;
  const client = getGraphClient();

  const result = await client.api(`/me/drive/items/${itemId}`).patch({ name: newName });
  return {
    id: result.id,
    name: result.name,
    webUrl: result.webUrl,
  };
}

export async function copyOneDriveItem(params: {
  itemId: string;
  destinationParentId: string;
  newName?: string;
}) {
  const { itemId, destinationParentId, newName } = params;
  const client = getGraphClient();

  const body: Record<string, unknown> = {
    parentReference: { id: destinationParentId },
  };
  if (newName) body.name = newName;

  await client.api(`/me/drive/items/${itemId}/copy`).post(body);
  return { success: true, message: "Element wird kopiert (asynchron)." };
}
