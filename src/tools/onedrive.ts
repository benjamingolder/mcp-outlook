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
