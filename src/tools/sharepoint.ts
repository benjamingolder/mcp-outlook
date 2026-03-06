import { getGraphClient } from "../graph.js";

export async function listSharepointSites(params: { search?: string; top?: number }) {
  const { search = "", top = 10 } = params;
  const client = getGraphClient();

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

export async function listSharepointFiles(params: {
  siteId: string;
  driveId?: string;
  folderId?: string;
  top?: number;
}) {
  const { siteId, driveId, folderId, top = 20 } = params;
  const client = getGraphClient();

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

export async function searchSharepoint(params: { query: string; top?: number }) {
  const { query, top = 10 } = params;
  const client = getGraphClient();

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
