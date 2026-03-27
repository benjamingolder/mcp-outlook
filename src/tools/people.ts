import { Client } from "@microsoft/microsoft-graph-client";

export async function listRelevantPeople(client: Client, params: { top?: number; search?: string }) {
  const { top = 20, search } = params;

  let req = client.api("/me/people").top(top);
  if (search) req = req.search(`"${search}"`);

  const result = await req.get();
  return result.value.map((p: any) => ({
    id: p.id,
    displayName: p.displayName,
    jobTitle: p.jobTitle,
    companyName: p.companyName,
    department: p.department,
    emailAddresses: p.emailAddresses,
    phones: p.phones,
    relevanceScore: p.relevanceScore,
  }));
}

export async function listTrendingDocuments(client: Client, params: { top?: number }) {
  const { top = 10 } = params;
  const result = await client.api("/me/insights/trending").top(top).get();
  return result.value.map((item: any) => ({
    id: item.id,
    resourceType: item.resourceVisualization?.type,
    title: item.resourceVisualization?.title,
    previewText: item.resourceVisualization?.previewText,
    webUrl: item.resourceReference?.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
  }));
}

export async function listUsedDocuments(client: Client, params: { top?: number }) {
  const { top = 10 } = params;
  const result = await client.api("/me/insights/used").top(top).get();
  return result.value.map((item: any) => ({
    id: item.id,
    resourceType: item.resourceVisualization?.type,
    title: item.resourceVisualization?.title,
    webUrl: item.resourceReference?.webUrl,
    lastUsedDateTime: item.lastUsed?.lastAccessedDateTime,
  }));
}

export async function listSharedDocuments(client: Client, params: { top?: number }) {
  const { top = 10 } = params;
  const result = await client.api("/me/insights/shared").top(top).get();
  return result.value.map((item: any) => ({
    id: item.id,
    resourceType: item.resourceVisualization?.type,
    title: item.resourceVisualization?.title,
    webUrl: item.resourceReference?.webUrl,
    sharedBy: item.lastShared?.sharedBy?.user?.displayName ?? null,
    sharedDateTime: item.lastShared?.sharedDateTime,
  }));
}
