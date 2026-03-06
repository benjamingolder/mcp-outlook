import { getGraphClient } from "../graph.js";

export async function listNotebooks(params: { top?: number }) {
  const { top = 20 } = params;
  const client = getGraphClient();
  const result = await client.api("/me/onenote/notebooks").top(top).get();
  return result.value.map((n: any) => ({
    id: n.id,
    displayName: n.displayName,
    createdDateTime: n.createdDateTime,
    lastModifiedDateTime: n.lastModifiedDateTime,
    webUrl: n.links?.oneNoteWebUrl?.href ?? null,
  }));
}

export async function listSections(params: { notebookId: string }) {
  const { notebookId } = params;
  const client = getGraphClient();
  const result = await client
    .api(`/me/onenote/notebooks/${notebookId}/sections`)
    .get();
  return result.value.map((s: any) => ({
    id: s.id,
    displayName: s.displayName,
    createdDateTime: s.createdDateTime,
    lastModifiedDateTime: s.lastModifiedDateTime,
  }));
}

export async function listPages(params: { sectionId: string; top?: number }) {
  const { sectionId, top = 20 } = params;
  const client = getGraphClient();
  const result = await client
    .api(`/me/onenote/sections/${sectionId}/pages`)
    .top(top)
    .get();
  return result.value.map((p: any) => ({
    id: p.id,
    title: p.title,
    createdDateTime: p.createdDateTime,
    lastModifiedDateTime: p.lastModifiedDateTime,
    webUrl: p.links?.oneNoteWebUrl?.href ?? null,
  }));
}

export async function getPage(params: { pageId: string }) {
  const { pageId } = params;
  const client = getGraphClient();
  const content = await client.api(`/me/onenote/pages/${pageId}/content`).getStream();
  const chunks: Buffer[] = [];
  for await (const chunk of content) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return { content: Buffer.concat(chunks).toString("utf8") };
}

export async function createPage(params: {
  sectionId: string;
  title: string;
  content?: string;
}) {
  const { sectionId, title, content = "" } = params;
  const client = getGraphClient();
  const html = `<!DOCTYPE html><html><head><title>${title}</title></head><body>${content}</body></html>`;
  const result = await client
    .api(`/me/onenote/sections/${sectionId}/pages`)
    .header("Content-Type", "text/html")
    .post(html);
  return {
    id: result.id,
    title: result.title,
    webUrl: result.links?.oneNoteWebUrl?.href ?? null,
  };
}
