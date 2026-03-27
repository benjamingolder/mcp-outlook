import { Client } from "@microsoft/microsoft-graph-client";

export async function listTodoLists(client: Client) {
  const result = await client.api("/me/todo/lists").get();
  return result.value.map((l: any) => ({
    id: l.id,
    displayName: l.displayName,
    isOwner: l.isOwner,
    isShared: l.isShared,
  }));
}

export async function listTasks(client: Client, params: { listId: string; filter?: string; top?: number }) {
  const { listId, filter, top = 20 } = params;

  let req = client
    .api(`/me/todo/lists/${listId}/tasks`)
    .top(top)
    .orderby("createdDateTime DESC");

  if (filter) req = req.filter(filter);

  const result = await req.get();
  return result.value.map((t: any) => ({
    id: t.id,
    title: t.title,
    status: t.status,
    importance: t.importance,
    dueDateTime: t.dueDateTime?.dateTime ?? null,
    completedDateTime: t.completedDateTime?.dateTime ?? null,
    body: t.body?.content ?? null,
  }));
}

export async function createTask(client: Client, params: {
  listId: string;
  title: string;
  body?: string;
  dueDateTime?: string;
  importance?: "low" | "normal" | "high";
}) {
  const { listId, title, body, dueDateTime, importance = "normal" } = params;

  const task: Record<string, unknown> = { title, importance };
  if (body) task.body = { contentType: "text", content: body };
  if (dueDateTime) task.dueDateTime = { dateTime: dueDateTime, timeZone: "UTC" };

  const result = await client.api(`/me/todo/lists/${listId}/tasks`).post(task);
  return { id: result.id, title: result.title, status: result.status };
}

export async function updateTask(client: Client, params: {
  listId: string;
  taskId: string;
  title?: string;
  status?: "notStarted" | "inProgress" | "completed" | "waitingOnOthers" | "deferred";
  importance?: "low" | "normal" | "high";
  dueDateTime?: string;
  body?: string;
}) {
  const { listId, taskId, title, status, importance, dueDateTime, body } = params;

  const patch: Record<string, unknown> = {};
  if (title) patch.title = title;
  if (status) patch.status = status;
  if (importance) patch.importance = importance;
  if (dueDateTime) patch.dueDateTime = { dateTime: dueDateTime, timeZone: "UTC" };
  if (body) patch.body = { contentType: "text", content: body };

  await client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).patch(patch);
  return { success: true, message: "Aufgabe aktualisiert." };
}

export async function deleteTask(client: Client, params: { listId: string; taskId: string }) {
  const { listId, taskId } = params;
  await client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).delete();
  return { success: true, message: "Aufgabe gelöscht." };
}
