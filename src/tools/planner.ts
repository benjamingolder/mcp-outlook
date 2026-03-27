import { Client } from "@microsoft/microsoft-graph-client";

export async function listPlans(client: Client, params: { groupId: string }) {
  const { groupId } = params;
  const result = await client.api(`/groups/${groupId}/planner/plans`).get();
  return result.value.map((p: any) => ({
    id: p.id,
    title: p.title,
    createdDateTime: p.createdDateTime,
    owner: p.owner,
  }));
}

export async function listMyPlannerTasks(client: Client, params: { top?: number }) {
  const { top = 20 } = params;
  const result = await client.api("/me/planner/tasks").top(top).get();
  return result.value.map((t: any) => ({
    id: t.id,
    title: t.title,
    planId: t.planId,
    bucketId: t.bucketId,
    percentComplete: t.percentComplete,
    priority: t.priority,
    dueDateTime: t.dueDateTime,
    startDateTime: t.startDateTime,
    createdDateTime: t.createdDateTime,
    assignedTo: Object.keys(t.assignments ?? {}),
  }));
}

export async function listBuckets(client: Client, params: { planId: string }) {
  const { planId } = params;
  const result = await client.api(`/planner/plans/${planId}/buckets`).get();
  return result.value.map((b: any) => ({
    id: b.id,
    name: b.name,
    planId: b.planId,
    orderHint: b.orderHint,
  }));
}

export async function listPlanTasks(client: Client, params: { planId: string }) {
  const { planId } = params;
  const result = await client.api(`/planner/plans/${planId}/tasks`).get();
  return result.value.map((t: any) => ({
    id: t.id,
    title: t.title,
    bucketId: t.bucketId,
    percentComplete: t.percentComplete,
    priority: t.priority,
    dueDateTime: t.dueDateTime,
    assignedTo: Object.keys(t.assignments ?? {}),
  }));
}

export async function createPlannerTask(client: Client, params: {
  planId: string;
  title: string;
  bucketId?: string;
  dueDateTime?: string;
  assignedToUserIds?: string[];
  priority?: number;
}) {
  const { planId, title, bucketId, dueDateTime, assignedToUserIds = [], priority } = params;

  const task: Record<string, unknown> = { planId, title };
  if (bucketId) task.bucketId = bucketId;
  if (dueDateTime) task.dueDateTime = dueDateTime;
  if (priority !== undefined) task.priority = priority;
  if (assignedToUserIds.length > 0) {
    task.assignments = Object.fromEntries(
      assignedToUserIds.map((id) => [id, { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" }])
    );
  }

  const result = await client.api("/planner/tasks").post(task);
  return { id: result.id, title: result.title, planId: result.planId };
}

export async function updatePlannerTask(client: Client, params: {
  taskId: string;
  title?: string;
  percentComplete?: number;
  dueDateTime?: string;
  priority?: number;
  bucketId?: string;
}) {
  const { taskId, ...patch } = params;

  // Etag needed for PATCH on planner tasks
  const existing = await client.api(`/planner/tasks/${taskId}`).get();
  const etag = existing["@odata.etag"];

  await client
    .api(`/planner/tasks/${taskId}`)
    .header("If-Match", etag)
    .patch(patch);

  return { success: true, message: "Planner-Aufgabe aktualisiert." };
}

export async function deletePlannerTask(client: Client, params: { taskId: string }) {
  const { taskId } = params;

  const existing = await client.api(`/planner/tasks/${taskId}`).get();
  const etag = existing["@odata.etag"];

  await client
    .api(`/planner/tasks/${taskId}`)
    .header("If-Match", etag)
    .delete();

  return { success: true, message: "Planner-Aufgabe gelöscht." };
}
