import { Client } from "@microsoft/microsoft-graph-client";

export async function listEvents(client: Client, params: {
  startDateTime?: string;
  endDateTime?: string;
  top?: number;
}) {
  const {
    startDateTime = new Date().toISOString(),
    endDateTime = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString(),
    top = 20,
  } = params;

  const result = await client
    .api("/me/calendarView")
    .query({ startDateTime, endDateTime })
    .top(top)
    .orderby("start/dateTime")
    .get();

  return result.value.map((e: any) => ({
    id: e.id,
    subject: e.subject,
    start: e.start,
    end: e.end,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress,
    isAllDay: e.isAllDay,
    bodyPreview: e.bodyPreview,
    categories: e.categories ?? [],
    attendees: e.attendees?.map((a: any) => ({
      email: a.emailAddress,
      status: a.status?.response,
    })),
  }));
}

export async function getEvent(client: Client, id: string) {
  const e = await client.api(`/me/events/${id}`).get();
  return {
    id: e.id,
    subject: e.subject,
    start: e.start,
    end: e.end,
    location: e.location?.displayName,
    body: e.body?.content,
    bodyType: e.body?.contentType,
    organizer: e.organizer?.emailAddress,
    attendees: e.attendees?.map((a: any) => ({
      email: a.emailAddress,
      status: a.status?.response,
    })),
    isAllDay: e.isAllDay,
    categories: e.categories ?? [],
  };
}

export async function createEvent(client: Client, params: {
  subject: string;
  start: string;
  end: string;
  body?: string;
  location?: string;
  attendees?: string[];
  isAllDay?: boolean;
  bodyType?: "html" | "text";
  categories?: string[];
}) {
  const {
    subject,
    start,
    end,
    body = "",
    location,
    attendees = [],
    isAllDay = false,
    bodyType = "text",
    categories,
  } = params;

  const event = await client.api("/me/events").post({
    subject,
    isAllDay,
    start: { dateTime: start, timeZone: "UTC" },
    end: { dateTime: end, timeZone: "UTC" },
    body: {
      contentType: bodyType === "html" ? "HTML" : "Text",
      content: body,
    },
    ...(location && { location: { displayName: location } }),
    ...(categories && categories.length > 0 && { categories }),
    attendees: attendees.map((addr) => ({
      emailAddress: { address: addr },
      type: "required",
    })),
  });

  return {
    id: event.id,
    subject: event.subject,
    start: event.start,
    end: event.end,
    webLink: event.webLink,
  };
}

export async function updateEvent(client: Client, params: {
  id: string;
  subject?: string;
  start?: string;
  end?: string;
  body?: string;
  location?: string;
  categories?: string[];
}) {
  const { id, subject, start, end, body, location, categories } = params;

  const patch: Record<string, unknown> = {};
  if (subject) patch.subject = subject;
  if (start) patch.start = { dateTime: start, timeZone: "UTC" };
  if (end) patch.end = { dateTime: end, timeZone: "UTC" };
  if (body) patch.body = { contentType: "Text", content: body };
  if (location) patch.location = { displayName: location };
  if (categories !== undefined) patch.categories = categories;

  await client.api(`/me/events/${id}`).patch(patch);
  return { success: true, message: "Termin aktualisiert." };
}

export async function deleteEvent(client: Client, id: string) {
  await client.api(`/me/events/${id}`).delete();
  return { success: true, message: "Termin gelöscht." };
}
