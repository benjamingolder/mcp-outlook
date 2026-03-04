import { getGraphClient } from "../graph.js";

export async function listEvents(params: {
  startDateTime?: string;
  endDateTime?: string;
  top?: number;
}) {
  const {
    startDateTime = new Date().toISOString(),
    endDateTime = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString(),
    top = 20,
  } = params;

  const client = getGraphClient();
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
    attendees: e.attendees?.map((a: any) => ({
      email: a.emailAddress,
      status: a.status?.response,
    })),
  }));
}

export async function getEvent(id: string) {
  const client = getGraphClient();
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
  };
}

export async function createEvent(params: {
  subject: string;
  start: string;
  end: string;
  body?: string;
  location?: string;
  attendees?: string[];
  isAllDay?: boolean;
  bodyType?: "html" | "text";
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
  } = params;

  const client = getGraphClient();
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

export async function updateEvent(params: {
  id: string;
  subject?: string;
  start?: string;
  end?: string;
  body?: string;
  location?: string;
}) {
  const { id, subject, start, end, body, location } = params;
  const client = getGraphClient();

  const patch: Record<string, unknown> = {};
  if (subject) patch.subject = subject;
  if (start) patch.start = { dateTime: start, timeZone: "UTC" };
  if (end) patch.end = { dateTime: end, timeZone: "UTC" };
  if (body) patch.body = { contentType: "Text", content: body };
  if (location) patch.location = { displayName: location };

  await client.api(`/me/events/${id}`).patch(patch);
  return { success: true, message: "Termin aktualisiert." };
}

export async function deleteEvent(id: string) {
  const client = getGraphClient();
  await client.api(`/me/events/${id}`).delete();
  return { success: true, message: "Termin gelöscht." };
}
