import { getGraphClient } from "../graph.js";

export async function listEmails(params: {
  top?: number;
  folder?: string;
  filter?: string;
}) {
  const { top = 20, folder = "inbox", filter } = params;
  const client = getGraphClient();

  let req = client
    .api(`/me/mailFolders/${folder}/messages`)
    .top(top)
    .orderby("receivedDateTime DESC")
    .select("id,subject,from,receivedDateTime,bodyPreview,isRead");

  if (filter) req = req.filter(filter);

  const result = await req.get();
  return result.value.map((m: any) => ({
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress,
    receivedDateTime: m.receivedDateTime,
    bodyPreview: m.bodyPreview,
    isRead: m.isRead,
  }));
}

export async function readEmail(id: string) {
  const client = getGraphClient();
  const m = await client.api(`/me/messages/${id}`).get();
  return {
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress,
    to: m.toRecipients?.map((r: any) => r.emailAddress),
    cc: m.ccRecipients?.map((r: any) => r.emailAddress),
    receivedDateTime: m.receivedDateTime,
    body: m.body?.content,
    bodyType: m.body?.contentType,
  };
}

export async function sendEmail(params: {
  to: string[];
  subject: string;
  body: string;
  cc?: string[];
  bodyType?: "html" | "text";
}) {
  const { to, subject, body, cc = [], bodyType = "text" } = params;
  const client = getGraphClient();

  await client.api("/me/sendMail").post({
    message: {
      subject,
      body: {
        contentType: bodyType === "html" ? "HTML" : "Text",
        content: body,
      },
      toRecipients: to.map((addr) => ({ emailAddress: { address: addr } })),
      ccRecipients: cc.map((addr) => ({ emailAddress: { address: addr } })),
    },
  });

  return { success: true, message: "E-Mail erfolgreich gesendet." };
}

export async function replyToEmail(params: {
  id: string;
  body: string;
  replyAll?: boolean;
  bodyType?: "html" | "text";
}) {
  const { id, body, replyAll = false, bodyType = "text" } = params;
  const client = getGraphClient();

  const endpoint = replyAll
    ? `/me/messages/${id}/replyAll`
    : `/me/messages/${id}/reply`;

  await client.api(endpoint).post({
    message: {
      body: {
        contentType: bodyType === "html" ? "HTML" : "Text",
        content: body,
      },
    },
  });

  return { success: true, message: "Antwort erfolgreich gesendet." };
}
