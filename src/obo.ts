// On-Behalf-Of (OBO) flow: exchange user token for Microsoft Graph token

const tenantId = process.env.ENTRA_TENANT_ID!;
const clientId = process.env.ENTRA_CLIENT_ID!;
const clientSecret = process.env.ENTRA_CLIENT_SECRET!;

const GRAPH_SCOPES = [
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/Tasks.ReadWrite",
  "https://graph.microsoft.com/Contacts.ReadWrite",
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/Sites.ReadWrite.All",
  "https://graph.microsoft.com/Team.ReadBasic.All",
  "https://graph.microsoft.com/Channel.ReadBasic.All",
  "https://graph.microsoft.com/ChannelMessage.Send",
  "https://graph.microsoft.com/ChannelMessage.Read.All",
  "https://graph.microsoft.com/Chat.ReadWrite",
  "https://graph.microsoft.com/Notes.ReadWrite",
  "https://graph.microsoft.com/Presence.ReadWrite",
  "https://graph.microsoft.com/People.Read",
  "https://graph.microsoft.com/User.Read.All",
  "https://graph.microsoft.com/Group.ReadWrite.All",
  "https://graph.microsoft.com/Bookings.ReadWrite.All",
];

export async function getGraphTokenViaObo(userAccessToken: string): Promise<string> {
  const params = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    client_id: clientId,
    client_secret: clientSecret,
    assertion: userAccessToken,
    requested_token_use: "on_behalf_of",
    scope: GRAPH_SCOPES.join(" "),
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    }
  );

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`OBO Token-Exchange fehlgeschlagen: ${err}`);
  }

  const data = await res.json() as any;
  return data.access_token;
}
