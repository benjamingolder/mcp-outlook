import { ConfidentialClientApplication } from "@azure/msal-node";

const SCOPE = "https://graph.microsoft.com/.default";

let cca: ConfidentialClientApplication | null = null;

function getApp(): ConfidentialClientApplication {
  if (!cca) {
    cca = new ConfidentialClientApplication({
      auth: {
        clientId: process.env.CLIENT_ID!,
        clientSecret: process.env.CLIENT_SECRET!,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
      },
    });
  }
  return cca;
}

export async function getAccessToken(): Promise<string> {
  const result = await getApp().acquireTokenByClientCredential({
    scopes: [SCOPE],
  });

  if (!result) throw new Error("Token konnte nicht abgerufen werden.");
  return result.accessToken;
}
