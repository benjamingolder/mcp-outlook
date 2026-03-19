import {
  PublicClientApplication,
  type TokenCacheContext,
} from "@azure/msal-node";
import * as fs from "fs";
import * as path from "path";

const CACHE_PATH = process.env.TOKEN_CACHE_PATH ?? "/data/token-cache.json";

// Scopes ohne Admin Consent (für Geschäftskonten ohne eigene App Registration)
const SCOPES_WORK = [
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/Tasks.ReadWrite",
  "https://graph.microsoft.com/Contacts.ReadWrite",
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/Team.ReadBasic.All",
  "https://graph.microsoft.com/Channel.ReadBasic.All",
  "https://graph.microsoft.com/ChannelMessage.Send",
  "https://graph.microsoft.com/Chat.ReadWrite",
  "https://graph.microsoft.com/Notes.ReadWrite",
  "https://graph.microsoft.com/Presence.ReadWrite",
  "https://graph.microsoft.com/People.Read",
  "offline_access",
];

// Alle Scopes inkl. Admin Consent (für eigene App Registration)
const SCOPES_FULL = [
  ...SCOPES_WORK,
  "https://graph.microsoft.com/Sites.ReadWrite.All",
  "https://graph.microsoft.com/ChannelMessage.Read.All",
  "https://graph.microsoft.com/User.Read.All",
  "https://graph.microsoft.com/Group.ReadWrite.All",
  "https://graph.microsoft.com/Bookings.ReadWrite.All",
];

export const SCOPES =
  process.env.SCOPE_PRESET === "work" ? SCOPES_WORK : SCOPES_FULL;

const cachePlugin = {
  beforeCacheAccess: async (ctx: TokenCacheContext) => {
    if (fs.existsSync(CACHE_PATH)) {
      ctx.tokenCache.deserialize(await fs.promises.readFile(CACHE_PATH, "utf8"));
    }
  },
  afterCacheAccess: async (ctx: TokenCacheContext) => {
    if (ctx.cacheHasChanged) {
      await fs.promises.mkdir(path.dirname(CACHE_PATH), { recursive: true });
      await fs.promises.writeFile(CACHE_PATH, ctx.tokenCache.serialize());
    }
  },
};

let pca: PublicClientApplication | null = null;

function getApp(): PublicClientApplication {
  if (!pca) {
    // "common" oder "organizations" als TENANT_ID erlaubt Multi-Tenant-Login
    // (z.B. für Schulkonten in einem fremden Tenant)
    const tenantId = process.env.TENANT_ID ?? "common";
    pca = new PublicClientApplication({
      auth: {
        clientId: process.env.CLIENT_ID!,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: { cachePlugin },
    });
  }
  return pca;
}

export async function getAccessToken(): Promise<string> {
  const app = getApp();
  const accounts = await app.getTokenCache().getAllAccounts();

  if (accounts.length === 0) {
    throw new Error(
      "Nicht authentifiziert. Bitte zuerst ausführen: docker exec mcp-outlook node dist/auth-setup.js"
    );
  }

  const result = await app.acquireTokenSilent({
    account: accounts[0],
    scopes: SCOPES,
  });

  if (!result) throw new Error("Token konnte nicht erneuert werden.");
  return result.accessToken;
}

export async function doDeviceCodeFlow(): Promise<void> {
  const app = getApp();
  await app.acquireTokenByDeviceCode({
    deviceCodeCallback: (response) => {
      console.log("\n" + response.message + "\n");
    },
    scopes: SCOPES,
  });
}
