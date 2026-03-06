import {
  PublicClientApplication,
  type TokenCacheContext,
} from "@azure/msal-node";
import * as fs from "fs";
import * as path from "path";

const CACHE_PATH = process.env.TOKEN_CACHE_PATH ?? "/data/token-cache.json";

export const SCOPES = [
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.Read",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/Tasks.ReadWrite",
  "https://graph.microsoft.com/Sites.Read.All",
  "https://graph.microsoft.com/Files.Read.All",
  "offline_access",
];

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
    pca = new PublicClientApplication({
      auth: {
        clientId: process.env.CLIENT_ID!,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
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
