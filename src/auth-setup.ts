import "dotenv/config";
import { getAccessToken } from "./auth.js";

console.log("=== MCP Outlook – Verbindungstest ===\n");

getAccessToken()
  .then(() => {
    console.log("✓ Client Credentials erfolgreich! Der MCP Server kann gestartet werden.");
    process.exit(0);
  })
  .catch((err: Error) => {
    console.error("✗ Authentifizierung fehlgeschlagen:", err.message);
    process.exit(1);
  });
