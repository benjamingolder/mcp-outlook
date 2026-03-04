import "dotenv/config";
import { doDeviceCodeFlow } from "./auth.js";

console.log("=== MCP Outlook – Authentifizierung ===\n");
console.log("Folge den Anweisungen um dein Microsoft-Konto zu verbinden:\n");

doDeviceCodeFlow()
  .then(() => {
    console.log("\n✓ Authentifizierung erfolgreich! Der MCP Server kann nun gestartet werden.");
    process.exit(0);
  })
  .catch((err: Error) => {
    console.error("✗ Authentifizierung fehlgeschlagen:", err.message);
    process.exit(1);
  });
