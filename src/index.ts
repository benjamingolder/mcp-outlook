import "dotenv/config";
import express from "express";
import cors from "cors";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ErrorCode,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import { listEmails, readEmail, sendEmail, replyToEmail } from "./tools/mail.js";
import {
  listEvents,
  getEvent,
  createEvent,
  updateEvent,
  deleteEvent,
} from "./tools/calendar.js";

function createMcpServer(): Server {
  const server = new Server(
    { name: "outlook-mcp", version: "1.0.0" },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: "list_emails",
        description: "Listet E-Mails aus einem Outlook-Ordner auf",
        inputSchema: {
          type: "object",
          properties: {
            top: { type: "number", description: "Anzahl der E-Mails (Standard: 20)" },
            folder: { type: "string", description: "Ordner: inbox, sentitems, drafts (Standard: inbox)" },
            filter: { type: "string", description: "OData-Filterausdruck" },
          },
        },
      },
      {
        name: "read_email",
        description: "Liest den vollständigen Inhalt einer E-Mail",
        inputSchema: {
          type: "object",
          properties: {
            id: { type: "string", description: "E-Mail-ID" },
          },
          required: ["id"],
        },
      },
      {
        name: "send_email",
        description: "Sendet eine E-Mail über Outlook",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "array", items: { type: "string" }, description: "Empfänger-Adressen" },
            subject: { type: "string", description: "Betreff" },
            body: { type: "string", description: "Nachrichtentext" },
            cc: { type: "array", items: { type: "string" }, description: "CC-Empfänger" },
            bodyType: { type: "string", enum: ["text", "html"], description: "Textformat (Standard: text)" },
          },
          required: ["to", "subject", "body"],
        },
      },
      {
        name: "reply_to_email",
        description: "Antwortet auf eine E-Mail",
        inputSchema: {
          type: "object",
          properties: {
            id: { type: "string", description: "E-Mail-ID" },
            body: { type: "string", description: "Antworttext" },
            replyAll: { type: "boolean", description: "Allen antworten (Standard: false)" },
            bodyType: { type: "string", enum: ["text", "html"] },
          },
          required: ["id", "body"],
        },
      },
      {
        name: "list_events",
        description: "Listet Kalendertermine in einem Zeitraum auf",
        inputSchema: {
          type: "object",
          properties: {
            startDateTime: { type: "string", description: "Beginn (ISO 8601, Standard: jetzt)" },
            endDateTime: { type: "string", description: "Ende (ISO 8601, Standard: +7 Tage)" },
            top: { type: "number", description: "Anzahl der Termine (Standard: 20)" },
          },
        },
      },
      {
        name: "get_event",
        description: "Liest Details eines Kalendertermins",
        inputSchema: {
          type: "object",
          properties: {
            id: { type: "string", description: "Termin-ID" },
          },
          required: ["id"],
        },
      },
      {
        name: "create_event",
        description: "Erstellt einen neuen Kalendertermin",
        inputSchema: {
          type: "object",
          properties: {
            subject: { type: "string", description: "Titel" },
            start: { type: "string", description: "Startzeit (ISO 8601 UTC)" },
            end: { type: "string", description: "Endzeit (ISO 8601 UTC)" },
            body: { type: "string", description: "Beschreibung" },
            location: { type: "string", description: "Ort" },
            attendees: { type: "array", items: { type: "string" }, description: "Teilnehmer-E-Mails" },
            isAllDay: { type: "boolean", description: "Ganztägiger Termin" },
          },
          required: ["subject", "start", "end"],
        },
      },
      {
        name: "update_event",
        description: "Aktualisiert einen bestehenden Kalendertermin",
        inputSchema: {
          type: "object",
          properties: {
            id: { type: "string", description: "Termin-ID" },
            subject: { type: "string" },
            start: { type: "string", description: "Neue Startzeit (ISO 8601 UTC)" },
            end: { type: "string", description: "Neue Endzeit (ISO 8601 UTC)" },
            body: { type: "string" },
            location: { type: "string" },
          },
          required: ["id"],
        },
      },
      {
        name: "delete_event",
        description: "Löscht einen Kalendertermin",
        inputSchema: {
          type: "object",
          properties: {
            id: { type: "string", description: "Termin-ID" },
          },
          required: ["id"],
        },
      },
    ],
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      let result: unknown;

      switch (name) {
        case "list_emails":     result = await listEmails(args as any); break;
        case "read_email":      result = await readEmail((args as any).id); break;
        case "send_email":      result = await sendEmail(args as any); break;
        case "reply_to_email":  result = await replyToEmail(args as any); break;
        case "list_events":     result = await listEvents(args as any); break;
        case "get_event":       result = await getEvent((args as any).id); break;
        case "create_event":    result = await createEvent(args as any); break;
        case "update_event":    result = await updateEvent(args as any); break;
        case "delete_event":    result = await deleteEvent((args as any).id); break;
        default:
          throw new McpError(ErrorCode.MethodNotFound, `Unbekanntes Tool: ${name}`);
      }

      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (err) {
      if (err instanceof McpError) throw err;
      return {
        content: [{ type: "text", text: `Fehler: ${(err as Error).message}` }],
        isError: true,
      };
    }
  });

  return server;
}

// Express HTTP-Server
const app = express();
app.use(cors());
app.use(express.json());

// StreamableHTTP Transport (neueres Protokoll)
app.post("/mcp", async (req, res) => {
  console.log("POST /mcp - body:", JSON.stringify(req.body));
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
  });
  const server = createMcpServer();
  await server.connect(transport);
  await transport.handleRequest(req, res, req.body);
});

app.get("/mcp", (_req, res) => {
  res.json({ status: "ok", service: "mcp-outlook" });
});

// SSE Transport (älteres Protokoll)
const sseTransports = new Map<string, SSEServerTransport>();

app.get("/sse", async (req, res) => {
  console.log("GET /sse - neue Verbindung");
  const transport = new SSEServerTransport("/messages", res);
  sseTransports.set(transport.sessionId, transport);
  res.on("close", () => sseTransports.delete(transport.sessionId));
  const server = createMcpServer();
  await server.connect(transport);
});

app.post("/messages", async (req, res) => {
  const sessionId = req.query.sessionId as string;
  const transport = sseTransports.get(sessionId);
  if (!transport) {
    res.status(404).json({ error: "Session nicht gefunden" });
    return;
  }
  await transport.handlePostMessage(req, res);
});

app.get("/health", (_req, res) => {
  res.json({ status: "ok", service: "mcp-outlook" });
});

const PORT = parseInt(process.env.PORT ?? "3000");
app.listen(PORT, () => {
  console.log(`MCP Outlook Server läuft auf Port ${PORT}`);
});
