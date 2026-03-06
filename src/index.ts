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
import { listTodoLists, listTasks, createTask, updateTask, deleteTask } from "./tools/todo.js";
import { listSharepointSites, listSharepointFiles, searchSharepoint } from "./tools/sharepoint.js";
import { listOneDriveFiles, searchOneDrive, getOneDriveFileInfo } from "./tools/onedrive.js";

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
      // ── Microsoft To Do ──────────────────────────────────────────────
      {
        name: "list_todo_lists",
        description: "Listet alle Microsoft To Do Aufgabenlisten auf",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "list_tasks",
        description: "Listet Aufgaben einer To Do Liste auf",
        inputSchema: {
          type: "object",
          properties: {
            listId: { type: "string", description: "ID der Aufgabenliste" },
            filter: { type: "string", description: "OData-Filter, z.B. 'status eq \\'notStarted\\''" },
            top: { type: "number", description: "Anzahl der Aufgaben (Standard: 20)" },
          },
          required: ["listId"],
        },
      },
      {
        name: "create_task",
        description: "Erstellt eine neue Aufgabe in einer To Do Liste",
        inputSchema: {
          type: "object",
          properties: {
            listId: { type: "string", description: "ID der Aufgabenliste" },
            title: { type: "string", description: "Titel der Aufgabe" },
            body: { type: "string", description: "Beschreibung" },
            dueDateTime: { type: "string", description: "Fälligkeitsdatum (ISO 8601 UTC)" },
            importance: { type: "string", enum: ["low", "normal", "high"], description: "Priorität" },
          },
          required: ["listId", "title"],
        },
      },
      {
        name: "update_task",
        description: "Aktualisiert eine Aufgabe (z.B. als erledigt markieren)",
        inputSchema: {
          type: "object",
          properties: {
            listId: { type: "string", description: "ID der Aufgabenliste" },
            taskId: { type: "string", description: "ID der Aufgabe" },
            title: { type: "string" },
            status: { type: "string", enum: ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"] },
            importance: { type: "string", enum: ["low", "normal", "high"] },
            dueDateTime: { type: "string", description: "Fälligkeitsdatum (ISO 8601 UTC)" },
            body: { type: "string" },
          },
          required: ["listId", "taskId"],
        },
      },
      {
        name: "delete_task",
        description: "Löscht eine Aufgabe aus einer To Do Liste",
        inputSchema: {
          type: "object",
          properties: {
            listId: { type: "string", description: "ID der Aufgabenliste" },
            taskId: { type: "string", description: "ID der Aufgabe" },
          },
          required: ["listId", "taskId"],
        },
      },
      // ── SharePoint ───────────────────────────────────────────────────
      {
        name: "list_sharepoint_sites",
        description: "Listet SharePoint-Seiten auf oder sucht nach einer bestimmten",
        inputSchema: {
          type: "object",
          properties: {
            search: { type: "string", description: "Suchbegriff für Site-Name (leer = alle)" },
            top: { type: "number", description: "Anzahl der Ergebnisse (Standard: 10)" },
          },
        },
      },
      {
        name: "list_sharepoint_files",
        description: "Listet Dateien und Ordner in einer SharePoint-Dokumentbibliothek auf",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            driveId: { type: "string", description: "ID der Dokumentbibliothek (optional)" },
            folderId: { type: "string", description: "ID des Unterordners (optional)" },
            top: { type: "number", description: "Anzahl der Ergebnisse (Standard: 20)" },
          },
          required: ["siteId"],
        },
      },
      {
        name: "search_sharepoint",
        description: "Sucht nach Dateien und Inhalten in SharePoint und OneDrive",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Suchbegriff" },
            top: { type: "number", description: "Anzahl der Ergebnisse (Standard: 10)" },
          },
          required: ["query"],
        },
      },
      // ── OneDrive ─────────────────────────────────────────────────────
      {
        name: "list_onedrive_files",
        description: "Listet Dateien und Ordner in OneDrive auf",
        inputSchema: {
          type: "object",
          properties: {
            folderId: { type: "string", description: "ID des Ordners (leer = Root)" },
            top: { type: "number", description: "Anzahl der Ergebnisse (Standard: 20)" },
          },
        },
      },
      {
        name: "search_onedrive",
        description: "Sucht nach Dateien in OneDrive",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Suchbegriff (Dateiname oder Inhalt)" },
            top: { type: "number", description: "Anzahl der Ergebnisse (Standard: 20)" },
          },
          required: ["query"],
        },
      },
      {
        name: "get_onedrive_file_info",
        description: "Ruft Details und Download-Link einer OneDrive-Datei ab",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "ID der Datei" },
          },
          required: ["fileId"],
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
        case "delete_event":          result = await deleteEvent((args as any).id); break;
        // To Do
        case "list_todo_lists":       result = await listTodoLists(); break;
        case "list_tasks":            result = await listTasks(args as any); break;
        case "create_task":           result = await createTask(args as any); break;
        case "update_task":           result = await updateTask(args as any); break;
        case "delete_task":           result = await deleteTask(args as any); break;
        // SharePoint
        case "list_sharepoint_sites": result = await listSharepointSites(args as any); break;
        case "list_sharepoint_files": result = await listSharepointFiles(args as any); break;
        case "search_sharepoint":     result = await searchSharepoint(args as any); break;
        // OneDrive
        case "list_onedrive_files":   result = await listOneDriveFiles(args as any); break;
        case "search_onedrive":       result = await searchOneDrive(args as any); break;
        case "get_onedrive_file_info":result = await getOneDriveFileInfo(args as any); break;
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

app.use((req, _res, next) => {
  console.log(`${new Date().toISOString()} ${req.method} ${req.path}`);
  next();
});

// API Key Authentifizierung (außer /health)
app.use((req, res, next) => {
  if (req.path === "/health") return next();

  const apiKey = process.env.API_KEY;
  if (!apiKey) return next(); // Kein Key konfiguriert → offen

  const authHeader = req.headers["authorization"];
  const queryKey = req.query["apikey"];

  const validHeader = authHeader === `Bearer ${apiKey}`;
  const validQuery = queryKey === apiKey;

  if (!validHeader && !validQuery) {
    res.status(401).json({ error: "Unauthorized" });
    return;
  }
  next();
});

// StreamableHTTP Transport (neueres Protokoll)
app.post("/mcp", async (req, res) => {
  console.log("POST /mcp - body:", JSON.stringify(req.body));
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
  });
  const server = createMcpServer();
  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error("POST /mcp - Fehler:", err);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        id: req.body?.id ?? null,
        error: { code: -32603, message: (err as Error).message },
      });
    }
  } finally {
    await server.close();
  }
});

app.get("/mcp", (_req, res) => {
  res.status(405).json({ error: "Method Not Allowed. Use POST for MCP requests." });
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
