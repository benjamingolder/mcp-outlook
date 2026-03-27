import "dotenv/config";
import express from "express";
import cors from "cors";
import { readFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { Client } from "@microsoft/microsoft-graph-client";
import { entraAuthMiddleware } from "./entraAuth.js";
import { getGraphTokenViaObo } from "./obo.js";
import { getGraphClient } from "./graph.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
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
import {
  listSharepointSites, getSharepointSite,
  listSharepointFiles, searchSharepoint,
  listSharepointLists, getSharepointList, createSharepointList, updateSharepointList, deleteSharepointList,
  listSharepointListItems, getSharepointListItem,
  createSharepointListItem, updateSharepointListItem, deleteSharepointListItem,
  createSharepointFolder, uploadSharepointFile, deleteSharepointFile, moveSharepointFile,
} from "./tools/sharepoint.js";
import {
  listOneDriveFiles, searchOneDrive, getOneDriveFileInfo,
  createOneDriveFolder, uploadOneDriveFile, deleteOneDriveItem,
  moveOneDriveItem, renameOneDriveItem, copyOneDriveItem,
} from "./tools/onedrive.js";
import { listContacts, getContact, createContact, updateContact, deleteContact } from "./tools/contacts.js";
import { listTeams, listChannels, listChannelMessages, sendChannelMessage, listChats, listChatMessages, sendChatMessage } from "./tools/teams.js";
import { listNotebooks, listSections, listPages, getPage, createPage } from "./tools/onenote.js";
import { listPlans, listMyPlannerTasks, listBuckets, listPlanTasks, createPlannerTask, updatePlannerTask, deletePlannerTask } from "./tools/planner.js";
import { listWorksheets, getRange, updateRange, getUsedRange } from "./tools/excel.js";
import { listRelevantPeople, listTrendingDocuments, listUsedDocuments, listSharedDocuments } from "./tools/people.js";
import { listUsers, getUser, listGroups, listGroupMembers, addGroupMember, removeGroupMember } from "./tools/directory.js";
import { getMyPresence, getUserPresence, getPresenceForUsers, setMyPresence } from "./tools/presence.js";
import { listBookingBusinesses, listBookingServices, listBookingAppointments, createBookingAppointment, cancelBookingAppointment } from "./tools/bookings.js";

function createMcpServer(client: Client): Server {
  const server = new Server(
    {
      name: "outlook-mcp",
      version: "1.0.0",
      title: "Microsoft 365",
      description: "MCP Server für Microsoft 365 (Mail, Kalender, Teams, SharePoint, OneDrive u.v.m.)",
      icons: [
        {
          src: "data:image/svg+xml;base64,PHN2ZyB2aWV3Qm94PSIwIC05LjIgOTYwIDEwNzQuNTAwMDAwMDAwMDAwMiIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB3aWR0aD0iMjIxNSIgaGVpZ2h0PSIyNTAwIj48cmFkaWFsR3JhZGllbnQgaWQ9ImEiIGN4PSIzMjIiIGN5PSIyMDcuMyIgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2UiIHI9IjgwMC44Ij48c3RvcCBvZmZzZXQ9Ii4wNjQiIHN0b3AtY29sb3I9IiNhZTdmZTIiLz48c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDc4ZDQiLz48L3JhZGlhbEdyYWRpZW50PjxsaW5lYXJHcmFkaWVudCBpZD0iYiIgZ3JhZGllbnRVbml5cz0idXNlclNwYWNlT25Vc2UiIHgxPSIzMjQuMyIgeDI9IjIxMCIgeTE9Ijg2MC44IiB5Mj0iNjYzLjIiPjxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iIzExNGE4YiIvPjxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwNzhkNCIgc3RvcC1vcGFjaXR5PSIwIi8+PC9saW5lYXJHcmFkaWVudD48cmFkaWFsR3JhZGllbnQgaWQ9ImMiIGN4PSIxNTQuMyIgY3k9IjgyNC40IiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIgcj0iNzQ1LjIiPjxzdG9wIG9mZnNldD0iLjEzNCIgc3RvcC1jb2xvcj0iI2Q1OWRmZiIvPjxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzVlNDM4ZiIvPjwvcmFkaWFsR3JhZGllbnQ+PGxpbmVhckdyYWRpZW50IGlkPSJkIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIgeDE9Ijg3Mi42IiB4Mj0iNzUwLjEiIHkxPSI1NjEiIHkyPSI3MzYuNiI+PHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjNDkzNDc0Ii8+PHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjOGM2NmJhIiBzdG9wLW9wYWNpdHk9IjAiLz48L2xpbmVhckdyYWRpZW50PjxyYWRpYWxHcmFkaWVudCBpZD0iZSIgY3g9Ijg4OS4zIiBjeT0iNTg4LjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiByPSI1OTguMSI+PHN0b3Agb2Zmc2V0PSIuMDU4IiBzdG9wLWNvbG9yPSIjNTBlNmZmIi8+PHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjNDM2ZGNkIi8+PC9yYWRpYWxHcmFkaWVudD48bGluZWFyR3JhZGllbnQgaWQ9ImYiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiB4MT0iMzExLjQiIHgyPSI0OTEuNyIgeTE9IjI1LjQiIHkyPSIyNS40Ij48c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiMyZDNmODAiLz48c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiM0MzZkY2QiIHN0b3Atb3BhY2l0eT0iMCIvPjwvbGluZWFyR3JhZGllbnQ+PHBhdGggZD0iTTM4NiAyNC42bC01LjQgMy4zYy04LjUgNS4yLTE2LjYgMTEtMjQuMiAxNy4zTDM3MiAzNC4zaDEzMkw1MjggMjE2IDQwOCAzMzZsLTEyMCA4My40djk2LjJjMCA2Ny4yIDM1LjEgMTI5LjUgOTIuNiAxNjQuMmwxMjYuMyA3Ni41TDI0MCA5MTJoLTUxLjVsLTk1LjktNTguMUMzNS4xIDgxOS4xIDAgNzU2LjkgMCA2ODkuN1YzNjYuM0MwIDI5OS4xIDM1LjEgMjM2LjcgOTIuNiAyMDJsMjg4LTE3NC4ycTIuNy0xLjcgNS40LTMuMnoiIGZpbGw9InVybCgjYSkiLz48cGF0aCBkPSJNMzg2IDI0LjZsLTUuNCAzLjNjLTguNSA1LjItMTYuNiAxMS0yNC4yIDE3LjNMMzcyIDM0LjNoMTMyTDUyOCAyMTYgNDA4IDMzNmwtMTIwIDgzLjR2OTYuMmMwIDY3LjIgMzUuMSAxMjkuNSA5Mi42IDE2NC4ybDEyNi4zIDc2LjVMMjQwIDkxMmgtNTEuNWwtOTUuOS01OC4xQzM1LjEgODE5LjEgMCA3NTYuOSAwIDY4OS43VjM2Ni4zQzAgMjk5LjEgMzUuMSAyMzYuNyA5Mi42IDIwMmwyODgtMTc0LjJxMi43LTEuNyA1LjQtMy4yeiIgZmlsbD0idXJsKCNiKSIvPjxwYXRoIGQ9Ik05MzYgNTc2bDI0IDM2djc3LjdjMCA2Ny4xLTM1LjEgMTI5LjQtOTIuNiAxNjQuMmwtMjg4IDE3NC40Yy02MS4xIDM3LTEzNy43IDM3LTE5OC44IDBMOTkuMyA4NThjNTkuOSAzMy4xIDEzMy4yIDMxLjggMTkyLjEtMy45bDI4OC0xNzQuM0M2MzYuOSA2NDUgNjcyIDU4Mi43IDY3MiA1MTUuNVY0MDh6IiBmaWxsPSJ1cmwoI2MpIi8+PHBhdGggZD0iTTkzNiA1NzZsMjQgMzZ2NzcuN2MwIDY3LjEtMzUuMSAxMjkuNC05Mi42IDE2NC4ybC0yODggMTc0LjRjLTYxLjEgMzctMTM3LjcgMzctMTk4LjggMEw5OS4zIDg1OGM1OS45IDMzLjEgMTMzLjIgMzEuOCAxOTIuMS0zLjlsMjg4LTE3NC4zQzYzNi45IDY0NSA2NzIgNTgyLjcgNjcyIDUxNS41VjQwOHoiIGZpbGw9InVybCgjZCkiLz48cGF0aCBkPSJNOTYwIDM2Ni4zdjMyMy40cTAgMy4xLS4xIDYuM2MtMi4xLTY0LjgtMzYuOC0xMjQuMy05Mi41LTE1OGwtMjg4LTE3NC4yYy02MS4xLTM3LTEzNy43LTM3LTE5OC44IDBsLTkyLjYgNTZWMTkyLjJjMC02Ny4yIDM1LjEtMTI5LjUgOTIuNi0xNjQuM2w1LjctMy41QzQ0Ni41LTkuMiA1MjAuMi04IDU3OS40IDI3LjhsMjg4IDE3NC4yYzU3LjUgMzQuNyA5Mi42IDk3LjEgOTIuNiAxNjQuM3oiIGZpbGw9InVybCgjZSkiLz48cGF0aCBkPSJNOTYwIDM2Ni4zdjMyMy40cTAgMy4xLS4xIDYuM2MtMi4xLTY0LjgtMzYuOC0xMjQuMy05Mi41LTE1OGwtMjg4LTE3NC4yYy02MS4xLTM3LTEzNy43LTM3LTE5OC44IDBsLTkyLjYgNTZWMTkyLjJjMC02Ny4yIDM1LjEtMTI5LjUgOTIuNi0xNjQuM2w1LjctMy41QzQ0Ni41LTkuMiA1MjAuMiA4IDU3OS40IDI3LjhsMjg4IDE3NC4yYzU3LjUgMzQuNyA5Mi42IDk3LjEgOTIuNiAxNjQuM3oiIGZpbGw9InVybCgjZikiLz48L3N2Zz4=",
          mimeType: "image/svg+xml",
          sizes: ["any"],
        },
      ],
    },
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
            categories: { type: "array", items: { type: "string" }, description: "Kategorien (z.B. ['Wichtig', 'Privat'])" },
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
            categories: { type: "array", items: { type: "string" }, description: "Kategorien (z.B. ['Wichtig', 'Privat'])" },
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
      {
        name: "list_sharepoint_lists",
        description: "Listet alle Listen einer SharePoint-Site auf",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            top: { type: "number", description: "Anzahl der Ergebnisse (Standard: 20)" },
          },
          required: ["siteId"],
        },
      },
      {
        name: "create_sharepoint_list",
        description: "Erstellt eine neue SharePoint-Liste mit benutzerdefinierten Spalten",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            displayName: { type: "string", description: "Name der Liste" },
            description: { type: "string", description: "Beschreibung der Liste" },
            columns: {
              type: "array",
              description: "Spaltendefinitionen",
              items: {
                type: "object",
                properties: {
                  name: { type: "string" },
                  type: { type: "string", enum: ["text", "number", "boolean", "dateTime", "choice"] },
                  choices: { type: "array", items: { type: "string" }, description: "Nur bei Typ 'choice'" },
                },
                required: ["name", "type"],
              },
            },
          },
          required: ["siteId", "displayName", "columns"],
        },
      },
      {
        name: "list_sharepoint_list_items",
        description: "Listet Einträge einer SharePoint-Liste auf",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            listId: { type: "string", description: "ID oder Name der Liste" },
            top: { type: "number", description: "Anzahl der Ergebnisse (Standard: 20)" },
            filter: { type: "string", description: "OData-Filter" },
          },
          required: ["siteId", "listId"],
        },
      },
      {
        name: "get_sharepoint_list_item",
        description: "Liest einen einzelnen Eintrag einer SharePoint-Liste",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string" },
            listId: { type: "string" },
            itemId: { type: "string" },
          },
          required: ["siteId", "listId", "itemId"],
        },
      },
      {
        name: "create_sharepoint_list_item",
        description: "Erstellt einen neuen Eintrag in einer SharePoint-Liste",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string" },
            listId: { type: "string" },
            fields: { type: "object", description: "Feldinhalte als Key-Value Objekt" },
          },
          required: ["siteId", "listId", "fields"],
        },
      },
      {
        name: "update_sharepoint_list_item",
        description: "Aktualisiert einen Eintrag in einer SharePoint-Liste",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string" },
            listId: { type: "string" },
            itemId: { type: "string" },
            fields: { type: "object", description: "Zu aktualisierende Felder als Key-Value Objekt" },
          },
          required: ["siteId", "listId", "itemId", "fields"],
        },
      },
      {
        name: "delete_sharepoint_list_item",
        description: "Löscht einen Eintrag aus einer SharePoint-Liste",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string" },
            listId: { type: "string" },
            itemId: { type: "string" },
          },
          required: ["siteId", "listId", "itemId"],
        },
      },
      {
        name: "get_sharepoint_site",
        description: "Ruft Details einer SharePoint-Site ab",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
          },
          required: ["siteId"],
        },
      },
      {
        name: "get_sharepoint_list",
        description: "Ruft Details einer SharePoint-Liste ab",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            listId: { type: "string", description: "ID der Liste" },
          },
          required: ["siteId", "listId"],
        },
      },
      {
        name: "update_sharepoint_list",
        description: "Aktualisiert eine SharePoint-Liste (Name, Beschreibung)",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            listId: { type: "string", description: "ID der Liste" },
            displayName: { type: "string", description: "Neuer Anzeigename" },
            description: { type: "string", description: "Neue Beschreibung" },
          },
          required: ["siteId", "listId"],
        },
      },
      {
        name: "delete_sharepoint_list",
        description: "Löscht eine SharePoint-Liste",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            listId: { type: "string", description: "ID der Liste" },
          },
          required: ["siteId", "listId"],
        },
      },
      {
        name: "create_sharepoint_folder",
        description: "Erstellt einen Ordner in einer SharePoint-Dokumentbibliothek",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            driveId: { type: "string", description: "ID der Drive (optional)" },
            parentId: { type: "string", description: "ID des übergeordneten Ordners (leer = Root)" },
            folderName: { type: "string", description: "Name des neuen Ordners" },
          },
          required: ["siteId", "folderName"],
        },
      },
      {
        name: "upload_sharepoint_file",
        description: "Lädt eine Datei in eine SharePoint-Dokumentbibliothek hoch",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            driveId: { type: "string", description: "ID der Drive (optional)" },
            parentId: { type: "string", description: "ID des Zielordners (leer = Root)" },
            fileName: { type: "string", description: "Dateiname inkl. Endung" },
            content: { type: "string", description: "Dateiinhalt als Text" },
            mimeType: { type: "string", description: "MIME-Type (Standard: text/plain)" },
          },
          required: ["siteId", "fileName", "content"],
        },
      },
      {
        name: "delete_sharepoint_file",
        description: "Löscht eine Datei oder einen Ordner in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            driveId: { type: "string", description: "ID der Drive (optional)" },
            itemId: { type: "string", description: "ID der Datei oder des Ordners" },
          },
          required: ["siteId", "itemId"],
        },
      },
      {
        name: "move_sharepoint_file",
        description: "Verschiebt eine Datei oder einen Ordner in SharePoint (optional umbenennen)",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "ID der SharePoint-Site" },
            driveId: { type: "string", description: "ID der Drive (optional)" },
            itemId: { type: "string", description: "ID der Datei oder des Ordners" },
            destinationParentId: { type: "string", description: "ID des Zielordners" },
            newName: { type: "string", description: "Neuer Name (optional)" },
          },
          required: ["siteId", "itemId", "destinationParentId"],
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
      {
        name: "create_onedrive_folder",
        description: "Erstellt einen neuen Ordner in OneDrive",
        inputSchema: {
          type: "object",
          properties: {
            parentId: { type: "string", description: "ID des übergeordneten Ordners (leer = Root)" },
            folderName: { type: "string", description: "Name des neuen Ordners" },
          },
          required: ["folderName"],
        },
      },
      {
        name: "upload_onedrive_file",
        description: "Lädt eine Datei in OneDrive hoch (erstellt oder überschreibt)",
        inputSchema: {
          type: "object",
          properties: {
            parentId: { type: "string", description: "ID des Zielordners (leer = Root)" },
            fileName: { type: "string", description: "Dateiname inkl. Endung" },
            content: { type: "string", description: "Dateiinhalt als Text" },
            mimeType: { type: "string", description: "MIME-Type (Standard: text/plain)" },
          },
          required: ["fileName", "content"],
        },
      },
      {
        name: "delete_onedrive_item",
        description: "Löscht eine Datei oder einen Ordner in OneDrive",
        inputSchema: {
          type: "object",
          properties: {
            itemId: { type: "string", description: "ID der Datei oder des Ordners" },
          },
          required: ["itemId"],
        },
      },
      {
        name: "move_onedrive_item",
        description: "Verschiebt eine Datei oder einen Ordner in OneDrive (optional umbenennen)",
        inputSchema: {
          type: "object",
          properties: {
            itemId: { type: "string", description: "ID der Datei oder des Ordners" },
            destinationParentId: { type: "string", description: "ID des Zielordners" },
            newName: { type: "string", description: "Neuer Name (optional)" },
          },
          required: ["itemId", "destinationParentId"],
        },
      },
      {
        name: "rename_onedrive_item",
        description: "Benennt eine Datei oder einen Ordner in OneDrive um",
        inputSchema: {
          type: "object",
          properties: {
            itemId: { type: "string", description: "ID der Datei oder des Ordners" },
            newName: { type: "string", description: "Neuer Name" },
          },
          required: ["itemId", "newName"],
        },
      },
      {
        name: "copy_onedrive_item",
        description: "Kopiert eine Datei oder einen Ordner in OneDrive",
        inputSchema: {
          type: "object",
          properties: {
            itemId: { type: "string", description: "ID der Datei oder des Ordners" },
            destinationParentId: { type: "string", description: "ID des Zielordners" },
            newName: { type: "string", description: "Name der Kopie (optional)" },
          },
          required: ["itemId", "destinationParentId"],
        },
      },
      // ── Contacts ─────────────────────────────────────────────────────
      {
        name: "list_contacts",
        description: "Listet Outlook-Kontakte auf",
        inputSchema: {
          type: "object",
          properties: {
            top: { type: "number", description: "Anzahl (Standard: 20)" },
            filter: { type: "string", description: "OData-Filter" },
            search: { type: "string", description: "Suchbegriff (Name/Email)" },
          },
        },
      },
      {
        name: "get_contact",
        description: "Liest einen Kontakt",
        inputSchema: { type: "object", properties: { id: { type: "string" } }, required: ["id"] },
      },
      {
        name: "create_contact",
        description: "Erstellt einen neuen Outlook-Kontakt",
        inputSchema: {
          type: "object",
          properties: {
            givenName: { type: "string", description: "Vorname" },
            surname: { type: "string", description: "Nachname" },
            emailAddresses: { type: "array", items: { type: "object", properties: { address: { type: "string" }, name: { type: "string" } } } },
            mobilePhone: { type: "string" },
            businessPhones: { type: "array", items: { type: "string" } },
            jobTitle: { type: "string" },
            companyName: { type: "string" },
            department: { type: "string" },
          },
          required: ["givenName"],
        },
      },
      {
        name: "update_contact",
        description: "Aktualisiert einen Outlook-Kontakt",
        inputSchema: {
          type: "object",
          properties: {
            id: { type: "string" },
            givenName: { type: "string" }, surname: { type: "string" },
            emailAddresses: { type: "array", items: { type: "object" } },
            mobilePhone: { type: "string" }, jobTitle: { type: "string" },
            companyName: { type: "string" }, department: { type: "string" },
            personalNotes: { type: "string" },
          },
          required: ["id"],
        },
      },
      {
        name: "delete_contact",
        description: "Löscht einen Outlook-Kontakt",
        inputSchema: { type: "object", properties: { id: { type: "string" } }, required: ["id"] },
      },
      // ── Teams ─────────────────────────────────────────────────────────
      {
        name: "list_teams",
        description: "Listet alle Teams auf, denen du angehörst",
        inputSchema: { type: "object", properties: { top: { type: "number" } } },
      },
      {
        name: "list_channels",
        description: "Listet Kanäle eines Teams auf",
        inputSchema: { type: "object", properties: { teamId: { type: "string" } }, required: ["teamId"] },
      },
      {
        name: "list_channel_messages",
        description: "Liest Nachrichten aus einem Teams-Kanal",
        inputSchema: {
          type: "object",
          properties: {
            teamId: { type: "string" }, channelId: { type: "string" },
            top: { type: "number", description: "Anzahl (Standard: 20)" },
          },
          required: ["teamId", "channelId"],
        },
      },
      {
        name: "send_channel_message",
        description: "Sendet eine Nachricht in einen Teams-Kanal",
        inputSchema: {
          type: "object",
          properties: {
            teamId: { type: "string" }, channelId: { type: "string" },
            content: { type: "string" },
            contentType: { type: "string", enum: ["text", "html"] },
            subject: { type: "string" },
          },
          required: ["teamId", "channelId", "content"],
        },
      },
      {
        name: "list_chats",
        description: "Listet Teams-Chats auf",
        inputSchema: { type: "object", properties: { top: { type: "number" } } },
      },
      {
        name: "list_chat_messages",
        description: "Liest Nachrichten aus einem Teams-Chat",
        inputSchema: {
          type: "object",
          properties: { chatId: { type: "string" }, top: { type: "number" } },
          required: ["chatId"],
        },
      },
      {
        name: "send_chat_message",
        description: "Sendet eine Nachricht in einen Teams-Chat",
        inputSchema: {
          type: "object",
          properties: {
            chatId: { type: "string" }, content: { type: "string" },
            contentType: { type: "string", enum: ["text", "html"] },
          },
          required: ["chatId", "content"],
        },
      },
      // ── OneNote ───────────────────────────────────────────────────────
      {
        name: "list_notebooks",
        description: "Listet OneNote-Notizbücher auf",
        inputSchema: { type: "object", properties: { top: { type: "number" } } },
      },
      {
        name: "list_sections",
        description: "Listet Abschnitte eines Notizbuchs auf",
        inputSchema: { type: "object", properties: { notebookId: { type: "string" } }, required: ["notebookId"] },
      },
      {
        name: "list_pages",
        description: "Listet Seiten eines OneNote-Abschnitts auf",
        inputSchema: {
          type: "object",
          properties: { sectionId: { type: "string" }, top: { type: "number" } },
          required: ["sectionId"],
        },
      },
      {
        name: "get_page",
        description: "Liest den HTML-Inhalt einer OneNote-Seite",
        inputSchema: { type: "object", properties: { pageId: { type: "string" } }, required: ["pageId"] },
      },
      {
        name: "create_page",
        description: "Erstellt eine neue OneNote-Seite",
        inputSchema: {
          type: "object",
          properties: {
            sectionId: { type: "string" }, title: { type: "string" },
            content: { type: "string", description: "HTML-Inhalt der Seite" },
          },
          required: ["sectionId", "title"],
        },
      },
      // ── Planner ───────────────────────────────────────────────────────
      {
        name: "list_my_planner_tasks",
        description: "Listet alle Planner-Aufgaben auf, die dir zugewiesen sind",
        inputSchema: { type: "object", properties: { top: { type: "number" } } },
      },
      {
        name: "list_plans",
        description: "Listet Pläne einer Microsoft 365 Gruppe auf",
        inputSchema: { type: "object", properties: { groupId: { type: "string" } }, required: ["groupId"] },
      },
      {
        name: "list_buckets",
        description: "Listet Buckets (Spalten) eines Planner-Plans auf",
        inputSchema: { type: "object", properties: { planId: { type: "string" } }, required: ["planId"] },
      },
      {
        name: "list_plan_tasks",
        description: "Listet alle Aufgaben eines Planner-Plans auf",
        inputSchema: { type: "object", properties: { planId: { type: "string" } }, required: ["planId"] },
      },
      {
        name: "create_planner_task",
        description: "Erstellt eine neue Planner-Aufgabe",
        inputSchema: {
          type: "object",
          properties: {
            planId: { type: "string" }, title: { type: "string" },
            bucketId: { type: "string" }, dueDateTime: { type: "string" },
            assignedToUserIds: { type: "array", items: { type: "string" } },
            priority: { type: "number", description: "0 (dringend) bis 9 (unwichtig)" },
          },
          required: ["planId", "title"],
        },
      },
      {
        name: "update_planner_task",
        description: "Aktualisiert eine Planner-Aufgabe",
        inputSchema: {
          type: "object",
          properties: {
            taskId: { type: "string" }, title: { type: "string" },
            percentComplete: { type: "number", description: "0, 50 oder 100" },
            dueDateTime: { type: "string" }, priority: { type: "number" }, bucketId: { type: "string" },
          },
          required: ["taskId"],
        },
      },
      {
        name: "delete_planner_task",
        description: "Löscht eine Planner-Aufgabe",
        inputSchema: { type: "object", properties: { taskId: { type: "string" } }, required: ["taskId"] },
      },
      // ── Excel ─────────────────────────────────────────────────────────
      {
        name: "list_worksheets",
        description: "Listet Tabellenblätter einer Excel-Datei auf",
        inputSchema: {
          type: "object",
          properties: { fileId: { type: "string" }, driveId: { type: "string", description: "Optional: Drive-ID (z.B. SharePoint)" } },
          required: ["fileId"],
        },
      },
      {
        name: "get_range",
        description: "Liest einen Zellenbereich aus einer Excel-Datei",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string" }, worksheetId: { type: "string" },
            address: { type: "string", description: "z.B. A1:C10" }, driveId: { type: "string" },
          },
          required: ["fileId", "worksheetId", "address"],
        },
      },
      {
        name: "get_used_range",
        description: "Liest den gesamten benutzten Bereich eines Tabellenblatts",
        inputSchema: {
          type: "object",
          properties: { fileId: { type: "string" }, worksheetId: { type: "string" }, driveId: { type: "string" } },
          required: ["fileId", "worksheetId"],
        },
      },
      {
        name: "update_range",
        description: "Schreibt Werte in einen Zellenbereich einer Excel-Datei",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string" }, worksheetId: { type: "string" },
            address: { type: "string" },
            values: { type: "array", items: { type: "array" }, description: "2D-Array mit Zellenwerten" },
            driveId: { type: "string" },
          },
          required: ["fileId", "worksheetId", "address", "values"],
        },
      },
      // ── People & Insights ─────────────────────────────────────────────
      {
        name: "list_relevant_people",
        description: "Listet relevante Personen basierend auf deiner Kommunikation auf",
        inputSchema: { type: "object", properties: { top: { type: "number" }, search: { type: "string" } } },
      },
      {
        name: "list_trending_documents",
        description: "Listet Dokumente auf, die gerade in deinem Umfeld trending sind",
        inputSchema: { type: "object", properties: { top: { type: "number" } } },
      },
      {
        name: "list_used_documents",
        description: "Listet zuletzt verwendete Dokumente auf",
        inputSchema: { type: "object", properties: { top: { type: "number" } } },
      },
      {
        name: "list_shared_documents",
        description: "Listet Dokumente auf, die mit dir geteilt wurden",
        inputSchema: { type: "object", properties: { top: { type: "number" } } },
      },
      // ── Directory ─────────────────────────────────────────────────────
      {
        name: "list_users",
        description: "Listet Benutzer im Tenant auf",
        inputSchema: {
          type: "object",
          properties: { top: { type: "number" }, filter: { type: "string" }, search: { type: "string", description: "Suche nach Displayname" } },
        },
      },
      {
        name: "get_user",
        description: "Liest Details eines Benutzers",
        inputSchema: { type: "object", properties: { userId: { type: "string", description: "ID oder UPN" } }, required: ["userId"] },
      },
      {
        name: "list_groups",
        description: "Listet Gruppen im Tenant auf",
        inputSchema: {
          type: "object",
          properties: { top: { type: "number" }, filter: { type: "string" }, search: { type: "string" } },
        },
      },
      {
        name: "list_group_members",
        description: "Listet Mitglieder einer Gruppe auf",
        inputSchema: {
          type: "object",
          properties: { groupId: { type: "string" }, top: { type: "number" } },
          required: ["groupId"],
        },
      },
      {
        name: "add_group_member",
        description: "Fügt einen Benutzer zu einer Gruppe hinzu",
        inputSchema: {
          type: "object",
          properties: { groupId: { type: "string" }, userId: { type: "string" } },
          required: ["groupId", "userId"],
        },
      },
      {
        name: "remove_group_member",
        description: "Entfernt einen Benutzer aus einer Gruppe",
        inputSchema: {
          type: "object",
          properties: { groupId: { type: "string" }, userId: { type: "string" } },
          required: ["groupId", "userId"],
        },
      },
      // ── Presence ──────────────────────────────────────────────────────
      {
        name: "get_my_presence",
        description: "Liest deinen eigenen Teams-Präsenzstatus",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "get_user_presence",
        description: "Liest den Präsenzstatus eines bestimmten Benutzers",
        inputSchema: { type: "object", properties: { userId: { type: "string" } }, required: ["userId"] },
      },
      {
        name: "get_presence_for_users",
        description: "Liest den Präsenzstatus mehrerer Benutzer auf einmal",
        inputSchema: {
          type: "object",
          properties: { userIds: { type: "array", items: { type: "string" } } },
          required: ["userIds"],
        },
      },
      {
        name: "set_my_presence",
        description: "Setzt deinen eigenen Teams-Präsenzstatus",
        inputSchema: {
          type: "object",
          properties: {
            availability: { type: "string", enum: ["Available", "Busy", "DoNotDisturb", "BeRightBack", "Away", "Offline"] },
            activity: { type: "string", description: "z.B. Available, InACall, InAMeeting, Away" },
            expirationDuration: { type: "string", description: "ISO 8601 Dauer, z.B. PT1H (Standard: 1 Stunde)" },
          },
          required: ["availability", "activity"],
        },
      },
      // ── Bookings ──────────────────────────────────────────────────────
      {
        name: "list_booking_businesses",
        description: "Listet alle Microsoft Bookings Unternehmen auf",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "list_booking_services",
        description: "Listet Services eines Bookings-Unternehmens auf",
        inputSchema: { type: "object", properties: { businessId: { type: "string" } }, required: ["businessId"] },
      },
      {
        name: "list_booking_appointments",
        description: "Listet Termine eines Bookings-Unternehmens auf",
        inputSchema: {
          type: "object",
          properties: {
            businessId: { type: "string" },
            start: { type: "string", description: "Von (ISO 8601)" },
            end: { type: "string", description: "Bis (ISO 8601)" },
          },
          required: ["businessId"],
        },
      },
      {
        name: "create_booking_appointment",
        description: "Erstellt einen neuen Bookings-Termin",
        inputSchema: {
          type: "object",
          properties: {
            businessId: { type: "string" }, serviceId: { type: "string" },
            startDateTime: { type: "string" }, endDateTime: { type: "string" },
            timeZone: { type: "string", description: "Standard: Europe/Berlin" },
            customerName: { type: "string" }, customerEmail: { type: "string" },
            customerPhone: { type: "string" },
            staffMemberIds: { type: "array", items: { type: "string" } },
            notes: { type: "string" },
          },
          required: ["businessId", "serviceId", "startDateTime", "endDateTime", "customerName", "customerEmail"],
        },
      },
      {
        name: "cancel_booking_appointment",
        description: "Storniert einen Bookings-Termin",
        inputSchema: {
          type: "object",
          properties: {
            businessId: { type: "string" }, appointmentId: { type: "string" },
            reason: { type: "string" },
          },
          required: ["businessId", "appointmentId"],
        },
      },
    ],
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      let result: unknown;

      switch (name) {
        case "list_emails":     result = await listEmails(client, args as any); break;
        case "read_email":      result = await readEmail(client, (args as any).id); break;
        case "send_email":      result = await sendEmail(client, args as any); break;
        case "reply_to_email":  result = await replyToEmail(client, args as any); break;
        case "list_events":     result = await listEvents(client, args as any); break;
        case "get_event":       result = await getEvent(client, (args as any).id); break;
        case "create_event":    result = await createEvent(client, args as any); break;
        case "update_event":    result = await updateEvent(client, args as any); break;
        case "delete_event":          result = await deleteEvent(client, (args as any).id); break;
        // To Do
        case "list_todo_lists":       result = await listTodoLists(client); break;
        case "list_tasks":            result = await listTasks(client, args as any); break;
        case "create_task":           result = await createTask(client, args as any); break;
        case "update_task":           result = await updateTask(client, args as any); break;
        case "delete_task":           result = await deleteTask(client, args as any); break;
        // SharePoint
        case "list_sharepoint_sites":          result = await listSharepointSites(client, args as any); break;
        case "get_sharepoint_site":            result = await getSharepointSite(client, args as any); break;
        case "list_sharepoint_files":          result = await listSharepointFiles(client, args as any); break;
        case "search_sharepoint":              result = await searchSharepoint(client, args as any); break;
        case "list_sharepoint_lists":          result = await listSharepointLists(client, args as any); break;
        case "get_sharepoint_list":            result = await getSharepointList(client, args as any); break;
        case "create_sharepoint_list":         result = await createSharepointList(client, args as any); break;
        case "update_sharepoint_list":         result = await updateSharepointList(client, args as any); break;
        case "delete_sharepoint_list":         result = await deleteSharepointList(client, args as any); break;
        case "list_sharepoint_list_items":     result = await listSharepointListItems(client, args as any); break;
        case "get_sharepoint_list_item":       result = await getSharepointListItem(client, args as any); break;
        case "create_sharepoint_list_item":    result = await createSharepointListItem(client, args as any); break;
        case "update_sharepoint_list_item":    result = await updateSharepointListItem(client, args as any); break;
        case "delete_sharepoint_list_item":    result = await deleteSharepointListItem(client, args as any); break;
        case "create_sharepoint_folder":       result = await createSharepointFolder(client, args as any); break;
        case "upload_sharepoint_file":         result = await uploadSharepointFile(client, args as any); break;
        case "delete_sharepoint_file":         result = await deleteSharepointFile(client, args as any); break;
        case "move_sharepoint_file":           result = await moveSharepointFile(client, args as any); break;
        // OneDrive
        case "list_onedrive_files":            result = await listOneDriveFiles(client, args as any); break;
        case "search_onedrive":                result = await searchOneDrive(client, args as any); break;
        case "get_onedrive_file_info":         result = await getOneDriveFileInfo(client, args as any); break;
        case "create_onedrive_folder":         result = await createOneDriveFolder(client, args as any); break;
        case "upload_onedrive_file":           result = await uploadOneDriveFile(client, args as any); break;
        case "delete_onedrive_item":           result = await deleteOneDriveItem(client, args as any); break;
        case "move_onedrive_item":             result = await moveOneDriveItem(client, args as any); break;
        case "rename_onedrive_item":           result = await renameOneDriveItem(client, args as any); break;
        case "copy_onedrive_item":             result = await copyOneDriveItem(client, args as any); break;
        // Contacts
        case "list_contacts":                  result = await listContacts(client, args as any); break;
        case "get_contact":                    result = await getContact(client, (args as any).id); break;
        case "create_contact":                 result = await createContact(client, args as any); break;
        case "update_contact":                 result = await updateContact(client, args as any); break;
        case "delete_contact":                 result = await deleteContact(client, (args as any).id); break;
        // Teams
        case "list_teams":                     result = await listTeams(client, args as any); break;
        case "list_channels":                  result = await listChannels(client, args as any); break;
        case "list_channel_messages":          result = await listChannelMessages(client, args as any); break;
        case "send_channel_message":           result = await sendChannelMessage(client, args as any); break;
        case "list_chats":                     result = await listChats(client, args as any); break;
        case "list_chat_messages":             result = await listChatMessages(client, args as any); break;
        case "send_chat_message":              result = await sendChatMessage(client, args as any); break;
        // OneNote
        case "list_notebooks":                 result = await listNotebooks(client, args as any); break;
        case "list_sections":                  result = await listSections(client, args as any); break;
        case "list_pages":                     result = await listPages(client, args as any); break;
        case "get_page":                       result = await getPage(client, args as any); break;
        case "create_page":                    result = await createPage(client, args as any); break;
        // Planner
        case "list_my_planner_tasks":          result = await listMyPlannerTasks(client, args as any); break;
        case "list_plans":                     result = await listPlans(client, args as any); break;
        case "list_buckets":                   result = await listBuckets(client, args as any); break;
        case "list_plan_tasks":                result = await listPlanTasks(client, args as any); break;
        case "create_planner_task":            result = await createPlannerTask(client, args as any); break;
        case "update_planner_task":            result = await updatePlannerTask(client, args as any); break;
        case "delete_planner_task":            result = await deletePlannerTask(client, args as any); break;
        // Excel
        case "list_worksheets":                result = await listWorksheets(client, args as any); break;
        case "get_range":                      result = await getRange(client, args as any); break;
        case "get_used_range":                 result = await getUsedRange(client, args as any); break;
        case "update_range":                   result = await updateRange(client, args as any); break;
        // People & Insights
        case "list_relevant_people":           result = await listRelevantPeople(client, args as any); break;
        case "list_trending_documents":        result = await listTrendingDocuments(client, args as any); break;
        case "list_used_documents":            result = await listUsedDocuments(client, args as any); break;
        case "list_shared_documents":          result = await listSharedDocuments(client, args as any); break;
        // Directory
        case "list_users":                     result = await listUsers(client, args as any); break;
        case "get_user":                       result = await getUser(client, args as any); break;
        case "list_groups":                    result = await listGroups(client, args as any); break;
        case "list_group_members":             result = await listGroupMembers(client, args as any); break;
        case "add_group_member":               result = await addGroupMember(client, args as any); break;
        case "remove_group_member":            result = await removeGroupMember(client, args as any); break;
        // Presence
        case "get_my_presence":                result = await getMyPresence(client); break;
        case "get_user_presence":              result = await getUserPresence(client, args as any); break;
        case "get_presence_for_users":         result = await getPresenceForUsers(client, args as any); break;
        case "set_my_presence":                result = await setMyPresence(client, args as any); break;
        // Bookings
        case "list_booking_businesses":        result = await listBookingBusinesses(client); break;
        case "list_booking_services":          result = await listBookingServices(client, args as any); break;
        case "list_booking_appointments":      result = await listBookingAppointments(client, args as any); break;
        case "create_booking_appointment":     result = await createBookingAppointment(client, args as any); break;
        case "cancel_booking_appointment":     result = await cancelBookingAppointment(client, args as any); break;
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

// ── Express Setup ──────────────────────────────────────────────────────────────

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use((req, _res, next) => {
  console.log(`${new Date().toISOString()} ${req.method} ${req.path}`);
  next();
});

app.get("/health", (_req, res) => {
  res.json({ status: "ok", service: "mcp-outlook" });
});

app.get("/icon.svg", (_req, res) => {
  const icon = readFileSync(join(__dirname, "icon.svg"));
  res.setHeader("Content-Type", "image/svg+xml");
  res.send(icon);
});

// ── OAuth 2.0 Dynamic Client Registration (DCR) ────────────────────────────────

const tenantId = process.env.ENTRA_TENANT_ID!;
const clientId = process.env.ENTRA_CLIENT_ID!;
const copilotClientId = process.env.COPILOT_CLIENT_ID!;
const copilotClientSecret = process.env.COPILOT_CLIENT_SECRET!;
const appUrl = process.env.APP_URL!;
const scope = `api://${clientId}/access`;

const authServerMetadata = {
  issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
  authorization_endpoint: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
  token_endpoint: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
  registration_endpoint: `${appUrl}/register`,
  scopes_supported: [scope, "openid", "profile", "offline_access"],
  response_types_supported: ["code"],
  grant_types_supported: ["authorization_code", "refresh_token"],
  code_challenge_methods_supported: ["S256"],
  token_endpoint_auth_methods_supported: ["none"],
};

const protectedResourceMetadata = {
  resource: appUrl,
  authorization_servers: [`https://login.microsoftonline.com/${tenantId}/v2.0`],
  scopes_supported: [scope],
  bearer_methods_supported: ["header"],
};

// Discovery endpoints – must be reachable without Bearer token
app.get(["/.well-known/oauth-authorization-server", "/.well-known/oauth-authorization-server/*"],
  (req, res) => {
    console.log(`[Discovery] oauth-authorization-server hit: ${req.path}`);
    res.json(authServerMetadata);
  });

app.get(["/.well-known/oauth-protected-resource", "/.well-known/oauth-protected-resource/*"],
  (req, res) => {
    console.log(`[Discovery] oauth-protected-resource hit: ${req.path}`);
    res.json(protectedResourceMetadata);
  });

app.post("/register", (req, res) => {
  console.log(`[Discovery] /register hit from ${req.ip}`);
  res.status(201).json({
    client_id: copilotClientId,
    client_secret: copilotClientSecret,
    client_name: "mcp-outlook-client",
    grant_types: ["authorization_code", "refresh_token"],
    response_types: ["code"],
    scope: `${scope} offline_access`,
    token_endpoint_auth_method: "client_secret_post",
  });
});

// ── All MCP endpoints require Entra ID Auth ────────────────────────────────────
app.use(entraAuthMiddleware);

// StreamableHTTP Transport (neueres Protokoll)
app.all("/mcp", async (req, res) => {
  const graphToken = await getGraphTokenViaObo(req.accessToken!);
  const client = getGraphClient(graphToken);
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  const server = createMcpServer(client);
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

// SSE Transport (älteres Protokoll)
interface SseSession {
  transport: SSEServerTransport;
  client: Client;
}
const sseSessions = new Map<string, SseSession>();

app.get("/sse", async (req, res) => {
  console.log("GET /sse - neue Verbindung");
  const graphToken = await getGraphTokenViaObo(req.accessToken!);
  const client = getGraphClient(graphToken);
  const transport = new SSEServerTransport("/messages", res);
  sseSessions.set(transport.sessionId, { transport, client });
  res.on("close", () => sseSessions.delete(transport.sessionId));
  const server = createMcpServer(client);
  await server.connect(transport);
});

app.post("/messages", async (req, res) => {
  const sessionId = req.query.sessionId as string;
  const session = sseSessions.get(sessionId);
  if (!session) {
    res.status(404).json({ error: "Session nicht gefunden" });
    return;
  }
  await session.transport.handlePostMessage(req, res);
});

const PORT = parseInt(process.env.PORT ?? "3000");
app.listen(PORT, () => {
  console.log(`MCP Outlook Server v2 läuft auf Port ${PORT}`);
});
