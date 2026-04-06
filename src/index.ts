import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { google, sheets_v4 } from "googleapis";
import { z } from "zod";
import * as fs from "fs";
import * as path from "path";
import express from "express";
import cors from "cors";

// ── Auth ─────────────────────────────────────────────────────────────────────

function getAuth() {
  // Priority 1: JSON content as env var (Railway/production)
  const keyJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (keyJson) {
    const credentials = JSON.parse(keyJson);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    return auth;
  }

  // Priority 2: Path to key file (local Claude Code)
  const keyPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!keyPath) {
    throw new Error(
      "Set GOOGLE_SERVICE_ACCOUNT_JSON (JSON content) or GOOGLE_SERVICE_ACCOUNT_KEY (file path)."
    );
  }
  const keyFile = path.resolve(keyPath);
  if (!fs.existsSync(keyFile)) {
    throw new Error(`Service account key file not found: ${keyFile}`);
  }
  const auth = new google.auth.GoogleAuth({
    keyFile,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return auth;
}

function getSheets(): sheets_v4.Sheets {
  const auth = getAuth();
  return google.sheets({ version: "v4", auth });
}

// ── Tool schemas ──────────────────────────────────────────────────────────────

const ReadSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID (from the URL)"),
  range: z.string().describe("A1 notation range, e.g. 'Sheet1!A1:D20' or 'A1:Z100'"),
});

const WriteSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
  range: z.string().describe("A1 notation range to write to, e.g. 'Sheet1!A1'"),
  values: z
    .array(z.array(z.any()))
    .describe("2D array of values. Each inner array is a row."),
  value_input_option: z
    .enum(["RAW", "USER_ENTERED"])
    .default("USER_ENTERED")
    .describe("RAW = literal values; USER_ENTERED = same as typing in Sheets (parses formulas, dates)"),
});

const AppendSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
  range: z.string().describe("Range to search for a table to append to, e.g. 'Sheet1!A:Z'"),
  values: z
    .array(z.array(z.any()))
    .describe("2D array of rows to append"),
  value_input_option: z
    .enum(["RAW", "USER_ENTERED"])
    .default("USER_ENTERED"),
});

const ClearSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
  range: z.string().describe("A1 notation range to clear, e.g. 'Sheet1!A1:Z100'"),
});

const ListSheetsSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
});

const GetInfoSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
});

const BatchReadSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
  ranges: z
    .array(z.string())
    .describe("List of A1 notation ranges to read in one call"),
});

const CreateSheetSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
  sheet_name: z.string().describe("Name for the new sheet tab"),
});

const RenameSheetSchema = z.object({
  spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
  sheet_id: z.number().describe("Numeric sheet ID (from list_sheets)"),
  new_name: z.string().describe("New name for the sheet tab"),
});

// ── Tool implementations ──────────────────────────────────────────────────────

async function sheetsRead(args: z.infer<typeof ReadSchema>) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: args.spreadsheet_id,
    range: args.range,
    valueRenderOption: "FORMATTED_VALUE",
    dateTimeRenderOption: "FORMATTED_STRING",
  });
  const values = res.data.values ?? [];
  if (values.length === 0) return "No data found in the specified range.";

  // Format as markdown table if there are headers
  const [header, ...rows] = values;
  if (!header) return "Empty range.";

  const colWidths = header.map((h: unknown, i: number) =>
    Math.max(
      String(h ?? "").length,
      ...rows.map((r) => String(r[i] ?? "").length)
    )
  );

  const fmt = (row: unknown[]) =>
    "| " +
    header.map((_: unknown, i: number) =>
      String(row[i] ?? "").padEnd(colWidths[i])
    ).join(" | ") +
    " |";

  const separator =
    "| " + colWidths.map((w) => "-".repeat(w)).join(" | ") + " |";

  const table = [fmt(header), separator, ...rows.map(fmt)].join("\n");
  return `**Range:** \`${args.range}\`\n**Rows:** ${rows.length}\n\n${table}`;
}

async function sheetsWrite(args: z.infer<typeof WriteSchema>) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId: args.spreadsheet_id,
    range: args.range,
    valueInputOption: args.value_input_option,
    requestBody: { values: args.values },
  });
  return (
    `Updated ${res.data.updatedCells} cell(s) in range ${res.data.updatedRange}.`
  );
}

async function sheetsAppend(args: z.infer<typeof AppendSchema>) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId: args.spreadsheet_id,
    range: args.range,
    valueInputOption: args.value_input_option,
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: args.values },
  });
  const updates = res.data.updates;
  return `Appended ${args.values.length} row(s). Updated range: ${updates?.updatedRange ?? "unknown"}.`;
}

async function sheetsClear(args: z.infer<typeof ClearSchema>) {
  const sheets = getSheets();
  await sheets.spreadsheets.values.clear({
    spreadsheetId: args.spreadsheet_id,
    range: args.range,
  });
  return `Cleared range ${args.range} successfully.`;
}

async function sheetsListSheets(args: z.infer<typeof ListSheetsSchema>) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.get({
    spreadsheetId: args.spreadsheet_id,
    fields: "sheets.properties",
  });
  const sheetList = (res.data.sheets ?? []).map((s) => ({
    id: s.properties?.sheetId,
    title: s.properties?.title,
    index: s.properties?.index,
    rowCount: s.properties?.gridProperties?.rowCount,
    columnCount: s.properties?.gridProperties?.columnCount,
  }));
  return JSON.stringify(sheetList, null, 2);
}

async function sheetsGetInfo(args: z.infer<typeof GetInfoSchema>) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.get({
    spreadsheetId: args.spreadsheet_id,
    fields: "spreadsheetId,properties,sheets.properties",
  });
  return JSON.stringify(
    {
      id: res.data.spreadsheetId,
      title: res.data.properties?.title,
      locale: res.data.properties?.locale,
      timeZone: res.data.properties?.timeZone,
      sheets: (res.data.sheets ?? []).map((s) => ({
        id: s.properties?.sheetId,
        title: s.properties?.title,
        index: s.properties?.index,
      })),
    },
    null,
    2
  );
}

async function sheetsBatchRead(args: z.infer<typeof BatchReadSchema>) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: args.spreadsheet_id,
    ranges: args.ranges,
    valueRenderOption: "FORMATTED_VALUE",
  });
  const result: Record<string, unknown[][]> = {};
  for (const vr of res.data.valueRanges ?? []) {
    result[vr.range ?? "unknown"] = vr.values ?? [];
  }
  return JSON.stringify(result, null, 2);
}

async function sheetsCreateSheet(args: z.infer<typeof CreateSheetSchema>) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: args.spreadsheet_id,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: { title: args.sheet_name },
          },
        },
      ],
    },
  });
  const newSheet = res.data.replies?.[0]?.addSheet?.properties;
  return `Created sheet "${newSheet?.title}" with ID ${newSheet?.sheetId}.`;
}

async function sheetsRenameSheet(args: z.infer<typeof RenameSheetSchema>) {
  const sheets = getSheets();
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: args.spreadsheet_id,
    requestBody: {
      requests: [
        {
          updateSheetProperties: {
            properties: {
              sheetId: args.sheet_id,
              title: args.new_name,
            },
            fields: "title",
          },
        },
      ],
    },
  });
  return `Sheet renamed to "${args.new_name}".`;
}

// ── MCP Server ────────────────────────────────────────────────────────────────

const server = new Server(
  { name: "mcp-google-sheets", version: "1.0.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: "sheets_read",
      description:
        "Read data from a Google Sheets range. Returns a formatted table. " +
        "Use A1 notation like 'Sheet1!A1:D50'. If you omit the sheet name, uses the first sheet.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string", description: "Spreadsheet ID from the URL" },
          range: { type: "string", description: "A1 notation range, e.g. 'Sheet1!A1:Z100'" },
        },
        required: ["spreadsheet_id", "range"],
      },
    },
    {
      name: "sheets_write",
      description:
        "Write or overwrite values in a Google Sheets range. " +
        "Provide a 2D array: each inner array is a row, each item is a cell.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
          range: { type: "string", description: "Top-left cell of where to write, e.g. 'Sheet1!A1'" },
          values: {
            type: "array",
            items: { type: "array", items: {} },
            description: "2D array, e.g. [[\"Name\",\"Score\"],[\"Alice\",95]]",
          },
          value_input_option: {
            type: "string",
            enum: ["RAW", "USER_ENTERED"],
            default: "USER_ENTERED",
          },
        },
        required: ["spreadsheet_id", "range", "values"],
      },
    },
    {
      name: "sheets_append",
      description:
        "Append new rows to the end of a table in Google Sheets. " +
        "Finds the last row with data and adds after it.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
          range: { type: "string", description: "Range of the table, e.g. 'Sheet1!A:Z'" },
          values: {
            type: "array",
            items: { type: "array", items: {} },
          },
          value_input_option: {
            type: "string",
            enum: ["RAW", "USER_ENTERED"],
            default: "USER_ENTERED",
          },
        },
        required: ["spreadsheet_id", "range", "values"],
      },
    },
    {
      name: "sheets_clear",
      description: "Clear all values in a Google Sheets range (keeps formatting).",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
          range: { type: "string", description: "Range to clear, e.g. 'Sheet1!A2:Z1000'" },
        },
        required: ["spreadsheet_id", "range"],
      },
    },
    {
      name: "sheets_list",
      description: "List all sheet tabs in a spreadsheet with their IDs and dimensions.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
        },
        required: ["spreadsheet_id"],
      },
    },
    {
      name: "sheets_info",
      description: "Get metadata about a spreadsheet: title, locale, timezone, and list of sheets.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
        },
        required: ["spreadsheet_id"],
      },
    },
    {
      name: "sheets_batch_read",
      description: "Read multiple ranges from a spreadsheet in a single API call.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
          ranges: {
            type: "array",
            items: { type: "string" },
            description: "List of A1 notation ranges",
          },
        },
        required: ["spreadsheet_id", "ranges"],
      },
    },
    {
      name: "sheets_create_sheet",
      description: "Create a new sheet tab inside an existing spreadsheet.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
          sheet_name: { type: "string", description: "Name for the new tab" },
        },
        required: ["spreadsheet_id", "sheet_name"],
      },
    },
    {
      name: "sheets_rename_sheet",
      description: "Rename an existing sheet tab. Use sheets_list first to get the sheet_id.",
      inputSchema: {
        type: "object",
        properties: {
          spreadsheet_id: { type: "string" },
          sheet_id: { type: "number", description: "Numeric sheet ID from sheets_list" },
          new_name: { type: "string" },
        },
        required: ["spreadsheet_id", "sheet_id", "new_name"],
      },
    },
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    let result: string;

    switch (name) {
      case "sheets_read":
        result = await sheetsRead(ReadSchema.parse(args));
        break;
      case "sheets_write":
        result = await sheetsWrite(WriteSchema.parse(args));
        break;
      case "sheets_append":
        result = await sheetsAppend(AppendSchema.parse(args));
        break;
      case "sheets_clear":
        result = await sheetsClear(ClearSchema.parse(args));
        break;
      case "sheets_list":
        result = await sheetsListSheets(ListSheetsSchema.parse(args));
        break;
      case "sheets_info":
        result = await sheetsGetInfo(GetInfoSchema.parse(args));
        break;
      case "sheets_batch_read":
        result = await sheetsBatchRead(BatchReadSchema.parse(args));
        break;
      case "sheets_create_sheet":
        result = await sheetsCreateSheet(CreateSheetSchema.parse(args));
        break;
      case "sheets_rename_sheet":
        result = await sheetsRenameSheet(RenameSheetSchema.parse(args));
        break;
      default:
        throw new Error(`Unknown tool: ${name}`);
    }

    return { content: [{ type: "text", text: result }] };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return {
      content: [{ type: "text", text: `Error: ${message}` }],
      isError: true,
    };
  }
});

// ── Transport: stdio (Claude Code) or HTTP/SSE (Claude.ai web) ───────────────

async function startStdio() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Google Sheets MCP Server running on stdio");
}

async function startHttp() {
  const app = express();
  app.use(cors());
  app.use(express.json());

  const PORT = parseInt(process.env.PORT ?? "3000", 10);

  app.get("/health", (_req, res) => {
    res.json({ status: "ok", server: "mcp-google-sheets" });
  });

  // SSE transport — compatible with Claude.ai
  const sseTransports: Record<string, SSEServerTransport> = {};

  app.get("/sse", async (_req, res) => {
    const transport = new SSEServerTransport("/messages", res);
    sseTransports[transport.sessionId] = transport;
    res.on("close", () => delete sseTransports[transport.sessionId]);
    await server.connect(transport);
  });

  app.post("/messages", async (req, res) => {
    const sessionId = req.query.sessionId as string;
    const transport = sseTransports[sessionId];
    if (!transport) { res.status(404).json({ error: "Session not found" }); return; }
    await transport.handlePostMessage(req, res);
  });

  app.listen(PORT, "0.0.0.0", () => {
    console.error(`Google Sheets MCP Server running on HTTP port ${PORT}`);
    console.error(`SSE endpoint: http://localhost:${PORT}/sse`);
  });
}

async function main() {
  const mode = process.env.MCP_TRANSPORT ?? "stdio";
  if (mode === "http") {
    await startHttp();
  } else {
    await startStdio();
  }
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
