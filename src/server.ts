import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ToolAnnotations } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { DriveClient, DriveAPIError } from "./client.js";

const SAFE: ToolAnnotations = {
  readOnlyHint: true,
  destructiveHint: false,
  openWorldHint: true,
};

function errorResult(err: unknown): {
  content: Array<{ type: "text"; text: string }>;
  isError: true;
} {
  const message =
    err instanceof DriveAPIError
      ? err.message
      : err instanceof Error
        ? err.message
        : String(err);
  return {
    content: [{ type: "text" as const, text: `Error: ${message}` }],
    isError: true,
  };
}

function jsonResult(data: unknown): { content: Array<{ type: "text"; text: string }> } {
  return {
    content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
  };
}

export function createServer(client: DriveClient): McpServer {
  const server = new McpServer({
    name: "gdrive-mcp",
    version: "0.1.0",
  });

  server.tool(
    "gdrive_search",
    "Search for files in Google Drive using full-text search or Drive query syntax. " +
      "Examples: 'quarterly report', \"name contains 'budget'\", \"mimeType='application/pdf'\".",
    {
      query: z.string().describe("Search query (Drive full-text or query syntax)"),
      max_results: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(20)
        .describe("Maximum files to return (1-100)"),
      page_token: z
        .string()
        .optional()
        .describe("Pagination token from a previous gdrive_search result"),
    },
    SAFE,
    async ({ query, max_results, page_token }) => {
      try {
        const result = await client.search(query, max_results, page_token);
        return jsonResult(result);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_get_file",
    "Get detailed metadata for a single file by its Google Drive file ID. " +
      "Returns name, mimeType, size, owners, dates, webViewLink, and more.",
    {
      file_id: z.string().describe("Google Drive file ID"),
    },
    SAFE,
    async ({ file_id }) => {
      try {
        const result = await client.getFile(file_id);
        return jsonResult(result);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_read_file",
    "Read the content of a file from Google Drive. " +
      "Google Docs are exported as Markdown, Sheets as CSV (first sheet), " +
      "Slides as plain text. Other text files are read directly. " +
      "Binary files return an error with a link to open in browser.",
    {
      file_id: z.string().describe("Google Drive file ID"),
      max_chars: z
        .number()
        .int()
        .min(0)
        .default(100_000)
        .describe("Truncate content to this many characters (default 100000)"),
    },
    SAFE,
    async ({ file_id, max_chars }) => {
      try {
        const result = await client.readFile(file_id, max_chars);
        return {
          content: [{ type: "text" as const, text: result.content }],
        };
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_list_files",
    "List files in a Google Drive folder. Defaults to the root folder. " +
      "Returns file names, types, sizes, and modification times with pagination support.",
    {
      folder_id: z
        .string()
        .default("root")
        .describe("Folder ID to list (default: root)"),
      max_results: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(50)
        .describe("Maximum files to return (1-100)"),
      page_token: z
        .string()
        .optional()
        .describe("Pagination token from a previous gdrive_list_files result"),
      order_by: z
        .string()
        .default("modifiedTime desc")
        .describe(
          "Sort order (default: most recently modified first). " +
            "Supported keys: createdTime, modifiedTime, name, quotaBytesUsed, etc.",
        ),
    },
    SAFE,
    async ({ folder_id, max_results, page_token, order_by }) => {
      try {
        const result = await client.listFiles(
          folder_id,
          max_results,
          page_token,
          order_by,
        );
        return jsonResult(result);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  return server;
}
