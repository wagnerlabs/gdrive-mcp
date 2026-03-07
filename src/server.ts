import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ToolAnnotations } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { DriveClient, DriveAPIError } from "./client.js";
import { SheetsClient, FormatOptions } from "./sheets-client.js";

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

function spreadsheetUrl(id: string): string {
  return `https://docs.google.com/spreadsheets/d/${id}/edit`;
}

function padValues(
  fetched: unknown[][],
  numRows: number,
  numCols: number,
): string[][] {
  const result: string[][] = [];
  for (let r = 0; r < numRows; r++) {
    const row: string[] = [];
    const fetchedRow = fetched[r] ?? [];
    for (let c = 0; c < numCols; c++) {
      const val = (fetchedRow as unknown[])[c];
      row.push(val !== undefined && val !== null ? String(val) : "");
    }
    result.push(row);
  }
  return result;
}

function describeShape(matrix: unknown[][]): string {
  const rows = matrix.length;
  const cols = rows > 0 ? Math.max(0, ...matrix.map((r) => (r as unknown[]).length)) : 0;
  return `${rows}x${cols}`;
}

/**
 * Returns `null` when actual values match expected, or a descriptive error
 * string when they don't. Also rejects mismatched shapes up-front.
 */
function checkPrecondition(
  actual: string[][],
  expected: (string | number | boolean)[][],
  values: (string | number | boolean)[][],
): string | null {
  if (expected.length !== values.length) {
    return (
      `expected_current_values must have the same shape as values ` +
      `(got ${describeShape(expected)}, expected ${describeShape(values)}).`
    );
  }
  for (let r = 0; r < values.length; r++) {
    if (expected[r].length !== values[r].length) {
      return (
        `expected_current_values must have the same shape as values ` +
        `(row ${r} has ${expected[r].length} columns, expected ${values[r].length}).`
      );
    }
  }
  for (let r = 0; r < actual.length; r++) {
    for (let c = 0; c < actual[r].length; c++) {
      if (actual[r][c] !== String(expected[r][c])) {
        return (
          "Precondition failed: current values do not match expected values.\n" +
          `Actual values: ${JSON.stringify(actual)}`
        );
      }
    }
  }
  return null;
}

const CellValue = z.union([z.string(), z.number(), z.boolean()]);
const ValuesArray = z.array(z.array(CellValue));

export function createServer(
  driveClient: DriveClient,
  sheetsClient: SheetsClient,
): McpServer {
  const server = new McpServer({
    name: "gdrive-mcp",
    version: "0.1.0",
  });

  const accessedSpreadsheets = new Set<string>();

  // ── Read-only Drive tools ──────────────────────────────────────────

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
        const result = await driveClient.search(query, max_results, page_token);
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
        const result = await driveClient.getFile(file_id);
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
        const meta = await driveClient.getFile(file_id);
        const result = await driveClient.readFile(file_id, max_chars);

        if (meta.mimeType === "application/vnd.google-apps.spreadsheet") {
          accessedSpreadsheets.add(file_id);
        } else if (
          meta.mimeType === "application/vnd.google-apps.shortcut" &&
          meta.shortcutTarget?.mimeType ===
            "application/vnd.google-apps.spreadsheet"
        ) {
          accessedSpreadsheets.add(meta.shortcutTarget.id);
        }

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
        const result = await driveClient.listFiles(
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

  // ── Read-only Sheets tool ──────────────────────────────────────────

  server.tool(
    "gdrive_get_spreadsheet_info",
    "Get spreadsheet metadata including all sheet tabs, their dimensions, and named ranges. " +
      "Use this to discover tab names and structure before reading or writing.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
    },
    SAFE,
    async ({ spreadsheet_id }) => {
      try {
        const info = await sheetsClient.getSpreadsheet(spreadsheet_id);
        accessedSpreadsheets.add(spreadsheet_id);
        return jsonResult(info);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  // ── Write tools — data operations ──────────────────────────────────

  server.tool(
    "gdrive_create_sheet",
    "Create a new Google Sheets spreadsheet in the user's root Drive folder.",
    {
      title: z.string().describe("Spreadsheet title"),
      sheet_names: z
        .array(z.string())
        .optional()
        .describe("Names for individual sheet tabs (default: ['Sheet1'])"),
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ title, sheet_names }) => {
      try {
        const result = await sheetsClient.createSpreadsheet(title, sheet_names);
        accessedSpreadsheets.add(result.spreadsheetId);
        return jsonResult(result);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_update_sheet",
    "Overwrite values in a specific cell range of a Google Sheets spreadsheet. " +
      "You must read the spreadsheet first (using gdrive_read_file or gdrive_get_spreadsheet_info). " +
      "value_input_option controls parsing: USER_ENTERED (default) parses formulas and formats " +
      "numbers/dates automatically; RAW stores values exactly as provided. " +
      "Provide expected_current_values (same shape as values) for small targeted edits as a safety " +
      "check — the write is refused if current values don't match. Skip it for bulk operations to " +
      "avoid doubling API calls. Set include_previous_values to true to see what was overwritten. " +
      "Use empty strings for blank cells, or gdrive_clear_values to clear a range.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      range: z.string().describe("A1 notation, e.g. 'Sheet1!A1:C3'"),
      values: ValuesArray.describe(
        "2D array of cell values, e.g. [[\"Name\",\"Age\"],[\"Alice\",30]]. " +
          "Each inner array is one row; cells can be strings, numbers, or booleans.",
      ),
      value_input_option: z
        .enum(["USER_ENTERED", "RAW"])
        .default("USER_ENTERED")
        .describe("USER_ENTERED parses formulas/dates; RAW stores literally"),
      expected_current_values: ValuesArray.optional().describe(
        "Expected current cell values (same shape as values), e.g. [[\"old\"]]. " +
          "If provided, the write is refused when actual values differ. Use for small targeted edits.",
      ),
      include_previous_values: z
        .boolean()
        .default(false)
        .describe("Include previous cell values in the response"),
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      spreadsheet_id,
      range,
      values,
      value_input_option,
      expected_current_values,
      include_previous_values,
    }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        let previousValues: string[][] | undefined;

        if (expected_current_values || include_previous_values) {
          const numRows = values.length;
          const numCols = Math.max(0, ...values.map((r) => r.length));
          const fetched = await sheetsClient.getValues(spreadsheet_id, range);
          previousValues = padValues(fetched, numRows, numCols);

          if (expected_current_values) {
            const mismatch = checkPrecondition(
              previousValues,
              expected_current_values,
              values,
            );
            if (mismatch) {
              return errorResult(new Error(mismatch));
            }
          }
        }

        const result = await sheetsClient.updateValues(
          spreadsheet_id,
          range,
          values,
          value_input_option,
        );

        const response: Record<string, unknown> = {
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          updatedRange: result.updatedRange,
          updatedRows: result.updatedRows,
          updatedColumns: result.updatedColumns,
          updatedCells: result.updatedCells,
        };
        if (previousValues) {
          response.previousValues = previousValues;
        }
        return jsonResult(response);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_append_sheet",
    "Append rows after existing data in a Google Sheets spreadsheet. " +
      "You must read the spreadsheet first. Rows are always inserted (never overwrite existing data). " +
      "value_input_option controls parsing: USER_ENTERED (default) parses formulas and formats; " +
      "RAW stores values literally.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      range: z
        .string()
        .describe("A1 notation of the table to append to, e.g. 'Sheet1!A:C'"),
      values: ValuesArray.describe("2D array of rows to append"),
      value_input_option: z
        .enum(["USER_ENTERED", "RAW"])
        .default("USER_ENTERED")
        .describe("USER_ENTERED parses formulas/dates; RAW stores literally"),
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ spreadsheet_id, range, values, value_input_option }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        const result = await sheetsClient.appendValues(
          spreadsheet_id,
          range,
          values,
          value_input_option,
        );
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          tableRange: result.tableRange,
          updatedRange: result.updatedRange,
          updatedRows: result.updatedRows,
          updatedCells: result.updatedCells,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_clear_values",
    "Clear all values from a cell range in a Google Sheets spreadsheet. " +
      "Formatting is preserved; only values are removed. You must read the spreadsheet first.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      range: z.string().describe("A1 notation, e.g. 'Sheet1!A1:C10'"),
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ spreadsheet_id, range }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        const result = await sheetsClient.clearValues(spreadsheet_id, range);
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          clearedRange: result.clearedRange,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  // ── Write tools — formatting ───────────────────────────────────────

  server.tool(
    "gdrive_format_cells",
    "Apply formatting to a cell range in a Google Sheets spreadsheet. " +
      "Provide at least one formatting parameter. Range must be bounded A1 notation with " +
      "explicit sheet name (e.g. 'Sheet1!A1:C5'). You must read the spreadsheet first.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      range: z
        .string()
        .describe(
          "Bounded A1 notation with sheet name, e.g. \"Sheet1!A1:C5\" or \"'Q1 Budget'!A1:C5\"",
        ),
      bold: z.boolean().optional().describe("Bold text"),
      italic: z.boolean().optional().describe("Italic text"),
      underline: z.boolean().optional().describe("Underline text"),
      strikethrough: z.boolean().optional().describe("Strikethrough text"),
      font_size: z.number().optional().describe("Font size in points"),
      font_family: z.string().optional().describe("Font family, e.g. 'Arial'"),
      text_color: z
        .string()
        .optional()
        .describe("Text color as hex, e.g. '#FF0000'"),
      background_color: z
        .string()
        .optional()
        .describe("Background color as hex, e.g. '#FFFF00'"),
      horizontal_alignment: z
        .enum(["LEFT", "CENTER", "RIGHT"])
        .optional()
        .describe("Horizontal text alignment"),
      vertical_alignment: z
        .enum(["TOP", "MIDDLE", "BOTTOM"])
        .optional()
        .describe("Vertical text alignment"),
      wrap_strategy: z
        .enum(["OVERFLOW_CELL", "CLIP", "WRAP"])
        .optional()
        .describe("Text wrapping strategy"),
      number_format_type: z
        .enum([
          "TEXT",
          "NUMBER",
          "PERCENT",
          "CURRENCY",
          "DATE",
          "TIME",
          "DATE_TIME",
          "SCIENTIFIC",
        ])
        .optional()
        .describe("Number format type"),
      number_format_pattern: z
        .string()
        .optional()
        .describe("Number format pattern, e.g. '$#,##0.00' or 'yyyy-mm-dd'"),
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      spreadsheet_id,
      range,
      bold,
      italic,
      underline,
      strikethrough,
      font_size,
      font_family,
      text_color,
      background_color,
      horizontal_alignment,
      vertical_alignment,
      wrap_strategy,
      number_format_type,
      number_format_pattern,
    }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }

      const options: FormatOptions = {};
      let hasFormat = false;
      if (bold !== undefined) { options.bold = bold; hasFormat = true; }
      if (italic !== undefined) { options.italic = italic; hasFormat = true; }
      if (underline !== undefined) { options.underline = underline; hasFormat = true; }
      if (strikethrough !== undefined) { options.strikethrough = strikethrough; hasFormat = true; }
      if (font_size !== undefined) { options.fontSize = font_size; hasFormat = true; }
      if (font_family !== undefined) { options.fontFamily = font_family; hasFormat = true; }
      if (text_color !== undefined) { options.textColor = text_color; hasFormat = true; }
      if (background_color !== undefined) { options.backgroundColor = background_color; hasFormat = true; }
      if (horizontal_alignment !== undefined) { options.horizontalAlignment = horizontal_alignment; hasFormat = true; }
      if (vertical_alignment !== undefined) { options.verticalAlignment = vertical_alignment; hasFormat = true; }
      if (wrap_strategy !== undefined) { options.wrapStrategy = wrap_strategy; hasFormat = true; }
      if (number_format_type !== undefined) { options.numberFormatType = number_format_type; hasFormat = true; }
      if (number_format_pattern !== undefined) { options.numberFormatPattern = number_format_pattern; hasFormat = true; }

      if (!hasFormat) {
        return errorResult(
          new Error("At least one formatting parameter must be provided."),
        );
      }

      try {
        await sheetsClient.formatCells(spreadsheet_id, range, options);
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          formattedRange: range,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  // ── Write tools — tab management ───────────────────────────────────

  server.tool(
    "gdrive_add_sheet_tab",
    "Add a new sheet tab to an existing Google Sheets spreadsheet. " +
      "You must read the spreadsheet first.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      title: z.string().describe("Name for the new sheet tab"),
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ spreadsheet_id, title }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        const result = await sheetsClient.addSheet(spreadsheet_id, title);
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          sheetId: result.sheetId,
          title: result.title,
          index: result.index,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_delete_sheet_tab",
    "Delete an entire sheet tab and all its data from a Google Sheets spreadsheet. " +
      "This cannot be undone. You must read the spreadsheet first.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      title: z.string().describe("Name of the sheet tab to delete"),
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ spreadsheet_id, title }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        await sheetsClient.deleteSheet(spreadsheet_id, title);
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          deletedTitle: title,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_rename_sheet_tab",
    "Rename an existing sheet tab in a Google Sheets spreadsheet. " +
      "You must read the spreadsheet first.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      current_title: z.string().describe("Current name of the sheet tab"),
      new_title: z.string().describe("New name for the sheet tab"),
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ spreadsheet_id, current_title, new_title }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        const result = await sheetsClient.renameSheet(
          spreadsheet_id,
          current_title,
          new_title,
        );
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          sheetId: result.sheetId,
          title: result.title,
          index: result.index,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  // ── Write tools — layout ───────────────────────────────────────────

  server.tool(
    "gdrive_insert_rows_columns",
    "Insert empty rows or columns into a sheet tab. " +
      "Inserts before the specified 0-based index. You must read the spreadsheet first.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      sheet_title: z.string().describe("Name of the sheet tab to modify"),
      dimension: z.enum(["ROWS", "COLUMNS"]).describe("Whether to insert rows or columns"),
      start_index: z
        .number()
        .int()
        .min(0)
        .describe("0-based index to insert before"),
      count: z.number().int().min(1).describe("Number of rows or columns to insert"),
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ spreadsheet_id, sheet_title, dimension, start_index, count }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        await sheetsClient.insertDimension(
          spreadsheet_id,
          sheet_title,
          dimension,
          start_index,
          count,
        );
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          sheetTitle: sheet_title,
          dimension,
          startIndex: start_index,
          count,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_delete_rows_columns",
    "Delete rows or columns and all data in them from a sheet tab. " +
      "Deletion range is [start_index, start_index + count) (0-based, end-exclusive). " +
      "You must read the spreadsheet first.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      sheet_title: z.string().describe("Name of the sheet tab to modify"),
      dimension: z.enum(["ROWS", "COLUMNS"]).describe("Whether to delete rows or columns"),
      start_index: z
        .number()
        .int()
        .min(0)
        .describe("0-based start index (inclusive)"),
      count: z.number().int().min(1).describe("Number of rows or columns to delete"),
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ spreadsheet_id, sheet_title, dimension, start_index, count }) => {
      if (!accessedSpreadsheets.has(spreadsheet_id)) {
        return errorResult(
          new Error(
            "You must read this spreadsheet before writing to it. " +
              "Use gdrive_read_file or gdrive_get_spreadsheet_info first.",
          ),
        );
      }
      try {
        await sheetsClient.deleteDimension(
          spreadsheet_id,
          sheet_title,
          dimension,
          start_index,
          count,
        );
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          sheetTitle: sheet_title,
          dimension,
          startIndex: start_index,
          count,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  return server;
}
