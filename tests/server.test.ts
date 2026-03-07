import { describe, it, expect, vi } from "vitest";
import { createServer } from "../src/server.js";
import { DriveClient } from "../src/client.js";
import { SheetsClient } from "../src/sheets-client.js";

function makeMockDriveClient(): DriveClient {
  return {
    search: vi.fn(),
    getFile: vi.fn(),
    readFile: vi.fn(),
    listFiles: vi.fn(),
  } as unknown as DriveClient;
}

function makeMockSheetsClient(): SheetsClient {
  return {
    getSpreadsheet: vi.fn(),
    createSpreadsheet: vi.fn(),
    getValues: vi.fn(),
    updateValues: vi.fn(),
    appendValues: vi.fn(),
    clearValues: vi.fn(),
    formatCells: vi.fn(),
    addSheet: vi.fn(),
    deleteSheet: vi.fn(),
    renameSheet: vi.fn(),
    insertDimension: vi.fn(),
    deleteDimension: vi.fn(),
    getSheetId: vi.fn(),
  } as unknown as SheetsClient;
}

function makeServer() {
  const drive = makeMockDriveClient();
  const sheets = makeMockSheetsClient();
  const server = createServer(drive, sheets);
  return { server, drive, sheets };
}

function getTools(server: ReturnType<typeof createServer>) {
  return (server as any)._registeredTools as Record<string, any>;
}

// ── Tool registration ────────────────────────────────────────────────

describe("createServer — tool registration", () => {
  it("registers all 15 tools", () => {
    const { server } = makeServer();
    const names = Object.keys(getTools(server));

    expect(names).toHaveLength(15);
    for (const name of [
      "gdrive_search",
      "gdrive_get_file",
      "gdrive_read_file",
      "gdrive_list_files",
      "gdrive_get_spreadsheet_info",
      "gdrive_create_sheet",
      "gdrive_update_sheet",
      "gdrive_append_sheet",
      "gdrive_clear_values",
      "gdrive_format_cells",
      "gdrive_add_sheet_tab",
      "gdrive_delete_sheet_tab",
      "gdrive_rename_sheet_tab",
      "gdrive_insert_rows_columns",
      "gdrive_delete_rows_columns",
    ]) {
      expect(names, `missing ${name}`).toContain(name);
    }
  });

  it("read-only tools have SAFE annotations", () => {
    const tools = getTools(makeServer().server);
    for (const name of [
      "gdrive_search",
      "gdrive_get_file",
      "gdrive_read_file",
      "gdrive_list_files",
      "gdrive_get_spreadsheet_info",
    ]) {
      const ann = tools[name].annotations;
      expect(ann?.readOnlyHint, `${name} readOnlyHint`).toBe(true);
      expect(ann?.destructiveHint, `${name} destructiveHint`).toBe(false);
    }
  });

  it("each write tool has correct per-tool annotations", () => {
    const tools = getTools(makeServer().server);
    const expected: Record<
      string,
      { readOnly: boolean; destructive: boolean; idempotent: boolean }
    > = {
      gdrive_create_sheet: { readOnly: false, destructive: false, idempotent: false },
      gdrive_update_sheet: { readOnly: false, destructive: true, idempotent: true },
      gdrive_append_sheet: { readOnly: false, destructive: false, idempotent: false },
      gdrive_clear_values: { readOnly: false, destructive: true, idempotent: true },
      gdrive_format_cells: { readOnly: false, destructive: false, idempotent: true },
      gdrive_add_sheet_tab: { readOnly: false, destructive: false, idempotent: false },
      gdrive_delete_sheet_tab: { readOnly: false, destructive: true, idempotent: false },
      gdrive_rename_sheet_tab: { readOnly: false, destructive: true, idempotent: false },
      gdrive_insert_rows_columns: { readOnly: false, destructive: false, idempotent: false },
      gdrive_delete_rows_columns: { readOnly: false, destructive: true, idempotent: false },
    };
    for (const [name, exp] of Object.entries(expected)) {
      const ann = tools[name].annotations;
      expect(ann?.readOnlyHint, `${name} readOnlyHint`).toBe(exp.readOnly);
      expect(ann?.destructiveHint, `${name} destructiveHint`).toBe(exp.destructive);
      expect(ann?.idempotentHint, `${name} idempotentHint`).toBe(exp.idempotent);
      expect(ann?.openWorldHint, `${name} openWorldHint`).toBe(true);
    }
  });
});

// ── Read-before-write guard ──────────────────────────────────────────

describe("read-before-write guard", () => {
  const writeTools = [
    "gdrive_update_sheet",
    "gdrive_append_sheet",
    "gdrive_clear_values",
    "gdrive_format_cells",
    "gdrive_add_sheet_tab",
    "gdrive_delete_sheet_tab",
    "gdrive_rename_sheet_tab",
    "gdrive_insert_rows_columns",
    "gdrive_delete_rows_columns",
  ];

  for (const toolName of writeTools) {
    it(`${toolName} rejects when spreadsheet has not been read`, async () => {
      const { server } = makeServer();
      const handler = getTools(server)[toolName].handler;

      const args: Record<string, unknown> = { spreadsheet_id: "unread-id" };
      if (toolName === "gdrive_update_sheet") {
        args.range = "Sheet1!A1";
        args.values = [["x"]];
        args.value_input_option = "USER_ENTERED";
        args.include_previous_values = false;
      } else if (toolName === "gdrive_append_sheet") {
        args.range = "Sheet1!A:A";
        args.values = [["x"]];
        args.value_input_option = "USER_ENTERED";
      } else if (toolName === "gdrive_clear_values") {
        args.range = "Sheet1!A1";
      } else if (toolName === "gdrive_format_cells") {
        args.range = "Sheet1!A1";
        args.bold = true;
      } else if (toolName === "gdrive_add_sheet_tab") {
        args.title = "New Tab";
      } else if (toolName === "gdrive_delete_sheet_tab") {
        args.title = "Sheet1";
      } else if (toolName === "gdrive_rename_sheet_tab") {
        args.current_title = "Sheet1";
        args.new_title = "Renamed";
      } else if (toolName === "gdrive_insert_rows_columns") {
        args.sheet_title = "Sheet1";
        args.dimension = "ROWS";
        args.start_index = 0;
        args.count = 1;
      } else if (toolName === "gdrive_delete_rows_columns") {
        args.sheet_title = "Sheet1";
        args.dimension = "ROWS";
        args.start_index = 0;
        args.count = 1;
      }

      const result = await handler(args, {});
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("must read this spreadsheet");
    });
  }

  it("allows write after gdrive_get_spreadsheet_info", async () => {
    const { server, sheets } = makeServer();
    const tools = getTools(server);

    (sheets.getSpreadsheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      title: "Test",
      spreadsheetUrl: "https://docs.google.com/spreadsheets/d/s1/edit",
      sheets: [{ sheetId: 0, title: "Sheet1", index: 0, rowCount: 100, columnCount: 26 }],
      namedRanges: [],
    });
    (sheets.clearValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      clearedRange: "Sheet1!A1",
    });

    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});
    const result = await tools["gdrive_clear_values"].handler(
      { spreadsheet_id: "s1", range: "Sheet1!A1" },
      {},
    );

    expect(result.isError).toBeUndefined();
  });

  it("allows write after gdrive_read_file reads a spreadsheet", async () => {
    const { server, drive, sheets } = makeServer();
    const tools = getTools(server);

    (drive.getFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      id: "s1",
      name: "My Sheet",
      mimeType: "application/vnd.google-apps.spreadsheet",
    });
    (drive.readFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      content: "[Note: CSV export contains the first sheet only.]\n\na,b\n1,2",
      mimeType: "text/csv",
      truncated: false,
    });
    (sheets.addSheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      sheetId: 1,
      title: "Tab2",
      index: 1,
    });

    await tools["gdrive_read_file"].handler({ file_id: "s1", max_chars: 100000 }, {});
    const result = await tools["gdrive_add_sheet_tab"].handler(
      { spreadsheet_id: "s1", title: "Tab2" },
      {},
    );

    expect(result.isError).toBeUndefined();
  });

  it("gdrive_create_sheet adds the new spreadsheet to the read set", async () => {
    const { server, sheets } = makeServer();
    const tools = getTools(server);

    (sheets.createSpreadsheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "new-id",
      spreadsheetUrl: "https://docs.google.com/spreadsheets/d/new-id/edit",
      sheets: [{ sheetId: 0, title: "Sheet1", index: 0 }],
    });
    (sheets.appendValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "new-id",
      updatedRange: "Sheet1!A1:A1",
      updatedRows: 1,
      updatedCells: 1,
    });

    await tools["gdrive_create_sheet"].handler({ title: "New Sheet" }, {});
    const result = await tools["gdrive_append_sheet"].handler(
      { spreadsheet_id: "new-id", range: "Sheet1!A:A", values: [["data"]], value_input_option: "USER_ENTERED" },
      {},
    );

    expect(result.isError).toBeUndefined();
  });

  it("failed gdrive_read_file does NOT unlock writes", async () => {
    const { server, drive } = makeServer();
    const tools = getTools(server);

    (drive.getFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      id: "s1",
      name: "Big Sheet",
      mimeType: "application/vnd.google-apps.spreadsheet",
    });
    (drive.readFile as ReturnType<typeof vi.fn>).mockRejectedValue(
      new Error("File too large for export"),
    );

    const readResult = await tools["gdrive_read_file"].handler(
      { file_id: "s1", max_chars: 100000 },
      {},
    );
    expect(readResult.isError).toBe(true);

    const writeResult = await tools["gdrive_clear_values"].handler(
      { spreadsheet_id: "s1", range: "Sheet1!A1" },
      {},
    );
    expect(writeResult.isError).toBe(true);
    expect(writeResult.content[0].text).toContain("must read this spreadsheet");
  });

  it("shortcut to spreadsheet unlocks writes for the target ID", async () => {
    const { server, drive, sheets } = makeServer();
    const tools = getTools(server);

    (drive.getFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      id: "shortcut1",
      name: "My Shortcut",
      mimeType: "application/vnd.google-apps.shortcut",
      shortcutTarget: {
        id: "real-spreadsheet-id",
        mimeType: "application/vnd.google-apps.spreadsheet",
      },
    });
    (drive.readFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      content: "[Note: CSV export contains the first sheet only.]\n\na,b\n1,2",
      mimeType: "text/csv",
      truncated: false,
    });
    (sheets.clearValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "real-spreadsheet-id",
      clearedRange: "Sheet1!A1",
    });

    await tools["gdrive_read_file"].handler(
      { file_id: "shortcut1", max_chars: 100000 },
      {},
    );

    const result = await tools["gdrive_clear_values"].handler(
      { spreadsheet_id: "real-spreadsheet-id", range: "Sheet1!A1" },
      {},
    );
    expect(result.isError).toBeUndefined();
  });

  it("shortcut to spreadsheet does NOT unlock writes for the shortcut ID itself", async () => {
    const { server, drive } = makeServer();
    const tools = getTools(server);

    (drive.getFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      id: "shortcut1",
      name: "My Shortcut",
      mimeType: "application/vnd.google-apps.shortcut",
      shortcutTarget: {
        id: "real-spreadsheet-id",
        mimeType: "application/vnd.google-apps.spreadsheet",
      },
    });
    (drive.readFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      content: "[Note: CSV export contains the first sheet only.]\n\na,b\n1,2",
      mimeType: "text/csv",
      truncated: false,
    });

    await tools["gdrive_read_file"].handler(
      { file_id: "shortcut1", max_chars: 100000 },
      {},
    );

    const result = await tools["gdrive_clear_values"].handler(
      { spreadsheet_id: "shortcut1", range: "Sheet1!A1" },
      {},
    );
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("must read this spreadsheet");
  });

  it("gdrive_read_file does NOT add non-spreadsheet files to the read set", async () => {
    const { server, drive } = makeServer();
    const tools = getTools(server);

    (drive.getFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      id: "doc1",
      name: "My Doc",
      mimeType: "application/vnd.google-apps.document",
    });
    (drive.readFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      content: "# Hello",
      mimeType: "text/markdown",
      truncated: false,
    });

    await tools["gdrive_read_file"].handler({ file_id: "doc1", max_chars: 100000 }, {});

    const result = await tools["gdrive_clear_values"].handler(
      { spreadsheet_id: "doc1", range: "Sheet1!A1" },
      {},
    );
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("must read this spreadsheet");
  });
});

// ── Precondition check (gdrive_update_sheet) ─────────────────────────

describe("gdrive_update_sheet — precondition check", () => {
  function setupForUpdate() {
    const { server, sheets } = makeServer();
    const tools = getTools(server);
    // Pre-populate the read set via getSpreadsheet mock
    (sheets.getSpreadsheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      title: "Test",
      spreadsheetUrl: "https://docs.google.com/spreadsheets/d/s1/edit",
      sheets: [],
      namedRanges: [],
    });
    return { tools, sheets };
  }

  it("succeeds when expected_current_values match", async () => {
    const { tools, sheets } = setupForUpdate();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.getValues as ReturnType<typeof vi.fn>).mockResolvedValue([["old"]]);
    (sheets.updateValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      updatedRange: "Sheet1!A1",
      updatedRows: 1,
      updatedColumns: 1,
      updatedCells: 1,
    });

    const result = await tools["gdrive_update_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A1",
        values: [["new"]],
        value_input_option: "USER_ENTERED",
        expected_current_values: [["old"]],
        include_previous_values: false,
      },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(sheets.updateValues).toHaveBeenCalled();
    const data = JSON.parse(result.content[0].text);
    expect(data.previousValues).toEqual([["old"]]);
  });

  it("rejects when expected_current_values do not match", async () => {
    const { tools, sheets } = setupForUpdate();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.getValues as ReturnType<typeof vi.fn>).mockResolvedValue([["actual"]]);

    const result = await tools["gdrive_update_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A1",
        values: [["new"]],
        value_input_option: "USER_ENTERED",
        expected_current_values: [["expected"]],
        include_previous_values: false,
      },
      {},
    );

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("Precondition failed");
    expect(result.content[0].text).toContain("actual");
    expect(sheets.updateValues).not.toHaveBeenCalled();
  });

  it("rejects when expected_current_values has fewer rows than values", async () => {
    const { tools, sheets } = setupForUpdate();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.getValues as ReturnType<typeof vi.fn>).mockResolvedValue([["a"], ["b"]]);

    const result = await tools["gdrive_update_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A1:A2",
        values: [["x"], ["y"]],
        value_input_option: "USER_ENTERED",
        expected_current_values: [["a"]],
        include_previous_values: false,
      },
      {},
    );

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("same shape as values");
    expect(result.content[0].text).toContain("1x1");
    expect(result.content[0].text).toContain("2x1");
    expect(sheets.updateValues).not.toHaveBeenCalled();
  });

  it("rejects when expected_current_values has fewer columns in a row", async () => {
    const { tools, sheets } = setupForUpdate();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.getValues as ReturnType<typeof vi.fn>).mockResolvedValue([["a", "b", "c"]]);

    const result = await tools["gdrive_update_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A1:C1",
        values: [["x", "y", "z"]],
        value_input_option: "USER_ENTERED",
        expected_current_values: [["a"]],
        include_previous_values: false,
      },
      {},
    );

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("same shape as values");
    expect(result.content[0].text).toContain("row 0");
    expect(sheets.updateValues).not.toHaveBeenCalled();
  });

  it("skips precondition read when neither flag is set", async () => {
    const { tools, sheets } = setupForUpdate();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.updateValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      updatedRange: "Sheet1!A1",
      updatedRows: 1,
      updatedColumns: 1,
      updatedCells: 1,
    });

    const result = await tools["gdrive_update_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A1",
        values: [["new"]],
        value_input_option: "USER_ENTERED",
        include_previous_values: false,
      },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(sheets.getValues).not.toHaveBeenCalled();
    const data = JSON.parse(result.content[0].text);
    expect(data.previousValues).toBeUndefined();
  });

  it("includes previousValues when include_previous_values is true", async () => {
    const { tools, sheets } = setupForUpdate();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.getValues as ReturnType<typeof vi.fn>).mockResolvedValue([["before"]]);
    (sheets.updateValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      updatedRange: "Sheet1!A1",
      updatedRows: 1,
      updatedColumns: 1,
      updatedCells: 1,
    });

    const result = await tools["gdrive_update_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A1",
        values: [["after"]],
        value_input_option: "USER_ENTERED",
        include_previous_values: true,
      },
      {},
    );

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.previousValues).toEqual([["before"]]);
  });

  it("pads fetched values with empty strings for trailing blanks", async () => {
    const { tools, sheets } = setupForUpdate();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    // Google trims trailing empties — returns only [["hello", "world"], ["foo"]]
    (sheets.getValues as ReturnType<typeof vi.fn>).mockResolvedValue([
      ["hello", "world"],
      ["foo"],
    ]);
    (sheets.updateValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      updatedRange: "Sheet1!A1:C3",
      updatedRows: 3,
      updatedColumns: 3,
      updatedCells: 9,
    });

    const result = await tools["gdrive_update_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A1:C3",
        values: [
          ["hello", "world", ""],
          ["foo", "", ""],
          ["", "", ""],
        ],
        value_input_option: "USER_ENTERED",
        expected_current_values: [
          ["hello", "world", ""],
          ["foo", "", ""],
          ["", "", ""],
        ],
        include_previous_values: false,
      },
      {},
    );

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.previousValues).toEqual([
      ["hello", "world", ""],
      ["foo", "", ""],
      ["", "", ""],
    ]);
  });
});

// ── gdrive_append_sheet — INSERT_ROWS ────────────────────────────────

describe("gdrive_append_sheet", () => {
  it("calls appendValues with the correct arguments", async () => {
    const { server, sheets } = makeServer();
    const tools = getTools(server);

    (sheets.getSpreadsheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      title: "T",
      spreadsheetUrl: "url",
      sheets: [],
      namedRanges: [],
    });
    (sheets.appendValues as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      updatedRange: "Sheet1!A4:B4",
      updatedRows: 1,
      updatedCells: 2,
    });

    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});
    const result = await tools["gdrive_append_sheet"].handler(
      {
        spreadsheet_id: "s1",
        range: "Sheet1!A:B",
        values: [["x", "y"]],
        value_input_option: "USER_ENTERED",
      },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(sheets.appendValues).toHaveBeenCalledWith("s1", "Sheet1!A:B", [["x", "y"]], "USER_ENTERED");
  });
});

// ── gdrive_format_cells — validation ─────────────────────────────────

describe("gdrive_format_cells", () => {
  it("rejects when no formatting parameters are provided", async () => {
    const { server, sheets } = makeServer();
    const tools = getTools(server);

    (sheets.getSpreadsheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      title: "T",
      spreadsheetUrl: "url",
      sheets: [],
      namedRanges: [],
    });

    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});
    const result = await tools["gdrive_format_cells"].handler(
      { spreadsheet_id: "s1", range: "Sheet1!A1:B2" },
      {},
    );

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("At least one formatting parameter");
  });

  it("succeeds with valid formatting parameters", async () => {
    const { server, sheets } = makeServer();
    const tools = getTools(server);

    (sheets.getSpreadsheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      title: "T",
      spreadsheetUrl: "url",
      sheets: [],
      namedRanges: [],
    });
    (sheets.formatCells as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});
    const result = await tools["gdrive_format_cells"].handler(
      { spreadsheet_id: "s1", range: "Sheet1!A1:B2", bold: true, font_size: 14 },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(sheets.formatCells).toHaveBeenCalled();
    const data = JSON.parse(result.content[0].text);
    expect(data.formattedRange).toBe("Sheet1!A1:B2");
  });
});

// ── Write tool responses ─────────────────────────────────────────────

describe("write tool responses", () => {
  function setupRead() {
    const { server, sheets } = makeServer();
    const tools = getTools(server);
    (sheets.getSpreadsheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      spreadsheetId: "s1",
      title: "T",
      spreadsheetUrl: "https://docs.google.com/spreadsheets/d/s1/edit",
      sheets: [{ sheetId: 0, title: "Sheet1", index: 0, rowCount: 100, columnCount: 26 }],
      namedRanges: [],
    });
    return { tools, sheets };
  }

  it("gdrive_delete_sheet_tab returns deletedTitle", async () => {
    const { tools, sheets } = setupRead();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.deleteSheet as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    const result = await tools["gdrive_delete_sheet_tab"].handler(
      { spreadsheet_id: "s1", title: "Sheet1" },
      {},
    );

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.deletedTitle).toBe("Sheet1");
    expect(data.spreadsheetId).toBe("s1");
  });

  it("gdrive_rename_sheet_tab returns updated title", async () => {
    const { tools, sheets } = setupRead();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.renameSheet as ReturnType<typeof vi.fn>).mockResolvedValue({
      sheetId: 0,
      title: "NewName",
      index: 0,
    });

    const result = await tools["gdrive_rename_sheet_tab"].handler(
      { spreadsheet_id: "s1", current_title: "Sheet1", new_title: "NewName" },
      {},
    );

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.title).toBe("NewName");
    expect(data.sheetId).toBe(0);
  });

  it("gdrive_insert_rows_columns returns operation details", async () => {
    const { tools, sheets } = setupRead();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.insertDimension as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    const result = await tools["gdrive_insert_rows_columns"].handler(
      { spreadsheet_id: "s1", sheet_title: "Sheet1", dimension: "ROWS", start_index: 2, count: 3 },
      {},
    );

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.dimension).toBe("ROWS");
    expect(data.startIndex).toBe(2);
    expect(data.count).toBe(3);
  });

  it("gdrive_delete_rows_columns returns operation details", async () => {
    const { tools, sheets } = setupRead();
    await tools["gdrive_get_spreadsheet_info"].handler({ spreadsheet_id: "s1" }, {});

    (sheets.deleteDimension as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    const result = await tools["gdrive_delete_rows_columns"].handler(
      { spreadsheet_id: "s1", sheet_title: "Sheet1", dimension: "COLUMNS", start_index: 0, count: 1 },
      {},
    );

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.dimension).toBe("COLUMNS");
    expect(data.startIndex).toBe(0);
    expect(data.count).toBe(1);
  });
});
