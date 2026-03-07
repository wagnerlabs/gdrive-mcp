import { describe, it, expect, vi } from "vitest";
import { SheetsClient } from "../src/sheets-client.js";
import { DriveAPIError } from "../src/client.js";

function makeMockSheets() {
  return {
    spreadsheets: {
      get: vi.fn(),
      create: vi.fn(),
      batchUpdate: vi.fn(),
      values: {
        get: vi.fn(),
        update: vi.fn(),
        append: vi.fn(),
        clear: vi.fn(),
      },
    },
  };
}

function makeClient(
  mock: ReturnType<typeof makeMockSheets>,
): SheetsClient {
  const client = new SheetsClient({} as any);
  (client as any).sheets = mock;
  return client;
}

// ── getSpreadsheet ───────────────────────────────────────────────────

describe("SheetsClient.getSpreadsheet", () => {
  it("returns formatted spreadsheet metadata", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "My Sheet", locale: "en_US" },
        spreadsheetUrl: "https://docs.google.com/spreadsheets/d/s1/edit",
        sheets: [
          {
            properties: {
              sheetId: 0,
              title: "Sheet1",
              index: 0,
              gridProperties: { rowCount: 1000, columnCount: 26 },
            },
          },
        ],
        namedRanges: [],
      },
    });
    const client = makeClient(mock);

    const result = await client.getSpreadsheet("s1");

    expect(result.spreadsheetId).toBe("s1");
    expect(result.title).toBe("My Sheet");
    expect(result.locale).toBe("en_US");
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0]).toEqual({
      sheetId: 0,
      title: "Sheet1",
      index: 0,
      rowCount: 1000,
      columnCount: 26,
    });
    expect(mock.spreadsheets.get).toHaveBeenCalledWith(
      expect.objectContaining({
        spreadsheetId: "s1",
        fields: expect.stringContaining("sheets.properties"),
      }),
    );
  });
});

// ── createSpreadsheet ────────────────────────────────────────────────

describe("SheetsClient.createSpreadsheet", () => {
  it("creates a spreadsheet with default sheet name", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.create.mockResolvedValue({
      data: {
        spreadsheetId: "new1",
        spreadsheetUrl: "https://docs.google.com/spreadsheets/d/new1/edit",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0 } }],
      },
    });
    const client = makeClient(mock);

    const result = await client.createSpreadsheet("New Spreadsheet");

    expect(result.spreadsheetId).toBe("new1");
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].title).toBe("Sheet1");
    expect(mock.spreadsheets.create).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: expect.objectContaining({
          properties: { title: "New Spreadsheet" },
        }),
      }),
    );
  });

  it("creates a spreadsheet with custom sheet names", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.create.mockResolvedValue({
      data: {
        spreadsheetId: "new2",
        spreadsheetUrl: "url",
        sheets: [
          { properties: { sheetId: 0, title: "Data", index: 0 } },
          { properties: { sheetId: 1, title: "Summary", index: 1 } },
        ],
      },
    });
    const client = makeClient(mock);

    const result = await client.createSpreadsheet("Report", ["Data", "Summary"]);

    expect(result.sheets).toHaveLength(2);
    expect(result.sheets.map((s) => s.title)).toEqual(["Data", "Summary"]);
  });
});

// ── getValues ────────────────────────────────────────────────────────

describe("SheetsClient.getValues", () => {
  it("returns values from the API", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.values.get.mockResolvedValue({
      data: { values: [["a", "b"], ["1", "2"]] },
    });
    const client = makeClient(mock);

    const result = await client.getValues("s1", "Sheet1!A1:B2");

    expect(result).toEqual([["a", "b"], ["1", "2"]]);
  });

  it("returns empty array when no values exist", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.values.get.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    const result = await client.getValues("s1", "Sheet1!A1:B2");

    expect(result).toEqual([]);
  });
});

// ── updateValues ─────────────────────────────────────────────────────

describe("SheetsClient.updateValues", () => {
  it("updates values and returns statistics", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.values.update.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        updatedRange: "Sheet1!A1:B2",
        updatedRows: 2,
        updatedColumns: 2,
        updatedCells: 4,
      },
    });
    const client = makeClient(mock);

    const result = await client.updateValues("s1", "Sheet1!A1:B2", [["a", "b"], ["c", "d"]]);

    expect(result.updatedCells).toBe(4);
    expect(mock.spreadsheets.values.update).toHaveBeenCalledWith(
      expect.objectContaining({
        spreadsheetId: "s1",
        range: "Sheet1!A1:B2",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [["a", "b"], ["c", "d"]] },
      }),
    );
  });

  it("passes RAW value input option", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.values.update.mockResolvedValue({
      data: { spreadsheetId: "s1", updatedRange: "A1", updatedRows: 1, updatedColumns: 1, updatedCells: 1 },
    });
    const client = makeClient(mock);

    await client.updateValues("s1", "Sheet1!A1", [["=SUM()"]], "RAW");

    expect(mock.spreadsheets.values.update).toHaveBeenCalledWith(
      expect.objectContaining({ valueInputOption: "RAW" }),
    );
  });
});

// ── appendValues ─────────────────────────────────────────────────────

describe("SheetsClient.appendValues", () => {
  it("appends with INSERT_ROWS", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.values.append.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        tableRange: "Sheet1!A1:B3",
        updates: { updatedRange: "Sheet1!A4:B4", updatedRows: 1, updatedCells: 2 },
      },
    });
    const client = makeClient(mock);

    const result = await client.appendValues("s1", "Sheet1!A:B", [["x", "y"]]);

    expect(result.tableRange).toBe("Sheet1!A1:B3");
    expect(result.updatedRows).toBe(1);
    expect(mock.spreadsheets.values.append).toHaveBeenCalledWith(
      expect.objectContaining({ insertDataOption: "INSERT_ROWS" }),
    );
  });
});

// ── clearValues ──────────────────────────────────────────────────────

describe("SheetsClient.clearValues", () => {
  it("clears the specified range", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.values.clear.mockResolvedValue({
      data: { spreadsheetId: "s1", clearedRange: "Sheet1!A1:C10" },
    });
    const client = makeClient(mock);

    const result = await client.clearValues("s1", "Sheet1!A1:C10");

    expect(result.clearedRange).toBe("Sheet1!A1:C10");
  });
});

// ── addSheet ─────────────────────────────────────────────────────────

describe("SheetsClient.addSheet", () => {
  it("adds a new sheet tab and returns its properties", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.batchUpdate.mockResolvedValue({
      data: {
        replies: [{ addSheet: { properties: { sheetId: 42, title: "NewTab", index: 1 } } }],
      },
    });
    const client = makeClient(mock);

    const result = await client.addSheet("s1", "NewTab");

    expect(result.sheetId).toBe(42);
    expect(result.title).toBe("NewTab");
    expect(result.index).toBe(1);
  });
});

// ── deleteSheet ──────────────────────────────────────────────────────

describe("SheetsClient.deleteSheet", () => {
  it("looks up sheet ID by title and deletes", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 7, title: "ToDelete", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    mock.spreadsheets.batchUpdate.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    await client.deleteSheet("s1", "ToDelete");

    expect(mock.spreadsheets.batchUpdate).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: { requests: [{ deleteSheet: { sheetId: 7 } }] },
      }),
    );
  });

  it("throws when sheet title is not found", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    const client = makeClient(mock);

    await expect(client.deleteSheet("s1", "NonExistent")).rejects.toThrow(
      /Sheet tab "NonExistent" not found/,
    );
  });
});

// ── renameSheet ──────────────────────────────────────────────────────

describe("SheetsClient.renameSheet", () => {
  it("renames a sheet tab and returns new info", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 3, title: "OldName", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    mock.spreadsheets.batchUpdate.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    const result = await client.renameSheet("s1", "OldName", "NewName");

    expect(result.sheetId).toBe(3);
    expect(result.title).toBe("NewName");
    expect(result.index).toBe(0);
    expect(mock.spreadsheets.batchUpdate).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId: 3, title: "NewName" },
                fields: "title",
              },
            },
          ],
        },
      }),
    );
  });
});

// ── insertDimension ──────────────────────────────────────────────────

describe("SheetsClient.insertDimension", () => {
  it("inserts rows with correct range", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    mock.spreadsheets.batchUpdate.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    await client.insertDimension("s1", "Sheet1", "ROWS", 2, 3);

    expect(mock.spreadsheets.batchUpdate).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: { sheetId: 0, dimension: "ROWS", startIndex: 2, endIndex: 5 },
                inheritFromBefore: true,
              },
            },
          ],
        },
      }),
    );
  });

  it("sets inheritFromBefore to false when inserting at index 0", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    mock.spreadsheets.batchUpdate.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    await client.insertDimension("s1", "Sheet1", "COLUMNS", 0, 1);

    expect(mock.spreadsheets.batchUpdate).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: { sheetId: 0, dimension: "COLUMNS", startIndex: 0, endIndex: 1 },
                inheritFromBefore: false,
              },
            },
          ],
        },
      }),
    );
  });
});

// ── deleteDimension ──────────────────────────────────────────────────

describe("SheetsClient.deleteDimension", () => {
  it("computes endIndex from startIndex + count", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    mock.spreadsheets.batchUpdate.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    await client.deleteDimension("s1", "Sheet1", "ROWS", 1, 2);

    expect(mock.spreadsheets.batchUpdate).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: { sheetId: 0, dimension: "ROWS", startIndex: 1, endIndex: 3 },
              },
            },
          ],
        },
      }),
    );
  });
});

// ── formatCells ──────────────────────────────────────────────────────

describe("SheetsClient.formatCells", () => {
  it("builds a RepeatCellRequest with the correct fields mask", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    mock.spreadsheets.batchUpdate.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    await client.formatCells("s1", "Sheet1!A1:B2", { bold: true, fontSize: 14 });

    const call = mock.spreadsheets.batchUpdate.mock.calls[0][0];
    const request = call.requestBody.requests[0].repeatCell;
    expect(request.range).toEqual({
      sheetId: 0,
      startRowIndex: 0,
      startColumnIndex: 0,
      endRowIndex: 2,
      endColumnIndex: 2,
    });
    expect(request.cell.userEnteredFormat.textFormat.bold).toBe(true);
    expect(request.cell.userEnteredFormat.textFormat.fontSize).toBe(14);
    expect(request.fields).toContain("userEnteredFormat.textFormat.bold");
    expect(request.fields).toContain("userEnteredFormat.textFormat.fontSize");
  });

  it("handles background color and alignment", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    mock.spreadsheets.batchUpdate.mockResolvedValue({ data: {} });
    const client = makeClient(mock);

    await client.formatCells("s1", "Sheet1!C3", {
      backgroundColor: "#FF0000",
      horizontalAlignment: "CENTER",
    });

    const request = mock.spreadsheets.batchUpdate.mock.calls[0][0].requestBody.requests[0].repeatCell;
    expect(request.cell.userEnteredFormat.backgroundColor).toEqual({
      red: 1,
      green: 0,
      blue: 0,
    });
    expect(request.cell.userEnteredFormat.horizontalAlignment).toBe("CENTER");
    expect(request.fields).toContain("userEnteredFormat.backgroundColor");
    expect(request.fields).toContain("userEnteredFormat.horizontalAlignment");
  });
});

// ── getSheetId ───────────────────────────────────────────────────────

describe("SheetsClient.getSheetId", () => {
  it("returns the numeric sheet ID for a given title", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [
          { properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } },
          { properties: { sheetId: 99, title: "Data", index: 1, gridProperties: { rowCount: 10, columnCount: 5 } } },
        ],
      },
    });
    const client = makeClient(mock);

    const id = await client.getSheetId("s1", "Data");

    expect(id).toBe(99);
  });

  it("throws DriveAPIError when title is not found", async () => {
    const mock = makeMockSheets();
    mock.spreadsheets.get.mockResolvedValue({
      data: {
        spreadsheetId: "s1",
        properties: { title: "T" },
        spreadsheetUrl: "url",
        sheets: [{ properties: { sheetId: 0, title: "Sheet1", index: 0, gridProperties: { rowCount: 10, columnCount: 5 } } }],
      },
    });
    const client = makeClient(mock);

    await expect(client.getSheetId("s1", "Missing")).rejects.toThrow(DriveAPIError);
    await expect(client.getSheetId("s1", "Missing")).rejects.toThrow(/Missing/);
  });
});

// ── Static helpers ───────────────────────────────────────────────────

describe("SheetsClient.parseA1Range", () => {
  it("parses a standard range", () => {
    expect(SheetsClient.parseA1Range("Sheet1!A1:C5")).toEqual({
      sheetTitle: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 4,
      endCol: 2,
    });
  });

  it("parses a quoted sheet name", () => {
    expect(SheetsClient.parseA1Range("'Q1 Budget'!A1:C5")).toEqual({
      sheetTitle: "Q1 Budget",
      startRow: 0,
      startCol: 0,
      endRow: 4,
      endCol: 2,
    });
  });

  it("parses a single cell as a 1x1 range", () => {
    expect(SheetsClient.parseA1Range("Sheet1!B2")).toEqual({
      sheetTitle: "Sheet1",
      startRow: 1,
      startCol: 1,
      endRow: 1,
      endCol: 1,
    });
  });

  it("handles multi-letter columns", () => {
    expect(SheetsClient.parseA1Range("Sheet1!AA1:AZ10")).toEqual({
      sheetTitle: "Sheet1",
      startRow: 0,
      startCol: 26,
      endRow: 9,
      endCol: 51,
    });
  });

  it("handles lowercase column letters", () => {
    expect(SheetsClient.parseA1Range("Sheet1!a1:c5")).toEqual({
      sheetTitle: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 4,
      endCol: 2,
    });
  });

  it("rejects open-ended ranges", () => {
    expect(() => SheetsClient.parseA1Range("Sheet1!A:C")).toThrow(DriveAPIError);
  });

  it("rejects ranges without sheet name", () => {
    expect(() => SheetsClient.parseA1Range("A1:C5")).toThrow(DriveAPIError);
  });

  it("rejects invalid format", () => {
    expect(() => SheetsClient.parseA1Range("garbage")).toThrow(DriveAPIError);
  });
});

describe("SheetsClient.hexToColor", () => {
  it("converts #FF0000 to red", () => {
    expect(SheetsClient.hexToColor("#FF0000")).toEqual({
      red: 1,
      green: 0,
      blue: 0,
    });
  });

  it("converts #00FF00 to green", () => {
    expect(SheetsClient.hexToColor("#00FF00")).toEqual({
      red: 0,
      green: 1,
      blue: 0,
    });
  });

  it("converts without # prefix", () => {
    expect(SheetsClient.hexToColor("0000FF")).toEqual({
      red: 0,
      green: 0,
      blue: 1,
    });
  });

  it("converts a mid-range color", () => {
    const color = SheetsClient.hexToColor("#808080");
    expect(color.red).toBeCloseTo(0.502, 2);
    expect(color.green).toBeCloseTo(0.502, 2);
    expect(color.blue).toBeCloseTo(0.502, 2);
  });
});
