import { describe, it, expect, vi, beforeEach } from "vitest";
import { DriveClient, DriveAPIError, buildSearchQuery } from "../src/client.js";

function makeMockDrive(overrides: Record<string, unknown> = {}) {
  return {
    files: {
      list: vi.fn(),
      get: vi.fn(),
      export: vi.fn(),
      ...overrides,
    },
  };
}

function makeClient(mockDrive: ReturnType<typeof makeMockDrive>): DriveClient {
  const client = new DriveClient({} as any);
  (client as any).drive = mockDrive;
  return client;
}

describe("buildSearchQuery", () => {
  it("wraps plain text in fullText contains", () => {
    expect(buildSearchQuery("quarterly report")).toBe(
      "fullText contains 'quarterly report' and trashed = false",
    );
  });

  it("escapes single quotes in plain text", () => {
    expect(buildSearchQuery("it's here")).toBe(
      "fullText contains 'it\\'s here' and trashed = false",
    );
  });

  it("treats text with Drive keywords but no operators as plain text", () => {
    expect(buildSearchQuery("parents meeting notes")).toBe(
      "fullText contains 'parents meeting notes' and trashed = false",
    );
    expect(buildSearchQuery("name ideas for project")).toBe(
      "fullText contains 'name ideas for project' and trashed = false",
    );
  });

  it("passes Drive query syntax through unchanged", () => {
    expect(buildSearchQuery("name contains 'budget'")).toBe(
      "name contains 'budget' and trashed = false",
    );
  });

  it("wraps or-queries in parens before appending trashed filter", () => {
    expect(buildSearchQuery("mimeType='application/pdf' or mimeType='text/plain'")).toBe(
      "(mimeType='application/pdf' or mimeType='text/plain') and trashed = false",
    );
  });

  it("does not double-add trashed filter", () => {
    const q = "name contains 'x' and trashed = false";
    expect(buildSearchQuery(q)).toBe(q);
  });

  it("preserves explicit trashed = true", () => {
    const q = "name contains 'x' and trashed = true";
    expect(buildSearchQuery(q)).toBe(q);
  });
});

describe("DriveClient.search", () => {
  it("returns formatted file summaries", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.list.mockResolvedValue({
      data: {
        files: [
          { id: "f1", name: "doc.txt", mimeType: "text/plain", size: "100" },
        ],
        nextPageToken: "tok2",
      },
    });
    const client = makeClient(mockDrive);

    const result = await client.search("test query");

    expect(result.files).toHaveLength(1);
    expect(result.files[0]).toEqual(
      expect.objectContaining({ id: "f1", name: "doc.txt", mimeType: "text/plain" }),
    );
    expect(result.nextPageToken).toBe("tok2");
    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({
        q: "fullText contains 'test query' and trashed = false",
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      }),
    );
  });

  it("passes Drive query syntax to API", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: { files: [] } });
    const client = makeClient(mockDrive);

    await client.search("mimeType='application/pdf'");

    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({
        q: "mimeType='application/pdf' and trashed = false",
      }),
    );
  });

  it("returns empty array when no files match", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: {} });
    const client = makeClient(mockDrive);

    const result = await client.search("nothing");

    expect(result.files).toEqual([]);
    expect(result.nextPageToken).toBeUndefined();
  });
});

describe("DriveClient.readFile", () => {
  it("exports Google Docs as markdown", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: { id: "f1", name: "doc", mimeType: "application/vnd.google-apps.document" },
    });
    mockDrive.files.export.mockResolvedValue({ data: "# Hello" });
    const client = makeClient(mockDrive);

    const result = await client.readFile("f1");

    expect(result.mimeType).toBe("text/markdown");
    expect(result.content).toBe("# Hello");
    expect(mockDrive.files.export).toHaveBeenCalledWith(
      expect.objectContaining({ fileId: "f1", mimeType: "text/markdown" }),
    );
  });

  it("exports Google Sheets as CSV with first-sheet note", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: { id: "f1", name: "sheet", mimeType: "application/vnd.google-apps.spreadsheet" },
    });
    mockDrive.files.export.mockResolvedValue({ data: "a,b,c\n1,2,3" });
    const client = makeClient(mockDrive);

    const result = await client.readFile("f1");

    expect(result.mimeType).toBe("text/csv");
    expect(result.content).toContain("first sheet only");
    expect(result.content).toContain("a,b,c");
  });

  it("exports Google Slides as plain text", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: { id: "f1", name: "slides", mimeType: "application/vnd.google-apps.presentation" },
    });
    mockDrive.files.export.mockResolvedValue({ data: "Slide 1 content" });
    const client = makeClient(mockDrive);

    const result = await client.readFile("f1");

    expect(result.mimeType).toBe("text/plain");
    expect(result.content).toBe("Slide 1 content");
  });

  it("downloads text files directly via alt=media", async () => {
    const mockDrive = makeMockDrive();
    // First call: getFile metadata
    mockDrive.files.get
      .mockResolvedValueOnce({
        data: { id: "f1", name: "code.js", mimeType: "application/javascript" },
      })
      // Second call: download content
      .mockResolvedValueOnce({ data: "console.log('hi');" });
    const client = makeClient(mockDrive);

    const result = await client.readFile("f1");

    expect(result.mimeType).toBe("application/javascript");
    expect(result.content).toBe("console.log('hi');");
  });

  it("throws for binary files", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: {
        id: "f1",
        name: "image.png",
        mimeType: "image/png",
        webViewLink: "https://drive.google.com/file/d/f1",
      },
    });
    const client = makeClient(mockDrive);

    await expect(client.readFile("f1")).rejects.toThrow(DriveAPIError);
    await expect(client.readFile("f1")).rejects.toThrow(/binary file/i);
  });

  it("truncates content exceeding max_chars", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: { id: "f1", name: "doc", mimeType: "application/vnd.google-apps.document" },
    });
    mockDrive.files.export.mockResolvedValue({ data: "x".repeat(200) });
    const client = makeClient(mockDrive);

    const result = await client.readFile("f1", 50);

    expect(result.truncated).toBe(true);
    expect(result.content).toContain("[truncated]");
    expect(result.content.length).toBeLessThan(200);
  });

  it("returns drawing placeholder for PNG export", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: { id: "f1", name: "drawing", mimeType: "application/vnd.google-apps.drawing" },
    });
    mockDrive.files.export.mockResolvedValue({ data: Buffer.from("png-bytes") });
    const client = makeClient(mockDrive);

    const result = await client.readFile("f1");

    expect(result.mimeType).toBe("image/png");
    expect(result.content).toContain("Drawing exported as PNG");
  });
});

describe("DriveClient.readFile — shortcuts", () => {
  it("follows shortcut to target file", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get
      .mockResolvedValueOnce({
        data: {
          id: "shortcut1",
          name: "My Shortcut",
          mimeType: "application/vnd.google-apps.shortcut",
          shortcutDetails: { targetId: "real1", targetMimeType: "application/vnd.google-apps.document" },
        },
      })
      .mockResolvedValueOnce({
        data: { id: "real1", name: "Real Doc", mimeType: "application/vnd.google-apps.document" },
      });
    mockDrive.files.export.mockResolvedValue({ data: "# Target content" });
    const client = makeClient(mockDrive);

    const result = await client.readFile("shortcut1");

    expect(result.content).toBe("# Target content");
    expect(result.mimeType).toBe("text/markdown");
  });

  it("throws when shortcut has no target", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: {
        id: "shortcut1",
        name: "Broken Shortcut",
        mimeType: "application/vnd.google-apps.shortcut",
      },
    });
    const client = makeClient(mockDrive);

    await expect(client.readFile("shortcut1")).rejects.toThrow(/Shortcut target/);
  });
});

describe("DriveClient.listFiles", () => {
  it("scopes query to parent folder", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.list.mockResolvedValue({
      data: { files: [{ id: "f1", name: "sub.txt", mimeType: "text/plain" }] },
    });
    const client = makeClient(mockDrive);

    await client.listFiles("folder123");

    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({
        q: "'folder123' in parents and trashed = false",
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      }),
    );
  });

  it("defaults to root folder with modifiedTime desc ordering", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: { files: [] } });
    const client = makeClient(mockDrive);

    await client.listFiles();

    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({
        q: "'root' in parents and trashed = false",
        orderBy: "modifiedTime desc",
      }),
    );
  });
});
