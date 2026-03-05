import { describe, it, expect, vi, beforeEach } from "vitest";
import { DriveClient, DriveAPIError } from "../src/client.js";

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
        q: "test query",
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
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

  it("defaults to root folder", async () => {
    const mockDrive = makeMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: { files: [] } });
    const client = makeClient(mockDrive);

    await client.listFiles();

    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({
        q: "'root' in parents and trashed = false",
      }),
    );
  });
});
