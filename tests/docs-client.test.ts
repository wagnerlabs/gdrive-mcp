import { describe, it, expect, vi } from "vitest";
import { DocsClient } from "../src/docs-client.js";

function makeMockDocs() {
  return {
    documents: {
      get: vi.fn(),
      create: vi.fn(),
      batchUpdate: vi.fn(),
    },
  };
}

function makeMockDrive() {
  return {
    files: {
      get: vi.fn(),
      update: vi.fn(),
      copy: vi.fn(),
    },
  };
}

function makeClient(
  mockDocs: ReturnType<typeof makeMockDocs>,
  mockDrive: ReturnType<typeof makeMockDrive>,
): DocsClient {
  const client = new DocsClient({} as any);
  (client as any).docs = mockDocs;
  (client as any).drive = mockDrive;
  return client;
}

describe("DocsClient.getDocument", () => {
  it("returns normalized metadata and content with tab-aware reads", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDocs.documents.get.mockResolvedValue({
      data: {
        documentId: "doc1",
        title: "Project Brief",
        revisionId: "rev-1",
        tabs: [
          {
            tabProperties: { tabId: "tab-1", title: "Main", index: 0, nestingLevel: 0 },
            documentTab: {
              body: {
                content: [
                  {
                    startIndex: 1,
                    endIndex: 7,
                    paragraph: {
                      elements: [
                        {
                          startIndex: 1,
                          endIndex: 7,
                          textRun: {
                            content: "Hello\n",
                            textStyle: { bold: true },
                          },
                        },
                      ],
                      paragraphStyle: {
                        namedStyleType: "HEADING_1",
                        alignment: "CENTER",
                      },
                    },
                  },
                  {
                    startIndex: 7,
                    endIndex: 8,
                    table: {},
                  },
                ],
              },
              lists: {},
            },
          },
        ],
      },
    });
    const client = makeClient(mockDocs, mockDrive);

    const result = await client.getDocument("doc1", {
      includeContent: true,
      tabId: "tab-1",
      maxChars: 10_000,
      maxParagraphs: 20,
    });

    expect(mockDocs.documents.get).toHaveBeenCalledWith(
      expect.objectContaining({
        documentId: "doc1",
        includeTabsContent: true,
        suggestionsViewMode: "SUGGESTIONS_INLINE",
      }),
    );
    expect(result).toEqual(
      expect.objectContaining({
        documentId: "doc1",
        title: "Project Brief",
        revisionId: "rev-1",
        documentUrl: "https://docs.google.com/document/d/doc1/edit",
      }),
    );
    expect(result.tabs[0]).toEqual(
      expect.objectContaining({
        tabId: "tab-1",
        title: "Main",
      }),
    );
    expect(result.tabs[0].paragraphs?.[0]).toEqual(
      expect.objectContaining({
        text: "Hello\n",
        namedStyleType: "HEADING_1",
        alignment: "CENTER",
      }),
    );
    expect(result.tabs[0].paragraphs?.[1].elements[0]).toEqual(
      expect.objectContaining({
        type: "placeholder",
        placeholderKind: "table",
      }),
    );
  });

  it("uses metadata-first field selection when content is omitted", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDocs.documents.get.mockResolvedValue({
      data: {
        documentId: "doc1",
        title: "Doc",
        revisionId: "rev-1",
        tabs: [{ tabProperties: { tabId: "tab-1", title: "Main", index: 0, nestingLevel: 0 } }],
      },
    });
    const client = makeClient(mockDocs, mockDrive);

    await client.getDocument("doc1");

    expect(mockDocs.documents.get).toHaveBeenCalledWith(
      expect.objectContaining({
        fields: expect.stringContaining("tabProperties"),
      }),
    );
    expect(mockDocs.documents.get.mock.calls[0][0].fields).not.toContain("textRun(content");
  });
});

describe("DocsClient writes", () => {
  it("creates a document and moves it into a folder", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDocs.documents.create.mockResolvedValue({
      data: {
        documentId: "new-doc",
        title: "New Doc",
        revisionId: "rev-new",
      },
    });
    mockDrive.files.get.mockResolvedValue({ data: { parents: ["root"] } });
    mockDrive.files.update.mockResolvedValue({ data: { id: "new-doc" } });
    const client = makeClient(mockDocs, mockDrive);

    const result = await client.createDocument("New Doc", "folder123");

    expect(mockDrive.files.update).toHaveBeenCalledWith(
      expect.objectContaining({
        fileId: "new-doc",
        addParents: "folder123",
        removeParents: "root",
      }),
    );
    expect(result).toEqual(
      expect.objectContaining({
        documentId: "new-doc",
        folderId: "folder123",
        revisionId: "rev-new",
      }),
    );
  });

  it("shapes replaceText requests with strict write control", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDocs.documents.batchUpdate.mockResolvedValue({
      data: {
        documentId: "doc1",
        writeControl: { requiredRevisionId: "rev-2" },
      },
    });
    const client = makeClient(mockDocs, mockDrive);

    const result = await client.replaceText({
      documentId: "doc1",
      tabId: "tab-1",
      startIndex: 5,
      endIndex: 10,
      text: "Updated",
      revisionId: "rev-1",
      conflictMode: "strict",
    });

    const call = mockDocs.documents.batchUpdate.mock.calls[0][0];
    expect(call.requestBody.writeControl).toEqual({ requiredRevisionId: "rev-1" });
    expect(call.requestBody.requests[0]).toEqual({
      deleteContentRange: {
        range: { tabId: "tab-1", startIndex: 5, endIndex: 10 },
      },
    });
    expect(call.requestBody.requests[1]).toEqual({
      insertText: {
        text: "Updated",
        location: { tabId: "tab-1", index: 5 },
      },
    });
    expect(result.revisionId).toBe("rev-2");
  });

  it("shapes replaceAllText requests with tab scoping and merge write control", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDocs.documents.batchUpdate.mockResolvedValue({
      data: {
        documentId: "doc1",
        writeControl: { targetRevisionId: "rev-2" },
      },
    });
    const client = makeClient(mockDocs, mockDrive);

    await client.replaceAllText({
      documentId: "doc1",
      searchText: "Acme",
      replaceText: "Wagner Labs",
      tabId: "tab-1",
      matchCase: false,
      revisionId: "rev-1",
      conflictMode: "merge",
    });

    const request = mockDocs.documents.batchUpdate.mock.calls[0][0].requestBody.requests[0]
      .replaceAllText;
    expect(request.containsText).toEqual({
      text: "Acme",
      matchCase: false,
    });
    expect(request.tabsCriteria).toEqual({ tabIds: ["tab-1"] });
    expect(mockDocs.documents.batchUpdate.mock.calls[0][0].requestBody.writeControl).toEqual({
      targetRevisionId: "rev-1",
    });
  });

  it("builds UpdateTextStyleRequest field masks", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDocs.documents.batchUpdate.mockResolvedValue({
      data: {
        documentId: "doc1",
        writeControl: { requiredRevisionId: "rev-2" },
      },
    });
    const client = makeClient(mockDocs, mockDrive);

    await client.updateTextStyle({
      documentId: "doc1",
      tabId: "tab-1",
      startIndex: 1,
      endIndex: 6,
      bold: true,
      fontFamily: "Arial",
      fontSize: 14,
      foregroundColor: "#3366FF",
      linkUrl: "https://example.com",
      revisionId: "rev-1",
    });

    const request = mockDocs.documents.batchUpdate.mock.calls[0][0].requestBody.requests[0]
      .updateTextStyle;
    expect(request.fields).toContain("bold");
    expect(request.fields).toContain("weightedFontFamily.fontFamily");
    expect(request.fields).toContain("fontSize");
    expect(request.fields).toContain("foregroundColor");
    expect(request.fields).toContain("link");
    expect(request.textStyle.foregroundColor.color.rgbColor.red).toBeCloseTo(0.2, 1);
  });

  it("builds paragraph style and list update requests", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDocs.documents.batchUpdate.mockResolvedValue({
      data: {
        documentId: "doc1",
        writeControl: { requiredRevisionId: "rev-2" },
      },
    });
    const client = makeClient(mockDocs, mockDrive);

    await client.updateParagraphStyle({
      documentId: "doc1",
      tabId: "tab-1",
      startIndex: 1,
      endIndex: 20,
      namedStyleType: "HEADING_2",
      alignment: "JUSTIFIED",
      revisionId: "rev-1",
    });
    await client.updateList({
      documentId: "doc1",
      tabId: "tab-1",
      startIndex: 1,
      endIndex: 20,
      preset: "CHECKBOX",
      revisionId: "rev-1",
    });

    const paragraphRequest = mockDocs.documents.batchUpdate.mock.calls[0][0].requestBody
      .requests[0].updateParagraphStyle;
    const listRequest = mockDocs.documents.batchUpdate.mock.calls[1][0].requestBody.requests[0]
      .createParagraphBullets;

    expect(paragraphRequest.fields).toBe("namedStyleType,alignment");
    expect(paragraphRequest.paragraphStyle).toEqual({
      namedStyleType: "HEADING_2",
      alignment: "JUSTIFIED",
    });
    expect(listRequest.bulletPreset).toBe("BULLET_CHECKBOX");
  });

  it("duplicates a document and seeds the copied revision", async () => {
    const mockDocs = makeMockDocs();
    const mockDrive = makeMockDrive();
    mockDrive.files.get.mockResolvedValue({ data: { parents: ["folderA"] } });
    mockDrive.files.copy.mockResolvedValue({
      data: { id: "copy-1", name: "Copy" },
    });
    mockDocs.documents.get.mockResolvedValue({
      data: { revisionId: "rev-copy" },
    });
    const client = makeClient(mockDocs, mockDrive);

    const result = await client.duplicateDocument("doc1");

    expect(mockDrive.files.copy).toHaveBeenCalledWith(
      expect.objectContaining({
        fileId: "doc1",
        requestBody: expect.objectContaining({
          parents: ["folderA"],
        }),
      }),
    );
    expect(result).toEqual(
      expect.objectContaining({
        documentId: "copy-1",
        revisionId: "rev-copy",
        sourceDocumentId: "doc1",
      }),
    );
  });
});
