import { describe, it, expect, vi, beforeEach } from "vitest";
import { createServer } from "../src/server.js";
import { DriveClient } from "../src/client.js";
import { SheetsClient } from "../src/sheets-client.js";
import { DocsClient, NormalizedDocument } from "../src/docs-client.js";

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

function makeMockDocsClient(): DocsClient {
  return {
    getDocument: vi.fn(),
    createDocument: vi.fn(),
    getRevisionId: vi.fn(),
    insertText: vi.fn(),
    replaceText: vi.fn(),
    replaceAllText: vi.fn(),
    deleteText: vi.fn(),
    updateTextStyle: vi.fn(),
    updateParagraphStyle: vi.fn(),
    updateList: vi.fn(),
    renameDocument: vi.fn(),
    duplicateDocument: vi.fn(),
  } as unknown as DocsClient;
}

function makeServer() {
  const drive = makeMockDriveClient();
  const sheets = makeMockSheetsClient();
  const docs = makeMockDocsClient();
  const server = createServer(drive, sheets, docs);
  return { server, drive, sheets, docs };
}

function getTools(server: ReturnType<typeof createServer>) {
  return (server as any)._registeredTools as Record<string, any>;
}

function makeDocumentMetadata(revisionId: string = "rev-1"): NormalizedDocument {
  return {
    documentId: "doc1",
    title: "Doc",
    documentUrl: "https://docs.google.com/document/d/doc1/edit",
    revisionId,
    contentTruncated: false,
    tabs: [
      {
        tabId: "tab-1",
        title: "Main",
        index: 0,
        nestingLevel: 0,
      },
      {
        tabId: "tab-2",
        title: "Notes",
        index: 1,
        nestingLevel: 0,
      },
    ],
  };
}

function makeDocumentContent(revisionId: string = "rev-1"): NormalizedDocument {
  return {
    documentId: "doc1",
    title: "Doc",
    documentUrl: "https://docs.google.com/document/d/doc1/edit",
    revisionId,
    contentTruncated: false,
    tabs: [
      {
        tabId: "tab-1",
        title: "Main",
        index: 0,
        nestingLevel: 0,
        paragraphs: [
          {
            startIndex: 1,
            endIndex: 7,
            text: "Hello\n",
            list: null,
            elements: [
              {
                type: "textRun",
                startIndex: 1,
                endIndex: 7,
                text: "Hello\n",
                textStyle: null,
              },
            ],
          },
          {
            startIndex: 7,
            endIndex: 13,
            text: "Hello\n",
            list: null,
            elements: [
              {
                type: "textRun",
                startIndex: 7,
                endIndex: 13,
                text: "Hello\n",
                textStyle: null,
              },
            ],
          },
        ],
      },
      {
        tabId: "tab-2",
        title: "Notes",
        index: 1,
        nestingLevel: 0,
        paragraphs: [
          {
            startIndex: 1,
            endIndex: 15,
            text: "Status: Draft\n",
            list: null,
            elements: [
              {
                type: "textRun",
                startIndex: 1,
                endIndex: 15,
                text: "Status: Draft\n",
                textStyle: null,
              },
            ],
          },
        ],
      },
    ],
  };
}

describe("Docs server behavior", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("rejects existing-document write tools until the document has been read", async () => {
    const { server } = makeServer();
    const tools = getTools(server);
    const argsByTool: Record<string, Record<string, unknown>> = {
      gdrive_insert_doc_text: { document_id: "doc1", text: "Hello", position: "end", match_case: true, conflict_mode: "strict" },
      gdrive_replace_doc_text: {
        document_id: "doc1",
        replacement_text: "Updated",
        target_text: "Hello",
        match_case: true,
        conflict_mode: "strict",
      },
      gdrive_replace_all_doc_text: {
        document_id: "doc1",
        old_text: "Hello",
        new_text: "Updated",
        all_tabs: false,
        match_case: true,
        conflict_mode: "strict",
      },
      gdrive_delete_doc_text: {
        document_id: "doc1",
        target_text: "Hello",
        match_case: true,
        conflict_mode: "strict",
      },
      gdrive_update_doc_text_style: {
        document_id: "doc1",
        target_text: "Hello",
        bold: true,
        match_case: true,
        conflict_mode: "strict",
      },
      gdrive_update_doc_paragraph_style: {
        document_id: "doc1",
        target_text: "Hello",
        named_style_type: "HEADING_2",
        match_case: true,
        conflict_mode: "strict",
      },
      gdrive_update_doc_list: {
        document_id: "doc1",
        target_text: "Hello",
        preset: "BULLETED",
        match_case: true,
        conflict_mode: "strict",
      },
      gdrive_rename_doc: { document_id: "doc1", new_title: "Renamed" },
      gdrive_duplicate_doc: { document_id: "doc1" },
    };

    for (const [toolName, args] of Object.entries(argsByTool)) {
      const result = await tools[toolName].handler(args, {});
      expect(result.isError, toolName).toBe(true);
      expect(result.content[0].text, toolName).toContain("must read this document");
    }
  });

  it("unlocks document writes after gdrive_get_document_info", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue(
      makeDocumentMetadata(),
    );
    (docs.renameDocument as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      title: "Renamed",
      documentUrl: "https://docs.google.com/document/d/doc1/edit",
    });

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: false, max_chars: 20_000, max_paragraphs: 200 },
      {},
    );
    (docs.getDocument as ReturnType<typeof vi.fn>).mockClear();
    (docs.getRevisionId as ReturnType<typeof vi.fn>).mockClear();
    const result = await tools["gdrive_rename_doc"].handler(
      { document_id: "doc1", new_title: "Renamed" },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(docs.renameDocument).toHaveBeenCalledWith("doc1", "Renamed");
    expect(docs.getDocument).not.toHaveBeenCalled();
    expect(docs.getRevisionId).not.toHaveBeenCalled();
  });

  it("unlocks document writes after gdrive_read_file and caches the revision", async () => {
    const { server, drive, docs } = makeServer();
    const tools = getTools(server);

    (drive.getFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      id: "doc1",
      name: "Doc",
      mimeType: "application/vnd.google-apps.document",
    });
    (drive.readFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      content: "# Hello",
      mimeType: "text/markdown",
      truncated: false,
    });
    (docs.getRevisionId as ReturnType<typeof vi.fn>).mockResolvedValue("rev-1");
    (docs.renameDocument as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      title: "Renamed",
      documentUrl: "https://docs.google.com/document/d/doc1/edit",
    });

    await tools["gdrive_read_file"].handler({ file_id: "doc1", max_chars: 100_000 }, {});
    const result = await tools["gdrive_rename_doc"].handler(
      { document_id: "doc1", new_title: "Renamed" },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(docs.getRevisionId).toHaveBeenCalledWith("doc1");
    expect(docs.renameDocument).toHaveBeenCalledWith("doc1", "Renamed");
  });

  it("shortcut reads unlock the target document ID, not the shortcut ID", async () => {
    const { server, drive, docs } = makeServer();
    const tools = getTools(server);

    (drive.getFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      id: "shortcut1",
      mimeType: "application/vnd.google-apps.shortcut",
      shortcutTarget: {
        id: "real-doc",
        mimeType: "application/vnd.google-apps.document",
      },
    });
    (drive.readFile as ReturnType<typeof vi.fn>).mockResolvedValue({
      content: "# Hello",
      mimeType: "text/markdown",
      truncated: false,
    });
    (docs.getRevisionId as ReturnType<typeof vi.fn>).mockResolvedValue("rev-1");
    (docs.renameDocument as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "real-doc",
      title: "Renamed",
      documentUrl: "https://docs.google.com/document/d/real-doc/edit",
    });

    await tools["gdrive_read_file"].handler({ file_id: "shortcut1", max_chars: 100_000 }, {});

    const goodResult = await tools["gdrive_rename_doc"].handler(
      { document_id: "real-doc", new_title: "Renamed" },
      {},
    );
    const badResult = await tools["gdrive_rename_doc"].handler(
      { document_id: "shortcut1", new_title: "Renamed" },
      {},
    );

    expect(goodResult.isError).toBeUndefined();
    expect(badResult.isError).toBe(true);
    expect(badResult.content[0].text).toContain("must read this document");
  });

  it("gdrive_create_doc auto-unlocks the new document for follow-up inserts", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.createDocument as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "new-doc",
      title: "New Doc",
      documentUrl: "https://docs.google.com/document/d/new-doc/edit",
      revisionId: "rev-new",
      folderId: "root",
    });
    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "new-doc",
      title: "New Doc",
      documentUrl: "https://docs.google.com/document/d/new-doc/edit",
      revisionId: "rev-new",
      contentTruncated: false,
      tabs: [{ tabId: "tab-1", title: "Main", index: 0, nestingLevel: 0 }],
    });
    (docs.insertText as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "new-doc",
      revisionId: "rev-next",
    });

    await tools["gdrive_create_doc"].handler(
      { title: "New Doc", folder_id: "root" },
      {},
    );
    const result = await tools["gdrive_insert_doc_text"].handler(
      {
        document_id: "new-doc",
        text: "Hello",
        position: "end",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(docs.insertText).toHaveBeenCalledWith(
      expect.objectContaining({
        documentId: "new-doc",
        revisionId: "rev-new",
        atEnd: true,
      }),
    );
  });

  it("replace_all defaults to the first tab and supports explicit all-tabs mode", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue(
      makeDocumentMetadata(),
    );
    (docs.replaceAllText as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      revisionId: "rev-2",
      replies: [{ replaceAllText: { occurrencesChanged: 1 } }],
    });

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: false, max_chars: 20_000, max_paragraphs: 200 },
      {},
    );

    await tools["gdrive_replace_all_doc_text"].handler(
      {
        document_id: "doc1",
        old_text: "Hello",
        new_text: "Hi",
        all_tabs: false,
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );
    await tools["gdrive_replace_all_doc_text"].handler(
      {
        document_id: "doc1",
        old_text: "Hello",
        new_text: "Hi",
        all_tabs: true,
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(docs.replaceAllText).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        documentId: "doc1",
        tabId: "tab-1",
        allTabs: false,
      }),
    );
    expect(docs.replaceAllText).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        documentId: "doc1",
        tabId: undefined,
        allTabs: true,
      }),
    );
  });

  it("supports case-insensitive anchor resolution when match_case is false", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValueOnce(
      makeDocumentContent(),
    );
    (docs.replaceText as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      revisionId: "rev-2",
    });

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: true, tab_id: "tab-1", max_chars: 20_000, max_paragraphs: 200 },
      {},
    );

    (docs.getDocument as ReturnType<typeof vi.fn>).mockClear();
    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue(
      makeDocumentMetadata(),
    );

    const result = await tools["gdrive_replace_doc_text"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "hello",
        occurrence: 1,
        replacement_text: "Hi",
        match_case: false,
        conflict_mode: "strict",
      },
      {},
    );

    expect(result.isError).toBeUndefined();
    expect(docs.replaceText).toHaveBeenCalledWith(
      expect.objectContaining({
        startIndex: 1,
        endIndex: 6,
      }),
    );
  });

  it("reuses cached structured content for anchor resolution when the revision still matches", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValueOnce(
      makeDocumentContent(),
    );
    (docs.replaceText as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      revisionId: "rev-2",
    });

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: true, tab_id: "tab-1", max_chars: 20_000, max_paragraphs: 200 },
      {},
    );

    (docs.getDocument as ReturnType<typeof vi.fn>).mockClear();
    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue(makeDocumentMetadata());

    await tools["gdrive_replace_doc_text"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "Hello",
        occurrence: 1,
        replacement_text: "Hi",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(docs.getDocument).toHaveBeenCalledTimes(1);
    expect(docs.getDocument).toHaveBeenCalledWith("doc1", { includeContent: false });
    expect(docs.replaceText).toHaveBeenCalledWith(
      expect.objectContaining({
        startIndex: 1,
        endIndex: 6,
      }),
    );
  });

  it("fetches fresh structured content when anchor resolution misses the cache", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce(makeDocumentMetadata())
      .mockResolvedValueOnce(makeDocumentMetadata())
      .mockResolvedValueOnce(makeDocumentContent());
    (docs.replaceText as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      revisionId: "rev-2",
    });

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: false, max_chars: 20_000, max_paragraphs: 200 },
      {},
    );
    await tools["gdrive_replace_doc_text"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "Hello",
        occurrence: 1,
        replacement_text: "Hi",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(docs.getDocument).toHaveBeenCalledTimes(3);
    expect(docs.getDocument).toHaveBeenLastCalledWith(
      "doc1",
      expect.objectContaining({
        includeContent: true,
        tabId: "tab-1",
      }),
    );
  });

  it("returns actionable anchor errors for ambiguous, missing, and out-of-range matches", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce(makeDocumentContent())
      .mockResolvedValue(makeDocumentMetadata());

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: true, tab_id: "tab-1", max_chars: 20_000, max_paragraphs: 200 },
      {},
    );

    const ambiguous = await tools["gdrive_replace_doc_text"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "Hello",
        replacement_text: "Hi",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );
    const missing = await tools["gdrive_replace_doc_text"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "Missing",
        replacement_text: "Hi",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );
    const outOfRange = await tools["gdrive_replace_doc_text"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "Hello",
        occurrence: 3,
        replacement_text: "Hi",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(ambiguous.content[0].text).toContain("Found 2 matches");
    expect(missing.content[0].text).toContain("Inspect current content");
    expect(outOfRange.content[0].text).toContain("out of range");
  });

  it("fails fast for paragraph style requests with no style fields", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue(
      makeDocumentMetadata(),
    );

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: false, max_chars: 20_000, max_paragraphs: 200 },
      {},
    );

    (docs.getDocument as ReturnType<typeof vi.fn>).mockClear();

    const result = await tools["gdrive_update_doc_paragraph_style"].handler(
      {
        document_id: "doc1",
        target_text: "Hello",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain(
      "At least one paragraph style parameter must be provided.",
    );
    expect(docs.getDocument).not.toHaveBeenCalled();
  });

  it("fails fast for text style requests with no style fields", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue(
      makeDocumentMetadata(),
    );

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: false, max_chars: 20_000, max_paragraphs: 200 },
      {},
    );

    (docs.getDocument as ReturnType<typeof vi.fn>).mockClear();

    const result = await tools["gdrive_update_doc_text_style"].handler(
      {
        document_id: "doc1",
        target_text: "Hello",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain(
      "At least one text style parameter must be provided.",
    );
    expect(docs.getDocument).not.toHaveBeenCalled();
  });

  it("snaps paragraph style and list updates to paragraph boundaries", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce(makeDocumentContent())
      .mockResolvedValueOnce(makeDocumentMetadata())
      .mockResolvedValueOnce(makeDocumentMetadata("rev-2"))
      .mockResolvedValueOnce(makeDocumentContent("rev-2"));
    (docs.updateParagraphStyle as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      revisionId: "rev-2",
    });
    (docs.updateList as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "doc1",
      revisionId: "rev-3",
    });

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: true, tab_id: "tab-1", max_chars: 20_000, max_paragraphs: 200 },
      {},
    );

    await tools["gdrive_update_doc_paragraph_style"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "Hello",
        occurrence: 1,
        named_style_type: "HEADING_2",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );
    await tools["gdrive_update_doc_list"].handler(
      {
        document_id: "doc1",
        tab_id: "tab-1",
        target_text: "Hello",
        occurrence: 2,
        preset: "CHECKBOX",
        match_case: true,
        conflict_mode: "strict",
      },
      {},
    );

    expect(docs.updateParagraphStyle).toHaveBeenCalledWith(
      expect.objectContaining({
        startIndex: 1,
        endIndex: 7,
      }),
    );
    expect(docs.updateList).toHaveBeenCalledWith(
      expect.objectContaining({
        startIndex: 7,
        endIndex: 13,
        preset: "CHECKBOX",
      }),
    );
  });

  it("duplicate_doc seeds the copied document into the read set", async () => {
    const { server, docs } = makeServer();
    const tools = getTools(server);

    (docs.getDocument as ReturnType<typeof vi.fn>).mockResolvedValue(
      makeDocumentMetadata(),
    );
    (docs.duplicateDocument as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "copy-1",
      title: "Copy",
      documentUrl: "https://docs.google.com/document/d/copy-1/edit",
      revisionId: "rev-copy",
      sourceDocumentId: "doc1",
    });
    (docs.renameDocument as ReturnType<typeof vi.fn>).mockResolvedValue({
      documentId: "copy-1",
      title: "Copy Renamed",
      documentUrl: "https://docs.google.com/document/d/copy-1/edit",
    });

    await tools["gdrive_get_document_info"].handler(
      { document_id: "doc1", include_content: false, max_chars: 20_000, max_paragraphs: 200 },
      {},
    );
    await tools["gdrive_duplicate_doc"].handler(
      { document_id: "doc1" },
      {},
    );
    const renameResult = await tools["gdrive_rename_doc"].handler(
      { document_id: "copy-1", new_title: "Copy Renamed" },
      {},
    );

    expect(renameResult.isError).toBeUndefined();
    expect(docs.renameDocument).toHaveBeenCalledWith("copy-1", "Copy Renamed");
  });
});
