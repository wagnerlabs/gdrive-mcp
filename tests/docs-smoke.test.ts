import { describe, expect, it } from "vitest";
import { google } from "googleapis";
import { loadCredentials } from "../src/auth.js";
import { DriveClient } from "../src/client.js";
import { DocsClient } from "../src/docs-client.js";
import { createServer } from "../src/server.js";
import { SheetsClient } from "../src/sheets-client.js";

const RUN_LIVE_GOOGLE_TESTS = process.env.RUN_LIVE_GOOGLE_TESTS === "1";
const describeLive = RUN_LIVE_GOOGLE_TESTS ? describe : describe.skip;

function getTools(server: ReturnType<typeof createServer>) {
  return (server as any)._registeredTools as Record<string, any>;
}

function parseJsonResult(result: {
  content: Array<{ type: "text"; text: string }>;
  isError?: true;
}): any {
  if (result.isError) {
    throw new Error(result.content[0]?.text ?? "Unknown MCP error");
  }
  return JSON.parse(result.content[0]?.text ?? "{}");
}

describeLive("live Docs smoke test", () => {
  it(
    "exercises revision-safe rich content, native lists, and native tables",
    async () => {
      const auth = await loadCredentials();
      const drive = google.drive({ version: "v3", auth });
      const server = createServer(
        new DriveClient(auth),
        new SheetsClient(auth),
        new DocsClient(auth),
      );
      const tools = getTools(server);
      const title = `docs-smoke-${Date.now()}`;
      let documentId: string | undefined;

      try {
        const created = parseJsonResult(
          await tools["gdrive_create_doc"].handler(
            { title, folder_id: "root" },
            {},
          ),
        );
        documentId = created.documentId;

        expect(created).toEqual(
          expect.objectContaining({
            documentId: expect.any(String),
            title,
          }),
        );

        const inserted = parseJsonResult(
          await tools["gdrive_insert_doc_text"].handler(
            {
              document_id: documentId,
              text: "hi",
              position: "start",
              match_case: true,
              conflict_mode: "strict",
            },
            {},
          ),
        );

        expect(inserted).toEqual(
          expect.objectContaining({
            documentId,
            insertedText: "hi",
            index: 1,
            revisionId: expect.any(String),
          }),
        );

        const documentInfo = parseJsonResult(
          await tools["gdrive_get_document_info"].handler(
            {
              document_id: documentId,
              include_content: true,
              max_chars: 1_000,
              max_paragraphs: 20,
            },
            {},
          ),
        );

        expect(documentInfo).toEqual(
          expect.objectContaining({
            documentId,
            revisionId: expect.any(String),
          }),
        );
        expect(
          documentInfo.tabs[0]?.paragraphs?.some((paragraph: { text: string }) =>
            paragraph.text.includes("hi"),
          ),
        ).toBe(true);

        const tabId = documentInfo.tabs[0]?.tabId;
        expect(tabId).toEqual(expect.any(String));

        const rich = parseJsonResult(
          await tools["gdrive_insert_doc_content"].handler(
            {
              document_id: documentId,
              tab_id: tabId,
              position: "end",
              blocks: [
                {
                  segments: [{ text: "Live smoke heading", bold: true }],
                  named_style_type: "HEADING_2",
                  space_below_points: 6,
                },
                {
                  segments: [{ text: "First item" }],
                  list: { preset: "NUMBERED", nesting_level: 0 },
                },
                {
                  segments: [{ text: "Nested item" }],
                  list: { preset: "NUMBERED", nesting_level: 1 },
                },
              ],
              match_case: true,
              inherit_neighbor_style: false,
              conflict_mode: "strict",
            },
            {},
          ),
        );
        expect(rich).toEqual(expect.objectContaining({
          documentId,
          blocksInserted: 3,
          revisionId: expect.any(String),
        }));

        const insertedTable = parseJsonResult(
          await tools["gdrive_insert_doc_table"].handler(
            {
              document_id: documentId,
              tab_id: tabId,
              rows: 2,
              columns: 2,
              values: [["A", "B"], ["C", "D"]],
              position: "end",
              match_case: true,
              conflict_mode: "strict",
            },
            {},
          ),
        );
        expect(insertedTable.table).toEqual(expect.objectContaining({
          type: "table",
          rows: 2,
          columns: 2,
          tablePath: expect.any(Array),
        }));

        const structured = parseJsonResult(
          await tools["gdrive_get_document_content"].handler(
            {
              document_id: documentId,
              tab_id: tabId,
              max_blocks: 50,
            },
            {},
          ),
        );
        expect(structured.tab.blocks).toEqual(expect.arrayContaining([
          expect.objectContaining({ type: "table", rows: 2, columns: 2 }),
        ]));
      } finally {
        if (documentId) {
          await drive.files.update({
            fileId: documentId,
            requestBody: { trashed: true },
            fields: "id,trashed",
            supportsAllDrives: true,
          });
        }
      }
    },
    120_000,
  );
});
