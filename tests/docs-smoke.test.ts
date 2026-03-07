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
    "creates a doc and inserts text at the start through the MCP server flow",
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
    60_000,
  );
});
