import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ToolAnnotations } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { DriveClient, DriveAPIError } from "./client.js";
import { SheetsClient, FormatOptions } from "./sheets-client.js";
import {
  DocListPreset,
  DocNamedStyleType,
  DocParagraphAlignment,
  DocsClient,
  DocsConflictMode,
  NormalizedDocElement,
  NormalizedDocTab,
  NormalizedDocument,
} from "./docs-client.js";

const SAFE: ToolAnnotations = {
  readOnlyHint: true,
  destructiveHint: false,
  openWorldHint: true,
};

const DOC_PLACEHOLDER_CHAR = "\uFFFC";
const DOC_TERMINAL_NEWLINE_WARNING =
  "Excluded the trailing paragraph newline from the resolved range because Google Docs cannot delete the final newline of a segment.";
const HEX_COLOR_PATTERN = /^#?[0-9A-Fa-f]{6}$/;

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

function documentUrl(id: string): string {
  return `https://docs.google.com/document/d/${id}/edit`;
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

interface CachedDocumentTabContent {
  documentId: string;
  revisionId: string;
  tabId: string;
  tab: NormalizedDocTab;
  searchableText: string;
  baseIndex: number;
}

interface ResolvedDocumentRange {
  tabId: string;
  startIndex: number;
  endIndex: number;
  actualText: string;
  targetMode: "anchor" | "explicit";
  snapshot: CachedDocumentTabContent;
}

function docCacheKey(documentId: string, revisionId: string, tabId: string): string {
  return `${documentId}::${revisionId}::${tabId}`;
}

function coerceSegmentLength(text: string, expectedLength: number): string {
  if (expectedLength <= 0) {
    return "";
  }
  if (text.length === expectedLength) {
    return text;
  }
  if (text.length > expectedLength) {
    return text.slice(0, expectedLength);
  }
  return text + DOC_PLACEHOLDER_CHAR.repeat(expectedLength - text.length);
}

function searchableTextForElement(element: NormalizedDocElement): string {
  const expectedLength = Math.max(0, element.endIndex - element.startIndex);
  if (element.type === "textRun") {
    return coerceSegmentLength(element.text, expectedLength);
  }
  return DOC_PLACEHOLDER_CHAR.repeat(Math.max(1, expectedLength));
}

function buildSearchableTabText(tab: NormalizedDocTab): {
  searchableText: string;
  baseIndex: number;
} {
  const paragraphs = (tab.paragraphs ?? []).slice().sort((left, right) => left.startIndex - right.startIndex);
  if (paragraphs.length === 0) {
    return { searchableText: "", baseIndex: 1 };
  }

  let searchableText = "";
  const baseIndex = paragraphs[0].startIndex;
  let cursor = baseIndex;

  for (const paragraph of paragraphs) {
    if (paragraph.startIndex > cursor) {
      searchableText += DOC_PLACEHOLDER_CHAR.repeat(paragraph.startIndex - cursor);
      cursor = paragraph.startIndex;
    }

    for (const element of paragraph.elements) {
      if (element.startIndex > cursor) {
        searchableText += DOC_PLACEHOLDER_CHAR.repeat(element.startIndex - cursor);
        cursor = element.startIndex;
      }

      searchableText += searchableTextForElement(element);
      cursor = element.endIndex;
    }

    if (paragraph.endIndex > cursor) {
      searchableText += DOC_PLACEHOLDER_CHAR.repeat(paragraph.endIndex - cursor);
      cursor = paragraph.endIndex;
    }
  }

  return { searchableText, baseIndex };
}

function extractRangeText(
  snapshot: CachedDocumentTabContent,
  startIndex: number,
  endIndex: number,
): string {
  const startOffset = startIndex - snapshot.baseIndex;
  const endOffset = endIndex - snapshot.baseIndex;
  if (startOffset < 0 || endOffset < startOffset) {
    return "";
  }
  return snapshot.searchableText.slice(startOffset, endOffset);
}

function findTextMatches(
  haystack: string,
  needle: string,
  matchCase: boolean,
): Array<{ startOffset: number; endOffset: number }> {
  if (!needle) {
    return [];
  }

  const source = matchCase ? haystack : haystack.toLowerCase();
  const target = matchCase ? needle : needle.toLowerCase();
  const matches: Array<{ startOffset: number; endOffset: number }> = [];

  let fromIndex = 0;
  while (true) {
    const offset = source.indexOf(target, fromIndex);
    if (offset === -1) {
      break;
    }
    matches.push({ startOffset: offset, endOffset: offset + target.length });
    fromIndex = offset + Math.max(target.length, 1);
  }

  return matches;
}

export function createServer(
  driveClient: DriveClient,
  sheetsClient: SheetsClient,
  docsClient: DocsClient,
): McpServer {
  const server = new McpServer({
    name: "gdrive-mcp",
    version: "0.1.0",
  });

  const accessedSpreadsheets = new Set<string>();
  const accessedDocs = new Set<string>();
  const lastSeenDocRevision = new Map<string, string>();
  const structuredDocCache = new Map<string, CachedDocumentTabContent>();

  function evictStructuredDocumentCache(
    documentId: string,
    keepRevisionId?: string,
  ): void {
    for (const [cacheKey, cached] of structuredDocCache.entries()) {
      if (
        cached.documentId === documentId &&
        cached.revisionId !== keepRevisionId
      ) {
        structuredDocCache.delete(cacheKey);
      }
    }
  }

  function rememberDocumentRead(documentId: string, revisionId?: string): void {
    accessedDocs.add(documentId);
    if (revisionId) {
      const previousRevisionId = lastSeenDocRevision.get(documentId);
      lastSeenDocRevision.set(documentId, revisionId);
      if (previousRevisionId !== revisionId) {
        evictStructuredDocumentCache(documentId, revisionId);
      }
    }
  }

  function cacheStructuredDocument(document: NormalizedDocument): void {
    if (!document.revisionId || document.contentTruncated) {
      return;
    }

    for (const tab of document.tabs) {
      if (!tab.paragraphs) {
        continue;
      }
      const search = buildSearchableTabText(tab);
      structuredDocCache.set(docCacheKey(document.documentId, document.revisionId, tab.tabId), {
        documentId: document.documentId,
        revisionId: document.revisionId,
        tabId: tab.tabId,
        tab,
        searchableText: search.searchableText,
        baseIndex: search.baseIndex,
      });
    }
  }

  function rememberDocumentSnapshot(document: NormalizedDocument): void {
    rememberDocumentRead(document.documentId, document.revisionId);
    cacheStructuredDocument(document);
  }

  function unreadDocumentError(): Error {
    return new Error(
      "You must read this document before writing to it. " +
        "Use gdrive_read_file or gdrive_get_document_info first.",
    );
  }

  async function fetchDocumentMetadata(documentId: string): Promise<NormalizedDocument> {
    const metadata = await docsClient.getDocument(documentId, {
      includeContent: false,
    });
    rememberDocumentSnapshot(metadata);
    return metadata;
  }

  async function ensureDocumentRevision(documentId: string): Promise<string> {
    const cached = lastSeenDocRevision.get(documentId);
    if (cached) {
      return cached;
    }

    const metadata = await fetchDocumentMetadata(documentId);
    if (!metadata.revisionId) {
      throw new Error(
        "Could not determine the current document revision. " +
          "Make sure you have edit access, then read the document again.",
      );
    }
    return metadata.revisionId;
  }

  async function resolveDocumentTab(
    documentId: string,
    requestedTabId?: string,
  ): Promise<{ document: NormalizedDocument; tab: NormalizedDocTab }> {
    const document = await fetchDocumentMetadata(documentId);
    if (document.tabs.length === 0) {
      throw new Error("This document does not expose any editable tabs.");
    }

    const tab = requestedTabId
      ? document.tabs.find((candidate) => candidate.tabId === requestedTabId)
      : document.tabs[0];

    if (!tab) {
      throw new Error(`Tab "${requestedTabId}" not found in document.`);
    }

    return { document, tab };
  }

  function assertRangeWithinSnapshot(
    snapshot: CachedDocumentTabContent,
    startIndex: number,
    endIndex: number,
  ): void {
    const lowerBound = snapshot.baseIndex;
    const upperBound = snapshot.baseIndex + snapshot.searchableText.length;
    if (startIndex < lowerBound || endIndex > upperBound || startIndex >= endIndex) {
      throw new Error(
        `Range [${startIndex}, ${endIndex}) is outside the current content for tab "${snapshot.tab.title}". ` +
          "Use gdrive_get_document_info include_content=true to inspect the latest indices.",
      );
    }
  }

  // This helper is body/tab scoped today. If we later support headers,
  // footers, or footnotes, terminal newline detection needs to become
  // segment-aware rather than assuming the last paragraph in the tab body.
  function segmentTerminalNewlineIndex(
    snapshot: CachedDocumentTabContent,
  ): number | undefined {
    const paragraphs = snapshot.tab.paragraphs ?? [];
    const lastParagraph = paragraphs[paragraphs.length - 1];
    if (!lastParagraph?.text.endsWith("\n")) {
      return undefined;
    }
    return lastParagraph.endIndex - 1;
  }

  function rangeIncludesSegmentTerminalNewline(
    resolved: ResolvedDocumentRange,
  ): boolean {
    const terminalNewlineIndex = segmentTerminalNewlineIndex(resolved.snapshot);
    return (
      terminalNewlineIndex !== undefined &&
      resolved.endIndex === terminalNewlineIndex + 1 &&
      resolved.actualText.endsWith("\n")
    );
  }

  function terminalNewlineEditError(targetMode: ResolvedDocumentRange["targetMode"]): Error {
    const guidance =
      targetMode === "anchor"
        ? 'Retry with `target_text` that matches `paragraph.displayText`, or omit the trailing "\\n".'
        : 'Adjust `end_index` to exclude the trailing "\\n".';
    return new Error(
      "Google Docs does not allow deleting or replacing the final paragraph newline of the current tab. " +
        guidance,
    );
  }

  function normalizeRangeForTextMutation(resolved: ResolvedDocumentRange): {
    startIndex: number;
    endIndex: number;
    actualText: string;
    warnings?: string[];
  } {
    if (!rangeIncludesSegmentTerminalNewline(resolved)) {
      return {
        startIndex: resolved.startIndex,
        endIndex: resolved.endIndex,
        actualText: resolved.actualText,
      };
    }

    const trimmedEndIndex = resolved.endIndex - 1;
    if (resolved.targetMode !== "anchor" || trimmedEndIndex <= resolved.startIndex) {
      throw terminalNewlineEditError(resolved.targetMode);
    }

    return {
      startIndex: resolved.startIndex,
      endIndex: trimmedEndIndex,
      actualText: resolved.actualText.slice(0, -1),
      warnings: [DOC_TERMINAL_NEWLINE_WARNING],
    };
  }

  function rewriteTerminalNewlineMutationError(err: unknown): unknown {
    const message = err instanceof Error ? err.message : String(err);
    if (
      message.includes("deleteContentRange") &&
      message.includes("newline character at the end of the segment")
    ) {
      return new Error(
        'Google Docs rejected the edit because the requested range included the final paragraph newline of the current tab. ' +
          'Use `paragraph.displayText` or omit the trailing "\\n" in `target_text`, or adjust `end_index` to exclude it.',
      );
    }
    return err;
  }

  // Brand-new Docs can expose a non-text placeholder paragraph at index 0.
  function firstInsertableTextIndex(snapshot: CachedDocumentTabContent): number | undefined {
    for (const paragraph of snapshot.tab.paragraphs ?? []) {
      for (const element of paragraph.elements) {
        if (element.type === "textRun" && element.startIndex < element.endIndex) {
          return element.startIndex;
        }
      }
    }
    return undefined;
  }

  function isValidInsertionIndex(
    snapshot: CachedDocumentTabContent,
    index: number,
  ): boolean {
    return (snapshot.tab.paragraphs ?? []).some((paragraph) =>
      paragraph.elements.some(
        (element) =>
          element.type === "textRun" &&
          index >= element.startIndex &&
          index < element.endIndex,
      ),
    );
  }

  async function getStructuredDocumentTab(
    documentId: string,
    requestedTabId?: string,
  ): Promise<CachedDocumentTabContent> {
    const { tab } = await resolveDocumentTab(documentId, requestedTabId);
    const revisionId = lastSeenDocRevision.get(documentId);
    if (revisionId) {
      const cached = structuredDocCache.get(docCacheKey(documentId, revisionId, tab.tabId));
      if (cached) {
        return cached;
      }
    }

    const snapshot = await docsClient.getDocument(documentId, {
      includeContent: true,
      tabId: tab.tabId,
      maxChars: 250_000,
      maxParagraphs: 5_000,
    });
    rememberDocumentSnapshot(snapshot);

    if (snapshot.contentTruncated) {
      throw new Error(
        "Document content was truncated during internal anchor resolution. " +
          "Use gdrive_get_document_info with include_content=true and a narrower tab, or provide explicit indices.",
      );
    }

    const nextRevisionId = snapshot.revisionId;
    if (nextRevisionId) {
      const cached = structuredDocCache.get(docCacheKey(documentId, nextRevisionId, tab.tabId));
      if (cached) {
        return cached;
      }
    }

    const snapshotTab = snapshot.tabs.find((candidate) => candidate.tabId === tab.tabId);
    if (!snapshotTab?.paragraphs) {
      throw new Error(`Tab "${tab.tabId}" did not return structured content.`);
    }

    const search = buildSearchableTabText(snapshotTab);
    const built: CachedDocumentTabContent = {
      documentId,
      revisionId: snapshot.revisionId ?? lastSeenDocRevision.get(documentId) ?? "",
      tabId: tab.tabId,
      tab: snapshotTab,
      searchableText: search.searchableText,
      baseIndex: search.baseIndex,
    };

    if (built.revisionId) {
      structuredDocCache.set(docCacheKey(documentId, built.revisionId, tab.tabId), built);
    }

    return built;
  }

  async function resolveDocumentRange(options: {
    documentId: string;
    tabId?: string;
    startIndex?: number;
    endIndex?: number;
    targetText?: string;
    occurrence?: number;
    matchCase?: boolean;
    expectedText?: string;
  }): Promise<ResolvedDocumentRange> {
    const hasExplicitRange = options.startIndex !== undefined || options.endIndex !== undefined;
    const hasTargetText = options.targetText !== undefined;

    if (hasExplicitRange === hasTargetText) {
      throw new Error(
        "Provide exactly one target mode: either start_index/end_index or target_text.",
      );
    }

    const snapshot = await getStructuredDocumentTab(options.documentId, options.tabId);

    if (hasExplicitRange) {
      if (options.startIndex === undefined || options.endIndex === undefined) {
        throw new Error("Both start_index and end_index are required when using explicit ranges.");
      }
      assertRangeWithinSnapshot(snapshot, options.startIndex, options.endIndex);
      const actualText = extractRangeText(snapshot, options.startIndex, options.endIndex);
      if (options.expectedText !== undefined && actualText !== options.expectedText) {
        throw new Error(
          "Precondition failed: current document text does not match expected_text.\n" +
            `Actual text: ${JSON.stringify(actualText)}`,
        );
      }
      return {
        tabId: snapshot.tabId,
        startIndex: options.startIndex,
        endIndex: options.endIndex,
        actualText,
        targetMode: "explicit",
        snapshot,
      };
    }

    const targetText = options.targetText ?? "";
    const matches = findTextMatches(
      snapshot.searchableText,
      targetText,
      options.matchCase ?? true,
    );

    if (matches.length === 0) {
      throw new Error(
        `Could not find ${JSON.stringify(targetText)} in the current tab content. ` +
          "Inspect current content with gdrive_get_document_info include_content=true.",
      );
    }

    if (options.occurrence === undefined && matches.length > 1) {
      throw new Error(
        `Found ${matches.length} matches for ${JSON.stringify(targetText)}. ` +
          "Provide occurrence (1-based) or inspect current content with gdrive_get_document_info include_content=true.",
      );
    }

    const occurrence = options.occurrence ?? 1;
    if (occurrence < 1) {
      throw new Error("occurrence must be 1 or greater.");
    }
    if (occurrence > matches.length) {
      throw new Error(
        `occurrence ${occurrence} is out of range; only ${matches.length} match(es) were found for ${JSON.stringify(targetText)}.`,
      );
    }

    const match = matches[occurrence - 1];
    const startIndex = snapshot.baseIndex + match.startOffset;
    const endIndex = snapshot.baseIndex + match.endOffset;
    const actualText = extractRangeText(snapshot, startIndex, endIndex);

    if (options.expectedText !== undefined && actualText !== options.expectedText) {
      throw new Error(
        "Precondition failed: current document text does not match expected_text.\n" +
          `Actual text: ${JSON.stringify(actualText)}`,
      );
    }

    return {
      tabId: snapshot.tabId,
      startIndex,
      endIndex,
      actualText,
      targetMode: "anchor",
      snapshot,
    };
  }

  async function resolveInsertTarget(options: {
    documentId: string;
    tabId?: string;
    index?: number;
    position?: "start" | "end";
    beforeText?: string;
    afterText?: string;
    occurrence?: number;
    matchCase?: boolean;
  }): Promise<{ tabId: string; index?: number; atEnd: boolean }> {
    const modes = [
      options.index !== undefined,
      options.position !== undefined,
      options.beforeText !== undefined,
      options.afterText !== undefined,
    ].filter(Boolean).length;

    if (modes !== 1) {
      throw new Error(
        "Provide exactly one insertion target: index, position, before_text, or after_text.",
      );
    }

    if (options.index !== undefined) {
      const snapshot = await getStructuredDocumentTab(options.documentId, options.tabId);
      if (!isValidInsertionIndex(snapshot, options.index)) {
        throw new Error(
          `Insertion index ${options.index} is not inside an editable text paragraph. ` +
            "Use gdrive_get_document_info include_content=true to inspect valid indices, " +
            "or prefer position:'end' for a blank/newly created doc.",
        );
      }
      return { tabId: snapshot.tabId, index: options.index, atEnd: false };
    }

    if (options.position === "end") {
      const { tab } = await resolveDocumentTab(options.documentId, options.tabId);
      return { tabId: tab.tabId, atEnd: true };
    }

    if (options.position === "start") {
      const snapshot = await getStructuredDocumentTab(options.documentId, options.tabId);
      const firstIndex = firstInsertableTextIndex(snapshot);
      if (firstIndex === undefined) {
        return { tabId: snapshot.tabId, atEnd: true };
      }
      return {
        tabId: snapshot.tabId,
        index: firstIndex,
        atEnd: false,
      };
    }

    const anchorRange = await resolveDocumentRange({
      documentId: options.documentId,
      tabId: options.tabId,
      targetText: options.beforeText ?? options.afterText,
      occurrence: options.occurrence,
      matchCase: options.matchCase,
    });

    return {
      tabId: anchorRange.tabId,
      index: options.beforeText ? anchorRange.startIndex : anchorRange.endIndex,
      atEnd: false,
    };
  }

  function snapRangeToParagraphs(
    snapshot: CachedDocumentTabContent,
    startIndex: number,
    endIndex: number,
  ): { startIndex: number; endIndex: number } {
    const overlapping = (snapshot.tab.paragraphs ?? []).filter(
      (paragraph) => paragraph.endIndex > startIndex && paragraph.startIndex < endIndex,
    );
    if (overlapping.length === 0) {
      throw new Error(
        "The selected range does not overlap any paragraph boundaries in the current tab content.",
      );
    }
    return {
      startIndex: overlapping[0].startIndex,
      endIndex: overlapping[overlapping.length - 1].endIndex,
    };
  }

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
        const resolvedTargetId =
          meta.mimeType === "application/vnd.google-apps.shortcut"
            ? meta.shortcutTarget?.id
            : file_id;
        const resolvedTargetMimeType =
          meta.mimeType === "application/vnd.google-apps.shortcut"
            ? meta.shortcutTarget?.mimeType
            : meta.mimeType;

        if (resolvedTargetMimeType === "application/vnd.google-apps.spreadsheet" && resolvedTargetId) {
          accessedSpreadsheets.add(resolvedTargetId);
        }

        if (resolvedTargetMimeType === "application/vnd.google-apps.document" && resolvedTargetId) {
          rememberDocumentRead(resolvedTargetId);
          try {
            const revisionId = await docsClient.getRevisionId(resolvedTargetId);
            rememberDocumentRead(resolvedTargetId, revisionId);
          } catch {
            // Preserve Markdown read behavior even if revision caching is unavailable.
          }
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

  // ── Read-only Docs tool ────────────────────────────────────────────

  server.tool(
    "gdrive_get_document_info",
    "Get structured Google Docs metadata and optional tab-scoped content. " +
      "By default this is metadata-first so large documents stay compact. " +
      "Set include_content=true to retrieve paragraph-level content for one tab or all tabs. " +
      "Paragraphs include both raw `text` and anchor-friendly `displayText` (without a trailing paragraph newline). " +
      "Example: {\"document_id\":\"doc123\",\"include_content\":true,\"tab_id\":\"tab-1\",\"max_chars\":12000,\"max_paragraphs\":80}.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      include_content: z
        .boolean()
        .default(false)
        .describe("Include structured paragraph content in the response"),
      tab_id: z
        .string()
        .optional()
        .describe("Optional tab ID to scope returned content to a specific tab"),
      max_chars: z
        .number()
        .int()
        .min(0)
        .default(20_000)
        .describe("Maximum characters of structured content to return when include_content=true"),
      max_paragraphs: z
        .number()
        .int()
        .min(1)
        .default(200)
        .describe("Maximum paragraphs to return when include_content=true"),
    },
    SAFE,
    async ({ document_id, include_content, tab_id, max_chars, max_paragraphs }) => {
      try {
        const info = await docsClient.getDocument(document_id, {
          includeContent: include_content,
          tabId: tab_id,
          maxChars: max_chars,
          maxParagraphs: max_paragraphs,
        });
        rememberDocumentSnapshot(info);
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

  // ── Write tools — Docs content and formatting ──────────────────────

  const docConflictModeSchema = z
    .enum(["strict", "merge"])
    .default("strict")
    .describe(
      "strict uses requiredRevisionId and fails on concurrent edits; merge uses targetRevisionId for collaborative merges.",
    );
  const docNamedStyleSchema = z
    .enum([
      "NORMAL_TEXT",
      "TITLE",
      "SUBTITLE",
      "HEADING_1",
      "HEADING_2",
      "HEADING_3",
      "HEADING_4",
      "HEADING_5",
      "HEADING_6",
    ])
    .describe("Paragraph named style, e.g. HEADING_2");
  const docAlignmentSchema = z
    .enum(["START", "CENTER", "END", "JUSTIFIED"])
    .describe("Paragraph alignment");
  const docColorSchema = z
    .string()
    .regex(HEX_COLOR_PATTERN, "Expected a 6-digit hex color such as #3366FF.")
    .describe("Hex color such as #3366FF");
  const docListPresetSchema = z
    .enum(["BULLETED", "NUMBERED", "CHECKBOX", "REMOVE"])
    .describe("High-level list preset or REMOVE to clear bullets");

  server.tool(
    "gdrive_create_doc",
    "Create a blank Google Doc. Optionally move it into a specific folder after creation.",
    {
      title: z.string().describe("Document title"),
      folder_id: z
        .string()
        .default("root")
        .describe("Destination folder ID, or 'root' for the user's root Drive folder"),
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ title, folder_id }) => {
      try {
        const result = await docsClient.createDocument(title, folder_id);
        rememberDocumentRead(result.documentId, result.revisionId);
        return jsonResult(result);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_insert_doc_text",
    "Insert text into an existing Google Doc. For a newly created or blank doc, prefer position:'end' for the first write. " +
      "Prefer position:'start'|'end' or before_text/after_text anchors instead of raw indices when possible, and use raw indices only after inspecting gdrive_get_document_info include_content=true. " +
      "Examples: {\"document_id\":\"doc123\",\"position\":\"end\",\"text\":\"Hello\"} or {\"document_id\":\"doc123\",\"before_text\":\"TODO\",\"occurrence\":1,\"text\":\"- \"}.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().optional().describe("Optional tab ID; defaults to the first tab"),
      text: z.string().describe("Text to insert"),
      index: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Power-user fallback: explicit UTF-16 insertion index inside an existing text paragraph"),
      position: z
        .enum(["start", "end"])
        .optional()
        .describe("LLM-friendly insertion target without index arithmetic; for a brand-new or blank doc, 'end' is the safest first write"),
      before_text: z
        .string()
        .optional()
        .describe("Insert immediately before this exact text match"),
      after_text: z
        .string()
        .optional()
        .describe("Insert immediately after this exact text match"),
      occurrence: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe("1-based match occurrence when before_text/after_text is repeated"),
      match_case: z
        .boolean()
        .default(true)
        .describe("Whether before_text/after_text matching is case-sensitive"),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      document_id,
      tab_id,
      text,
      index,
      position,
      before_text,
      after_text,
      occurrence,
      match_case,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      try {
        const target = await resolveInsertTarget({
          documentId: document_id,
          tabId: tab_id,
          index,
          position,
          beforeText: before_text,
          afterText: after_text,
          occurrence,
          matchCase: match_case,
        });
        const revisionId = await ensureDocumentRevision(document_id);
        const result = await docsClient.insertText({
          documentId: document_id,
          text,
          tabId: target.tabId,
          index: target.index,
          atEnd: target.atEnd,
          revisionId,
          conflictMode: conflict_mode as DocsConflictMode,
        });
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: target.tabId,
          revisionId: result.revisionId,
          insertedText: text,
          position: target.atEnd ? "end" : undefined,
          index: target.index,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_replace_doc_text",
    "Replace a targeted text range in an existing Google Doc. Use target_text for anchored replacements or explicit start_index/end_index as a fallback. " +
      "Example: {\"document_id\":\"doc123\",\"target_text\":\"Draft\",\"replacement_text\":\"Final\",\"occurrence\":1,\"expected_text\":\"Draft\"}.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().optional().describe("Optional tab ID; defaults to the first tab"),
      replacement_text: z.string().describe("Replacement text; use an empty string only if you intentionally want a delete-style replacement"),
      start_index: z.number().int().min(0).optional().describe("Explicit UTF-16 start index"),
      end_index: z.number().int().min(0).optional().describe("Explicit UTF-16 end index (exclusive)"),
      target_text: z
        .string()
        .optional()
        .describe(
          "Anchor to replace this exact text match instead of passing indices. Prefer `paragraph.displayText` from gdrive_get_document_info; raw paragraph text may include a trailing newline.",
        ),
      occurrence: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe("1-based match occurrence when target_text appears multiple times"),
      match_case: z
        .boolean()
        .default(true)
        .describe("Whether target_text matching is case-sensitive"),
      expected_text: z
        .string()
        .optional()
        .describe("Optional optimistic safety check for the current text in the resolved range"),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      document_id,
      tab_id,
      replacement_text,
      start_index,
      end_index,
      target_text,
      occurrence,
      match_case,
      expected_text,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      try {
        const resolved = await resolveDocumentRange({
          documentId: document_id,
          tabId: tab_id,
          startIndex: start_index,
          endIndex: end_index,
          targetText: target_text,
          occurrence,
          matchCase: match_case,
          expectedText: expected_text,
        });
        const normalized = normalizeRangeForTextMutation(resolved);
        const revisionId = await ensureDocumentRevision(document_id);
        const result = await docsClient.replaceText({
          documentId: document_id,
          text: replacement_text,
          tabId: resolved.tabId,
          startIndex: normalized.startIndex,
          endIndex: normalized.endIndex,
          revisionId,
          conflictMode: conflict_mode as DocsConflictMode,
        });
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: resolved.tabId,
          revisionId: result.revisionId,
          previousText: normalized.actualText,
          replacementText: replacement_text,
          replacedRange: {
            startIndex: normalized.startIndex,
            endIndex: normalized.endIndex,
          },
          warnings: normalized.warnings,
        });
      } catch (err) {
        return errorResult(rewriteTerminalNewlineMutationError(err));
      }
    },
  );

  server.tool(
    "gdrive_replace_all_doc_text",
    "Replace every exact text match in a Google Doc. For safety this defaults to the first tab unless you pass tab_id or set all_tabs=true explicitly. " +
      "Set match_case=false when casing is uncertain. Example: {\"document_id\":\"doc123\",\"old_text\":\"Acme\",\"new_text\":\"Wagner Labs\",\"all_tabs\":true}.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      old_text: z.string().describe("Exact text to find"),
      new_text: z.string().describe("Replacement text"),
      tab_id: z
        .string()
        .optional()
        .describe("Optional tab ID to scope replacement to a specific tab"),
      all_tabs: z
        .boolean()
        .default(false)
        .describe("Explicit opt-in to replace across every tab in the document"),
      match_case: z
        .boolean()
        .default(true)
        .describe("Whether old_text matching is case-sensitive"),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ document_id, old_text, new_text, tab_id, all_tabs, match_case, conflict_mode }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      if (all_tabs && tab_id) {
        return errorResult(
          new Error("Pass either tab_id or all_tabs=true, not both."),
        );
      }

      try {
        const revisionId = await ensureDocumentRevision(document_id);
        const scopedTabId = all_tabs
          ? undefined
          : (await resolveDocumentTab(document_id, tab_id)).tab.tabId;
        const result = await docsClient.replaceAllText({
          documentId: document_id,
          searchText: old_text,
          replaceText: new_text,
          matchCase: match_case,
          tabId: scopedTabId,
          allTabs: all_tabs,
          revisionId,
          conflictMode: conflict_mode as DocsConflictMode,
        });
        rememberDocumentRead(document_id, result.revisionId);
        const occurrencesChanged = (result.replies?.[0] as { replaceAllText?: { occurrencesChanged?: number } } | undefined)
          ?.replaceAllText?.occurrencesChanged;
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          revisionId: result.revisionId,
          tabId: scopedTabId,
          allTabs: all_tabs,
          oldText: old_text,
          newText: new_text,
          matchCase: match_case,
          occurrencesChanged,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_delete_doc_text",
    "Delete a targeted text range from a Google Doc using explicit indices or target_text anchors.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().optional().describe("Optional tab ID; defaults to the first tab"),
      start_index: z.number().int().min(0).optional().describe("Explicit UTF-16 start index"),
      end_index: z.number().int().min(0).optional().describe("Explicit UTF-16 end index (exclusive)"),
      target_text: z
        .string()
        .optional()
        .describe(
          "Anchor to delete this exact text match instead of passing indices. Prefer `paragraph.displayText` from gdrive_get_document_info; raw paragraph text may include a trailing newline.",
        ),
      occurrence: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe("1-based match occurrence when target_text appears multiple times"),
      match_case: z
        .boolean()
        .default(true)
        .describe("Whether target_text matching is case-sensitive"),
      expected_text: z
        .string()
        .optional()
        .describe("Optional optimistic safety check for the current text in the resolved range"),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      document_id,
      tab_id,
      start_index,
      end_index,
      target_text,
      occurrence,
      match_case,
      expected_text,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      try {
        const resolved = await resolveDocumentRange({
          documentId: document_id,
          tabId: tab_id,
          startIndex: start_index,
          endIndex: end_index,
          targetText: target_text,
          occurrence,
          matchCase: match_case,
          expectedText: expected_text,
        });
        const normalized = normalizeRangeForTextMutation(resolved);
        const revisionId = await ensureDocumentRevision(document_id);
        const result = await docsClient.deleteText({
          documentId: document_id,
          tabId: resolved.tabId,
          startIndex: normalized.startIndex,
          endIndex: normalized.endIndex,
          revisionId,
          conflictMode: conflict_mode as DocsConflictMode,
        });
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: resolved.tabId,
          revisionId: result.revisionId,
          deletedText: normalized.actualText,
          deletedRange: {
            startIndex: normalized.startIndex,
            endIndex: normalized.endIndex,
          },
          warnings: normalized.warnings,
        });
      } catch (err) {
        return errorResult(rewriteTerminalNewlineMutationError(err));
      }
    },
  );

  server.tool(
    "gdrive_update_doc_text_style",
    "Apply character-level formatting in a Google Doc using explicit ranges or target_text anchors. " +
      "Supports bold, italic, underline, strikethrough, font family, font size, colors, and links.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().optional().describe("Optional tab ID; defaults to the first tab"),
      start_index: z.number().int().min(0).optional().describe("Explicit UTF-16 start index"),
      end_index: z.number().int().min(0).optional().describe("Explicit UTF-16 end index (exclusive)"),
      target_text: z
        .string()
        .optional()
        .describe("Anchor to style this exact text match instead of passing indices"),
      occurrence: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe("1-based match occurrence when target_text appears multiple times"),
      match_case: z
        .boolean()
        .default(true)
        .describe("Whether target_text matching is case-sensitive"),
      bold: z.boolean().optional().describe("Bold text"),
      italic: z.boolean().optional().describe("Italic text"),
      underline: z.boolean().optional().describe("Underline text"),
      strikethrough: z.boolean().optional().describe("Strikethrough text"),
      font_family: z.string().optional().describe("Font family, e.g. 'Arial'"),
      font_size: z.number().positive().optional().describe("Font size in points"),
      foreground_color: docColorSchema.optional(),
      background_color: docColorSchema.optional(),
      link_url: z.string().url().optional().describe("Optional hyperlink URL"),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      document_id,
      tab_id,
      start_index,
      end_index,
      target_text,
      occurrence,
      match_case,
      bold,
      italic,
      underline,
      strikethrough,
      font_family,
      font_size,
      foreground_color,
      background_color,
      link_url,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      const hasTextStyleUpdate =
        bold !== undefined ||
        italic !== undefined ||
        underline !== undefined ||
        strikethrough !== undefined ||
        font_family !== undefined ||
        font_size !== undefined ||
        foreground_color !== undefined ||
        background_color !== undefined ||
        link_url !== undefined;
      if (!hasTextStyleUpdate) {
        return errorResult(
          new Error("At least one text style parameter must be provided."),
        );
      }

      try {
        const resolved = await resolveDocumentRange({
          documentId: document_id,
          tabId: tab_id,
          startIndex: start_index,
          endIndex: end_index,
          targetText: target_text,
          occurrence,
          matchCase: match_case,
        });
        const revisionId = await ensureDocumentRevision(document_id);
        const result = await docsClient.updateTextStyle({
          documentId: document_id,
          tabId: resolved.tabId,
          startIndex: resolved.startIndex,
          endIndex: resolved.endIndex,
          bold,
          italic,
          underline,
          strikethrough,
          fontFamily: font_family,
          fontSize: font_size,
          foregroundColor: foreground_color,
          backgroundColor: background_color,
          linkUrl: link_url,
          revisionId,
          conflictMode: conflict_mode as DocsConflictMode,
        });
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: resolved.tabId,
          revisionId: result.revisionId,
          styledRange: {
            startIndex: resolved.startIndex,
            endIndex: resolved.endIndex,
          },
          targetText: resolved.actualText,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_update_doc_paragraph_style",
    "Apply paragraph-level formatting in a Google Doc. Supports headings and alignment, and snaps the resolved range to full paragraph boundaries server-side.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().optional().describe("Optional tab ID; defaults to the first tab"),
      start_index: z.number().int().min(0).optional().describe("Explicit UTF-16 start index"),
      end_index: z.number().int().min(0).optional().describe("Explicit UTF-16 end index (exclusive)"),
      target_text: z
        .string()
        .optional()
        .describe("Anchor to paragraphs overlapping this exact text match"),
      occurrence: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe("1-based match occurrence when target_text appears multiple times"),
      match_case: z
        .boolean()
        .default(true)
        .describe("Whether target_text matching is case-sensitive"),
      named_style_type: docNamedStyleSchema.optional(),
      alignment: docAlignmentSchema.optional(),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      document_id,
      tab_id,
      start_index,
      end_index,
      target_text,
      occurrence,
      match_case,
      named_style_type,
      alignment,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      if (named_style_type === undefined && alignment === undefined) {
        return errorResult(
          new Error("At least one paragraph style parameter must be provided."),
        );
      }

      try {
        const resolved = await resolveDocumentRange({
          documentId: document_id,
          tabId: tab_id,
          startIndex: start_index,
          endIndex: end_index,
          targetText: target_text,
          occurrence,
          matchCase: match_case,
        });
        const paragraphRange = snapRangeToParagraphs(
          resolved.snapshot,
          resolved.startIndex,
          resolved.endIndex,
        );
        const revisionId = await ensureDocumentRevision(document_id);
        const result = await docsClient.updateParagraphStyle({
          documentId: document_id,
          tabId: resolved.tabId,
          startIndex: paragraphRange.startIndex,
          endIndex: paragraphRange.endIndex,
          namedStyleType: named_style_type as DocNamedStyleType | undefined,
          alignment: alignment as DocParagraphAlignment | undefined,
          revisionId,
          conflictMode: conflict_mode as DocsConflictMode,
        });
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: resolved.tabId,
          revisionId: result.revisionId,
          paragraphRange,
          namedStyleType: named_style_type,
          alignment,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_update_doc_list",
    "Create, change, or remove list formatting in a Google Doc. The resolved range is snapped to full paragraph boundaries server-side.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().optional().describe("Optional tab ID; defaults to the first tab"),
      start_index: z.number().int().min(0).optional().describe("Explicit UTF-16 start index"),
      end_index: z.number().int().min(0).optional().describe("Explicit UTF-16 end index (exclusive)"),
      target_text: z
        .string()
        .optional()
        .describe("Anchor to paragraphs overlapping this exact text match"),
      occurrence: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe("1-based match occurrence when target_text appears multiple times"),
      match_case: z
        .boolean()
        .default(true)
        .describe("Whether target_text matching is case-sensitive"),
      preset: docListPresetSchema,
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({
      document_id,
      tab_id,
      start_index,
      end_index,
      target_text,
      occurrence,
      match_case,
      preset,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      try {
        const resolved = await resolveDocumentRange({
          documentId: document_id,
          tabId: tab_id,
          startIndex: start_index,
          endIndex: end_index,
          targetText: target_text,
          occurrence,
          matchCase: match_case,
        });
        const paragraphRange = snapRangeToParagraphs(
          resolved.snapshot,
          resolved.startIndex,
          resolved.endIndex,
        );
        const revisionId = await ensureDocumentRevision(document_id);
        const result = await docsClient.updateList({
          documentId: document_id,
          tabId: resolved.tabId,
          startIndex: paragraphRange.startIndex,
          endIndex: paragraphRange.endIndex,
          preset: preset as DocListPreset,
          revisionId,
          conflictMode: conflict_mode as DocsConflictMode,
        });
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: resolved.tabId,
          revisionId: result.revisionId,
          paragraphRange,
          preset,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  // ── Write tools — Docs file operations ─────────────────────────────

  server.tool(
    "gdrive_rename_doc",
    "Rename an existing Google Doc file. This is a Drive file operation, not a Docs content edit.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      new_title: z.string().describe("New document title"),
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ document_id, new_title }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      try {
        const result = await docsClient.renameDocument(document_id, new_title);
        return jsonResult(result);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_duplicate_doc",
    "Duplicate an existing Google Doc. When folder_id is omitted, the copy keeps the source document's parent folder placement.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      new_title: z.string().optional().describe("Optional title for the copied document"),
      folder_id: z
        .string()
        .optional()
        .describe("Optional destination folder ID for the copy"),
    },
    {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ document_id, new_title, folder_id }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      try {
        const result = await docsClient.duplicateDocument(document_id, new_title, folder_id);
        rememberDocumentRead(result.documentId, result.revisionId);
        return jsonResult(result);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  return server;
}
