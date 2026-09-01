import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ToolAnnotations } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { docs_v1 } from "googleapis";
import { DriveClient, DriveAPIError } from "./client.js";
import { SheetsClient, FormatOptions } from "./sheets-client.js";
import {
  DocListPreset,
  DocNamedStyleType,
  DocParagraphAlignment,
  DocsClient,
  DocsConflictMode,
  NormalizedDocBlock,
  NormalizedDocElement,
  NormalizedDocParagraph,
  NormalizedDocTableBlock,
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
const DOC_RESPONSE_MAX_BYTES = 48_000;

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

function docBulletPreset(preset: Exclude<DocListPreset, "REMOVE">): string {
  if (preset === "CHECKBOX") return "BULLET_CHECKBOX";
  if (preset === "NUMBERED") return "NUMBERED_DECIMAL_ALPHA_ROMAN";
  return "BULLET_DISC_CIRCLE_SQUARE";
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

interface DocumentPageToken {
  documentId: string;
  revisionId: string;
  tabId: string;
  offset: number;
}

function encodeDocumentPageToken(token: DocumentPageToken): string {
  return Buffer.from(JSON.stringify(token), "utf8").toString("base64url");
}

function decodeDocumentPageToken(value: string): DocumentPageToken {
  try {
    const decoded = JSON.parse(Buffer.from(value, "base64url").toString("utf8"));
    if (
      typeof decoded.documentId !== "string" ||
      typeof decoded.revisionId !== "string" ||
      typeof decoded.tabId !== "string" ||
      !Number.isInteger(decoded.offset) ||
      decoded.offset < 0
    ) {
      throw new Error("invalid fields");
    }
    return decoded as DocumentPageToken;
  } catch {
    throw new Error("Invalid document page_token.");
  }
}

function splitLargeDocumentBlocks(blocks: NormalizedDocBlock[]): NormalizedDocBlock[] {
  const fragments: NormalizedDocBlock[] = [];
  for (const block of blocks) {
    if (block.type === "table" && Buffer.byteLength(JSON.stringify(block), "utf8") > 24_000) {
      for (const row of block.tableRows) {
        const rowFragment: NormalizedDocTableBlock = {
          ...block,
          rows: block.rows,
          tableRows: [row],
          fragment: { rowIndex: row.rowIndex },
        };
        if (Buffer.byteLength(JSON.stringify(rowFragment), "utf8") <= 24_000) {
          fragments.push(rowFragment);
          continue;
        }

        for (const cell of row.cells) {
          const cellFragment: NormalizedDocTableBlock = {
            ...block,
            tableRows: [{ ...row, cells: [cell] }],
            fragment: { rowIndex: row.rowIndex, columnIndex: cell.columnIndex },
          };
          if (Buffer.byteLength(JSON.stringify(cellFragment), "utf8") <= 24_000) {
            fragments.push(cellFragment);
            continue;
          }

          const chunkSize = 6_000;
          for (let offset = 0; offset < cell.text.length; offset += chunkSize) {
            const text = cell.text.slice(offset, offset + chunkSize);
            fragments.push({
              ...block,
              tableRows: [{
                ...row,
                cells: [{
                  ...cell,
                  startIndex: cell.startIndex + 1 + offset,
                  endIndex: cell.startIndex + 1 + offset + text.length,
                  text,
                  blocks: [],
                  fragment: {
                    textOffset: offset,
                    totalTextLength: cell.text.length,
                  },
                }],
              }],
              fragment: { rowIndex: row.rowIndex, columnIndex: cell.columnIndex },
            });
          }
        }
      }
      continue;
    }
    if (block.type === "paragraph" && Buffer.byteLength(JSON.stringify(block), "utf8") > 24_000) {
      const chunkSize = 8_000;
      for (let offset = 0; offset < block.text.length; offset += chunkSize) {
        const text = block.text.slice(offset, offset + chunkSize);
        fragments.push({
          ...block,
          startIndex: block.startIndex + offset,
          endIndex: block.startIndex + offset + text.length,
          text,
          displayText: text.endsWith("\n") ? text.slice(0, -1) : text,
          elements: [{
            type: "textRun",
            startIndex: block.startIndex + offset,
            endIndex: block.startIndex + offset + text.length,
            text,
            textStyle: null,
          }],
        } as NormalizedDocParagraph);
      }
      continue;
    }
    fragments.push(block);
  }
  return fragments;
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
    version: "0.2.0",
  });

  const accessedSpreadsheets = new Set<string>();
  const accessedDocs = new Set<string>();
  // Only explicit tool reads and successful agent-initiated writes may advance
  // this map. Internal metadata/content fetches must never authorize a newer
  // collaborator revision.
  const observedDocRevision = new Map<string, string>();
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
      const previousRevisionId = observedDocRevision.get(documentId);
      observedDocRevision.set(documentId, revisionId);
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
        "Use gdrive_read_file, gdrive_get_document_info, or gdrive_get_document_content first.",
    );
  }

  function staleDocumentError(
    observedRevisionId: string,
    currentRevisionId: string,
  ): Error {
    return new Error(
      "STALE_DOCUMENT: The document changed after the agent read it. " +
        `Observed revision ${JSON.stringify(observedRevisionId)}; current revision ${JSON.stringify(currentRevisionId)}. ` +
        "Read the document again before retrying, or explicitly use conflict_mode:'merge'.",
    );
  }

  async function fetchDocumentMetadata(documentId: string): Promise<NormalizedDocument> {
    const metadata = await docsClient.getDocument(documentId, {
      includeContent: false,
    });
    return metadata;
  }

  async function ensureDocumentRevision(
    documentId: string,
    conflictMode: DocsConflictMode = "strict",
  ): Promise<string> {
    const cached = observedDocRevision.get(documentId);
    if (cached) {
      if (conflictMode === "strict") {
        const current = await docsClient.getRevisionId(documentId);
        if (current && current !== cached) {
          throw staleDocumentError(cached, current);
        }
      }
      return cached;
    }
    throw new Error(
      "Could not determine the revision that was explicitly read. " +
        "Read the document again before editing it.",
    );
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
    const revisionId = observedDocRevision.get(documentId);
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
    cacheStructuredDocument(snapshot);

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
      revisionId: snapshot.revisionId ?? observedDocRevision.get(documentId) ?? "",
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

  function collectTables(blocks: NormalizedDocBlock[] | undefined): NormalizedDocTableBlock[] {
    const tables: NormalizedDocTableBlock[] = [];
    for (const block of blocks ?? []) {
      if (block.type !== "table") continue;
      tables.push(block);
      for (const row of block.tableRows) {
        for (const cell of row.cells) {
          tables.push(...collectTables(cell.blocks));
        }
      }
    }
    return tables;
  }

  function resolveTableByPath(
    snapshot: CachedDocumentTabContent,
    tablePath: number[],
  ): NormalizedDocTableBlock {
    const table = collectTables(snapshot.tab.blocks).find(
      (candidate) =>
        candidate.tablePath.length === tablePath.length &&
        candidate.tablePath.every((part, index) => part === tablePath[index]),
    );
    if (!table) {
      throw new Error(
        `Table path [${tablePath.join(", ")}] was not found in tab "${snapshot.tab.title}". Reread the tab content to inspect current table paths.`,
      );
    }
    return table;
  }

  function resolveTableCell(
    table: NormalizedDocTableBlock,
    rowIndex: number,
    columnIndex: number,
  ) {
    const row = table.tableRows.find((candidate) => candidate.rowIndex === rowIndex);
    const cell = row?.cells.find((candidate) => candidate.columnIndex === columnIndex);
    if (!cell) {
      throw new Error(
        `Cell (${rowIndex}, ${columnIndex}) does not exist or is covered by a merged cell in table [${table.tablePath.join(", ")}].`,
      );
    }
    return cell;
  }

  async function refreshTable(
    documentId: string,
    tabId: string,
    tablePath: number[],
  ): Promise<NormalizedDocTableBlock> {
    const refreshed = await getStructuredDocumentTab(documentId, tabId);
    return resolveTableByPath(refreshed, tablePath);
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
        const resolvedTargetId =
          meta.mimeType === "application/vnd.google-apps.shortcut"
            ? meta.shortcutTarget?.id
            : file_id;
        const resolvedTargetMimeType =
          meta.mimeType === "application/vnd.google-apps.shortcut"
            ? meta.shortcutTarget?.mimeType
            : meta.mimeType;
        const docRevisionBefore =
          resolvedTargetMimeType === "application/vnd.google-apps.document" && resolvedTargetId
            ? await docsClient.getRevisionId(resolvedTargetId).catch(() => undefined)
            : undefined;
        const result = await driveClient.readFile(file_id, max_chars);

        if (resolvedTargetMimeType === "application/vnd.google-apps.spreadsheet" && resolvedTargetId) {
          accessedSpreadsheets.add(resolvedTargetId);
        }

        if (resolvedTargetMimeType === "application/vnd.google-apps.document" && resolvedTargetId) {
          const revisionAfter = await docsClient.getRevisionId(resolvedTargetId).catch(() => undefined);
          if (
            docRevisionBefore &&
            revisionAfter &&
            docRevisionBefore !== revisionAfter
          ) {
            throw new Error(
              "The document changed while its Markdown export was being read. Read it again before editing.",
            );
          }
          rememberDocumentRead(resolvedTargetId, revisionAfter ?? docRevisionBefore);
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

  server.tool(
    "gdrive_get_sheet_values",
    "Read a bounded A1 range from a spreadsheet. Use the returned values as expected_current_values for safe targeted updates.",
    {
      spreadsheet_id: z.string().describe("Google Sheets spreadsheet ID"),
      range: z.string().describe("Bounded A1 notation with a sheet name, e.g. 'Sheet1!A1:C20'"),
    },
    SAFE,
    async ({ spreadsheet_id, range }) => {
      try {
        const parsed = SheetsClient.parseA1Range(range);
        const cells =
          (parsed.endRow - parsed.startRow + 1) *
          (parsed.endCol - parsed.startCol + 1);
        if (cells > 10_000) {
          throw new Error(
            `Range contains ${cells} cells; read at most 10000 cells per call.`,
          );
        }
        const values = await sheetsClient.getValues(spreadsheet_id, range);
        accessedSpreadsheets.add(spreadsheet_id);
        return jsonResult({
          spreadsheetId: spreadsheet_id,
          spreadsheetUrl: spreadsheetUrl(spreadsheet_id),
          range,
          values: padValues(
            values,
            parsed.endRow - parsed.startRow + 1,
            parsed.endCol - parsed.startCol + 1,
          ),
        });
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
          maxChars: Math.min(max_chars, 20_000),
          maxParagraphs: Math.min(max_paragraphs, 200),
        });
        rememberDocumentSnapshot(info);
        return jsonResult(info);
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_get_document_content",
    "Read one Google Docs tab as bounded, paginated structured blocks. Returns paragraphs and native tables with a revision-bound nextPageToken.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().describe("Tab ID returned by gdrive_get_document_info"),
      page_token: z.string().optional().describe("Revision-bound continuation token from the previous page"),
      max_blocks: z.number().int().min(1).max(200).default(50),
    },
    SAFE,
    async ({ document_id, tab_id, page_token, max_blocks }) => {
      try {
        const decoded = page_token ? decodeDocumentPageToken(page_token) : undefined;
        if (
          decoded &&
          (decoded.documentId !== document_id || decoded.tabId !== tab_id)
        ) {
          throw new Error("page_token belongs to a different document or tab.");
        }

        const info = await docsClient.getDocument(document_id, {
          includeContent: true,
          tabId: tab_id,
          maxChars: 50_000_000,
          maxParagraphs: 1_000_000,
        });
        if (!info.revisionId) {
          throw new Error("Could not determine the document revision.");
        }
        if (decoded && decoded.revisionId !== info.revisionId) {
          throw new Error(
            "STALE_PAGE_TOKEN: The document changed between content pages. Restart the read from the first page.",
          );
        }

        const tab = info.tabs.find((candidate) => candidate.tabId === tab_id);
        if (!tab) {
          throw new Error(`Tab "${tab_id}" not found in document.`);
        }
        const allBlocks = splitLargeDocumentBlocks(tab.blocks ?? []);
        const offset = decoded?.offset ?? 0;
        const blocks: NormalizedDocBlock[] = [];

        for (let index = offset; index < allBlocks.length && blocks.length < max_blocks; index++) {
          const candidate = [...blocks, allBlocks[index]];
          const probe = {
            documentId: info.documentId,
            title: info.title,
            documentUrl: info.documentUrl,
            revisionId: info.revisionId,
            tab: {
              tabId: tab.tabId,
              title: tab.title,
              blocks: candidate,
            },
            nextPageToken: "x".repeat(256),
          };
          if (
            blocks.length > 0 &&
            Buffer.byteLength(JSON.stringify(probe, null, 2), "utf8") > DOC_RESPONSE_MAX_BYTES
          ) {
            break;
          }
          blocks.push(allBlocks[index]);
        }

        if (blocks.length === 0 && offset < allBlocks.length) {
          throw new Error(
            "A single document block exceeds the safe MCP response limit. Narrow the content or split the oversized table cell.",
          );
        }

        const nextOffset = offset + blocks.length;
        const nextPageToken = nextOffset < allBlocks.length
          ? encodeDocumentPageToken({
              documentId: info.documentId,
              revisionId: info.revisionId,
              tabId: tab.tabId,
              offset: nextOffset,
            })
          : undefined;

        rememberDocumentSnapshot(info);
        return jsonResult({
          documentId: info.documentId,
          title: info.title,
          documentUrl: info.documentUrl,
          revisionId: info.revisionId,
          tab: {
            tabId: tab.tabId,
            title: tab.title,
            blocks,
          },
          nextPageToken,
        });
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
              "Use gdrive_read_file, gdrive_get_spreadsheet_info, or gdrive_get_sheet_values first.",
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
              return errorResult(
                new Error(
                  `STALE_SHEET_VALUES for ${range}: ${mismatch}\n` +
                    `Expected values: ${JSON.stringify(expected_current_values)}`,
                ),
              );
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
  const docTextSegmentSchema = z.object({
    text: z.string(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    strikethrough: z.boolean().optional(),
    font_family: z.string().optional(),
    font_size: z.number().positive().optional(),
    foreground_color: docColorSchema.optional(),
    background_color: docColorSchema.optional(),
    link_url: z.string().url().optional(),
  });
  const docContentBlockSchema = z.object({
    segments: z.array(docTextSegmentSchema).min(1),
    named_style_type: docNamedStyleSchema.optional(),
    alignment: docAlignmentSchema.optional(),
    indent_start_points: z.number().nonnegative().optional(),
    indent_end_points: z.number().nonnegative().optional(),
    indent_first_line_points: z.number().optional(),
    space_above_points: z.number().nonnegative().optional(),
    space_below_points: z.number().nonnegative().optional(),
    line_spacing: z.number().positive().optional(),
    list: z.object({
      preset: z.enum(["BULLETED", "NUMBERED", "CHECKBOX"]),
      nesting_level: z.number().int().min(0).max(8).default(0),
    }).optional(),
  });
  const docBatchTargetFields = {
    target_text: z.string(),
    occurrence: z.number().int().min(1).optional(),
    match_case: z.boolean().default(true),
    expected_text: z.string().optional(),
  };
  const docBatchOperationSchema = z.discriminatedUnion("type", [
    z.object({
      type: z.literal("replace"),
      ...docBatchTargetFields,
      replacement_text: z.string(),
    }),
    z.object({
      type: z.literal("delete"),
      ...docBatchTargetFields,
    }),
    z.object({
      type: z.literal("text_style"),
      ...docBatchTargetFields,
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikethrough: z.boolean().optional(),
      foreground_color: docColorSchema.optional(),
      background_color: docColorSchema.optional(),
      link_url: z.string().url().optional(),
    }),
    z.object({
      type: z.literal("paragraph_style"),
      ...docBatchTargetFields,
      named_style_type: docNamedStyleSchema.optional(),
      alignment: docAlignmentSchema.optional(),
    }),
    z.object({
      type: z.literal("list"),
      ...docBatchTargetFields,
      preset: docListPresetSchema,
    }),
  ]);
  const docTablePathSchema = z.array(z.number().int().min(1)).min(1).describe(
    "One-based table path returned by gdrive_get_document_content, e.g. [2] or [2,1]",
  );
  const docTableStructureOperationSchema = z.discriminatedUnion("type", [
    z.object({ type: z.literal("insert_row"), row_index: z.number().int().min(0), position: z.enum(["before", "after"]) }),
    z.object({ type: z.literal("delete_row"), row_index: z.number().int().min(0) }),
    z.object({ type: z.literal("insert_column"), column_index: z.number().int().min(0), position: z.enum(["before", "after"]) }),
    z.object({ type: z.literal("delete_column"), column_index: z.number().int().min(0) }),
    z.object({
      type: z.literal("merge"),
      start_row: z.number().int().min(0),
      start_column: z.number().int().min(0),
      row_span: z.number().int().min(1),
      column_span: z.number().int().min(1),
    }),
    z.object({
      type: z.literal("unmerge"),
      start_row: z.number().int().min(0),
      start_column: z.number().int().min(0),
      row_span: z.number().int().min(1),
      column_span: z.number().int().min(1),
    }),
    z.object({ type: z.literal("delete_table") }),
  ]);

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
    "gdrive_insert_doc_content",
    "Insert deterministic rich paragraph blocks in one revision-controlled operation. Supports styled text segments, headings, spacing, indentation, and nested native lists.",
    {
      document_id: z.string().describe("Google Docs document ID"),
      tab_id: z.string().optional().describe("Optional tab ID; defaults to the first tab"),
      blocks: z.array(docContentBlockSchema).min(1).max(200),
      index: z.number().int().min(0).optional(),
      position: z.enum(["start", "end"]).optional(),
      before_text: z.string().optional(),
      after_text: z.string().optional(),
      occurrence: z.number().int().min(1).optional(),
      match_case: z.boolean().default(true),
      inherit_neighbor_style: z.boolean().default(false),
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
      blocks,
      index,
      position,
      before_text,
      after_text,
      occurrence,
      match_case,
      inherit_neighbor_style,
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
        const snapshot = await getStructuredDocumentTab(document_id, target.tabId);
        const insertionIndex = target.atEnd
          ? segmentTerminalNewlineIndex(snapshot) ?? firstInsertableTextIndex(snapshot) ?? 1
          : target.index!;

        const layouts: Array<{
          startIndex: number;
          endIndex: number;
          textStartIndex: number;
          segmentRanges: Array<{ startIndex: number; endIndex: number; segment: typeof blocks[number]["segments"][number] }>;
          block: typeof blocks[number];
        }> = [];
        let insertedText = "";
        let cursor = insertionIndex;
        for (const block of blocks) {
          const tabs = block.list ? "\t".repeat(block.list.nesting_level) : "";
          const blockStart = cursor;
          insertedText += tabs;
          cursor += tabs.length;
          const textStartIndex = cursor;
          const segmentRanges = [];
          for (const segment of block.segments) {
            const startIndex = cursor;
            insertedText += segment.text;
            cursor += segment.text.length;
            segmentRanges.push({ startIndex, endIndex: cursor, segment });
          }
          if (!insertedText.endsWith("\n")) {
            insertedText += "\n";
            cursor += 1;
          }
          layouts.push({
            startIndex: blockStart,
            endIndex: cursor,
            textStartIndex,
            segmentRanges,
            block,
          });
        }

        const requests: docs_v1.Schema$Request[] = [{
          insertText: {
            location: { tabId: target.tabId, index: insertionIndex },
            text: insertedText,
          },
        }];

        if (!inherit_neighbor_style) {
          requests.push({
            updateTextStyle: {
              range: {
                tabId: target.tabId,
                startIndex: insertionIndex,
                endIndex: insertionIndex + insertedText.length,
              },
              textStyle: {},
              fields: "bold,italic,underline,strikethrough,weightedFontFamily,fontSize,foregroundColor,backgroundColor,link",
            },
          });
          for (const layout of layouts.filter((item) => !item.block.list)) {
            requests.push({
              deleteParagraphBullets: {
                range: {
                  tabId: target.tabId,
                  startIndex: layout.startIndex,
                  endIndex: layout.endIndex,
                },
              },
            });
          }
        }

        for (const layout of layouts) {
          const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
          const fields: string[] = [];
          if (layout.block.named_style_type) {
            paragraphStyle.namedStyleType = layout.block.named_style_type;
            fields.push("namedStyleType");
          } else if (!inherit_neighbor_style) {
            paragraphStyle.namedStyleType = "NORMAL_TEXT";
            fields.push("namedStyleType");
          }
          if (layout.block.alignment) {
            paragraphStyle.alignment = layout.block.alignment;
            fields.push("alignment");
          }
          const dimensions: Array<[number | undefined, keyof docs_v1.Schema$ParagraphStyle, string]> = [
            [layout.block.indent_start_points, "indentStart", "indentStart"],
            [layout.block.indent_end_points, "indentEnd", "indentEnd"],
            [layout.block.indent_first_line_points, "indentFirstLine", "indentFirstLine"],
            [layout.block.space_above_points, "spaceAbove", "spaceAbove"],
            [layout.block.space_below_points, "spaceBelow", "spaceBelow"],
          ];
          for (const [value, key, field] of dimensions) {
            if (value !== undefined) {
              (paragraphStyle as Record<string, unknown>)[key] = { magnitude: value, unit: "PT" };
              fields.push(field);
            } else if (!inherit_neighbor_style && (field === "indentStart" || field === "indentFirstLine")) {
              fields.push(field);
            }
          }
          if (layout.block.line_spacing !== undefined) {
            paragraphStyle.lineSpacing = layout.block.line_spacing;
            fields.push("lineSpacing");
          }
          if (fields.length > 0) {
            requests.push({
              updateParagraphStyle: {
                range: { tabId: target.tabId, startIndex: layout.startIndex, endIndex: layout.endIndex },
                paragraphStyle,
                fields: fields.join(","),
              },
            });
          }

          for (const range of layout.segmentRanges) {
            const style: docs_v1.Schema$TextStyle = {};
            const styleFields: string[] = [];
            const segment = range.segment;
            for (const [field, value] of [
              ["bold", segment.bold],
              ["italic", segment.italic],
              ["underline", segment.underline],
              ["strikethrough", segment.strikethrough],
            ] as const) {
              if (value !== undefined) {
                (style as Record<string, unknown>)[field] = value;
                styleFields.push(field);
              }
            }
            if (segment.font_family) {
              style.weightedFontFamily = { fontFamily: segment.font_family };
              styleFields.push("weightedFontFamily.fontFamily");
            }
            if (segment.font_size !== undefined) {
              style.fontSize = { magnitude: segment.font_size, unit: "PT" };
              styleFields.push("fontSize");
            }
            if (segment.foreground_color) {
              style.foregroundColor = DocsClient.hexToOptionalColor(segment.foreground_color);
              styleFields.push("foregroundColor");
            }
            if (segment.background_color) {
              style.backgroundColor = DocsClient.hexToOptionalColor(segment.background_color);
              styleFields.push("backgroundColor");
            }
            if (segment.link_url) {
              style.link = { url: segment.link_url };
              styleFields.push("link");
            }
            if (styleFields.length > 0 && range.endIndex > range.startIndex) {
              requests.push({
                updateTextStyle: {
                  range: { tabId: target.tabId, startIndex: range.startIndex, endIndex: range.endIndex },
                  textStyle: style,
                  fields: styleFields.join(","),
                },
              });
            }
          }
        }

        const listGroups: Array<{ preset: "BULLETED" | "NUMBERED" | "CHECKBOX"; startIndex: number; endIndex: number }> = [];
        for (const layout of layouts) {
          if (!layout.block.list) continue;
          const previous = listGroups[listGroups.length - 1];
          if (previous && previous.preset === layout.block.list.preset && previous.endIndex === layout.startIndex) {
            previous.endIndex = layout.endIndex;
          } else {
            listGroups.push({
              preset: layout.block.list.preset,
              startIndex: layout.startIndex,
              endIndex: layout.endIndex,
            });
          }
        }
        for (const group of listGroups.sort((a, b) => b.startIndex - a.startIndex)) {
          requests.push({
            createParagraphBullets: {
              range: { tabId: target.tabId, startIndex: group.startIndex, endIndex: group.endIndex },
              bulletPreset: docBulletPreset(group.preset),
            },
          });
        }

        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const result = await docsClient.batchUpdateRequests(
          document_id,
          requests,
          revisionId,
          conflict_mode as DocsConflictMode,
          false,
        );
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: target.tabId,
          revisionId: result.revisionId,
          insertedRange: {
            startIndex: insertionIndex,
            endIndex: insertionIndex + insertedText.length,
          },
          blocksInserted: blocks.length,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_batch_update_doc",
    "Apply up to 100 disjoint anchored text, formatting, paragraph, or list changes against one immutable document snapshot.",
    {
      document_id: z.string(),
      tab_id: z.string().optional(),
      operations: z.array(docBatchOperationSchema).min(1).max(100),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ document_id, tab_id, operations, conflict_mode }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }
      try {
        const requests: docs_v1.Schema$Request[] = [];
        const mutationRanges: Array<{ startIndex: number; endIndex: number }> = [];
        const resolvedOperations: Array<Record<string, unknown>> = [];

        for (const operation of operations) {
          const resolved = await resolveDocumentRange({
            documentId: document_id,
            tabId: tab_id,
            targetText: operation.target_text,
            occurrence: operation.occurrence,
            matchCase: operation.match_case,
            expectedText: operation.expected_text,
          });

          if (operation.type === "replace" || operation.type === "delete") {
            const normalized = normalizeRangeForTextMutation(resolved);
            if (
              mutationRanges.some(
                (range) =>
                  range.endIndex > normalized.startIndex &&
                  range.startIndex < normalized.endIndex,
              )
            ) {
              throw new Error("Batch content mutations must target disjoint ranges.");
            }
            mutationRanges.push(normalized);
            requests.push({
              deleteContentRange: {
                range: {
                  tabId: resolved.tabId,
                  startIndex: normalized.startIndex,
                  endIndex: normalized.endIndex,
                },
              },
            });
            if (operation.type === "replace" && operation.replacement_text.length > 0) {
              requests.push({
                insertText: {
                  location: { tabId: resolved.tabId, index: normalized.startIndex },
                  text: operation.replacement_text,
                },
              });
            }
            resolvedOperations.push({
              type: operation.type,
              range: { startIndex: normalized.startIndex, endIndex: normalized.endIndex },
            });
            continue;
          }

          if (operation.type === "text_style") {
            const style: docs_v1.Schema$TextStyle = {};
            const fields: string[] = [];
            for (const [field, value] of [
              ["bold", operation.bold],
              ["italic", operation.italic],
              ["underline", operation.underline],
              ["strikethrough", operation.strikethrough],
            ] as const) {
              if (value !== undefined) {
                (style as Record<string, unknown>)[field] = value;
                fields.push(field);
              }
            }
            if (operation.foreground_color) {
              style.foregroundColor = DocsClient.hexToOptionalColor(operation.foreground_color);
              fields.push("foregroundColor");
            }
            if (operation.background_color) {
              style.backgroundColor = DocsClient.hexToOptionalColor(operation.background_color);
              fields.push("backgroundColor");
            }
            if (operation.link_url) {
              style.link = { url: operation.link_url };
              fields.push("link");
            }
            if (fields.length === 0) {
              throw new Error("A text_style operation must include at least one style field.");
            }
            requests.push({
              updateTextStyle: {
                range: {
                  tabId: resolved.tabId,
                  startIndex: resolved.startIndex,
                  endIndex: resolved.endIndex,
                },
                textStyle: style,
                fields: fields.join(","),
              },
            });
          } else {
            const paragraphRange = snapRangeToParagraphs(
              resolved.snapshot,
              resolved.startIndex,
              resolved.endIndex,
            );
            if (operation.type === "paragraph_style") {
              const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
              const fields: string[] = [];
              if (operation.named_style_type) {
                paragraphStyle.namedStyleType = operation.named_style_type;
                fields.push("namedStyleType");
              }
              if (operation.alignment) {
                paragraphStyle.alignment = operation.alignment;
                fields.push("alignment");
              }
              if (fields.length === 0) {
                throw new Error("A paragraph_style operation must include a style field.");
              }
              requests.push({
                updateParagraphStyle: {
                  range: { tabId: resolved.tabId, ...paragraphRange },
                  paragraphStyle,
                  fields: fields.join(","),
                },
              });
            } else {
              requests.push(operation.preset === "REMOVE"
                ? { deleteParagraphBullets: { range: { tabId: resolved.tabId, ...paragraphRange } } }
                : {
                    createParagraphBullets: {
                      range: { tabId: resolved.tabId, ...paragraphRange },
                      bulletPreset: docBulletPreset(operation.preset),
                    },
                  });
            }
          }
          resolvedOperations.push({
            type: operation.type,
            range: { startIndex: resolved.startIndex, endIndex: resolved.endIndex },
          });
        }

        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const result = await docsClient.batchUpdateRequests(
          document_id,
          requests,
          revisionId,
          conflict_mode as DocsConflictMode,
          true,
        );
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          revisionId: result.revisionId,
          operations: resolvedOperations,
        });
      } catch (err) {
        return errorResult(rewriteTerminalNewlineMutationError(err));
      }
    },
  );

  server.tool(
    "gdrive_insert_doc_table",
    "Insert a native Google Docs table at a position or text anchor and optionally populate its cells.",
    {
      document_id: z.string(),
      tab_id: z.string().optional(),
      rows: z.number().int().min(1).max(100),
      columns: z.number().int().min(1).max(100),
      values: z.array(z.array(z.string())).optional(),
      index: z.number().int().min(0).optional(),
      position: z.enum(["start", "end"]).optional(),
      before_text: z.string().optional(),
      after_text: z.string().optional(),
      occurrence: z.number().int().min(1).optional(),
      match_case: z.boolean().default(true),
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
      rows,
      columns,
      values,
      index,
      position,
      before_text,
      after_text,
      occurrence,
      match_case,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) return errorResult(unreadDocumentError());
      if (rows * columns > 1_000) {
        return errorResult(new Error("Create at most 1000 table cells per call."));
      }
      if (values && (values.length > rows || values.some((row) => row.length > columns))) {
        return errorResult(new Error("values must fit within the requested table dimensions."));
      }
      let insertedTable: { tabId: string; tablePath: number[]; revisionId?: string } | undefined;
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const first = await docsClient.batchUpdateRequests(
          document_id,
          [{
            insertTable: target.atEnd
              ? {
                  rows,
                  columns,
                  endOfSegmentLocation: { tabId: target.tabId },
                }
              : {
                  rows,
                  columns,
                  location: { tabId: target.tabId, index: target.index },
                },
          }],
          revisionId,
          conflict_mode as DocsConflictMode,
          false,
        );
        rememberDocumentRead(document_id, first.revisionId);
        insertedTable = {
          tabId: target.tabId,
          tablePath: [],
          revisionId: first.revisionId,
        };

        const afterInsert = await getStructuredDocumentTab(document_id, target.tabId);
        const candidates = collectTables(afterInsert.tab.blocks).filter(
          (table) => table.tablePath.length === 1 && table.rows === rows && table.columns === columns,
        );
        const table = target.atEnd
          ? candidates[candidates.length - 1]
          : candidates.reduce<NormalizedDocTableBlock | undefined>((best, candidate) => {
              if (!best) return candidate;
              const wanted = target.index ?? 0;
              return Math.abs(candidate.startIndex - wanted) < Math.abs(best.startIndex - wanted)
                ? candidate
                : best;
            }, undefined);
        if (!table) {
          throw new Error(
            "The table was inserted, but its structure could not be resolved. Reread the document before continuing.",
          );
        }
        insertedTable = {
          tabId: target.tabId,
          tablePath: table.tablePath,
          revisionId: first.revisionId,
        };

        let finalRevisionId = first.revisionId;
        if (values?.some((row) => row.some((value) => value.length > 0))) {
          const requests: docs_v1.Schema$Request[] = [];
          for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
            for (let columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex++) {
              const text = values[rowIndex][columnIndex];
              if (!text) continue;
              const cell = resolveTableCell(table, rowIndex, columnIndex);
              requests.push({
                insertText: {
                  location: { tabId: target.tabId, index: cell.startIndex + 1 },
                  text,
                },
              });
            }
          }
          if (requests.length > 0) {
            const secondRevision = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
            const populated = await docsClient.batchUpdateRequests(
              document_id,
              requests,
              secondRevision,
              conflict_mode as DocsConflictMode,
              true,
            );
            finalRevisionId = populated.revisionId;
            rememberDocumentRead(document_id, populated.revisionId);
          }
        }

        const refreshed = await refreshTable(document_id, target.tabId, table.tablePath);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: target.tabId,
          revisionId: finalRevisionId,
          table: refreshed,
        });
      } catch (err) {
        if (insertedTable) {
          const message = err instanceof Error ? err.message : String(err);
          const location = insertedTable.tablePath.length > 0
            ? `at path [${insertedTable.tablePath.join(", ")}]`
            : "at an unresolved path";
          return errorResult(
            new Error(
              `TABLE_INSERT_PARTIAL: The native table was inserted ${location} ` +
                `in tab ${JSON.stringify(insertedTable.tabId)}, but a later population or verification step failed: ${message}. ` +
                "Reread the document before completing the table.",
            ),
          );
        }
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_update_doc_table_cells",
    "Replace multiple native table cells atomically using table paths and zero-based row/column coordinates.",
    {
      document_id: z.string(),
      tab_id: z.string(),
      table_path: docTablePathSchema,
      updates: z.array(z.object({
        row_index: z.number().int().min(0),
        column_index: z.number().int().min(0),
        text: z.string(),
        expected_text: z.string().optional(),
      })).min(1).max(500),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ document_id, tab_id, table_path, updates, conflict_mode }) => {
      if (!accessedDocs.has(document_id)) return errorResult(unreadDocumentError());
      try {
        const snapshot = await getStructuredDocumentTab(document_id, tab_id);
        const table = resolveTableByPath(snapshot, table_path);
        const seen = new Set<string>();
        const requests: docs_v1.Schema$Request[] = [];
        for (const update of updates) {
          const key = `${update.row_index}:${update.column_index}`;
          if (seen.has(key)) throw new Error(`Cell ${key} is updated more than once.`);
          seen.add(key);
          const cell = resolveTableCell(table, update.row_index, update.column_index);
          const currentText = cell.text.endsWith("\n") ? cell.text.slice(0, -1) : cell.text;
          if (update.expected_text !== undefined && update.expected_text !== currentText) {
            throw new Error(
              `STALE_TABLE_CELL: Cell (${update.row_index}, ${update.column_index}) contains ${JSON.stringify(currentText)}, not ${JSON.stringify(update.expected_text)}.`,
            );
          }
          const contentStart = cell.startIndex + 1;
          const contentEnd = Math.max(contentStart, cell.endIndex - 1);
          if (contentEnd > contentStart) {
            requests.push({
              deleteContentRange: {
                range: { tabId: tab_id, startIndex: contentStart, endIndex: contentEnd },
              },
            });
          }
          if (update.text.length > 0) {
            requests.push({
              insertText: {
                location: { tabId: tab_id, index: contentStart },
                text: update.text,
              },
            });
          }
        }
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const result = await docsClient.batchUpdateRequests(
          document_id,
          requests,
          revisionId,
          conflict_mode as DocsConflictMode,
          true,
        );
        rememberDocumentRead(document_id, result.revisionId);
        const refreshed = await refreshTable(document_id, tab_id, table_path);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: tab_id,
          revisionId: result.revisionId,
          table: refreshed,
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_modify_doc_table",
    "Insert or delete rows/columns, merge or unmerge cells, or delete a native Google Docs table. Operations run in the supplied order.",
    {
      document_id: z.string(),
      tab_id: z.string(),
      table_path: docTablePathSchema,
      operations: z.array(docTableStructureOperationSchema).min(1).max(100),
      conflict_mode: docConflictModeSchema,
    },
    {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    } satisfies ToolAnnotations,
    async ({ document_id, tab_id, table_path, operations, conflict_mode }) => {
      if (!accessedDocs.has(document_id)) return errorResult(unreadDocumentError());
      try {
        const snapshot = await getStructuredDocumentTab(document_id, tab_id);
        const table = resolveTableByPath(snapshot, table_path);
        const tableStartLocation = { tabId: tab_id, index: table.startIndex };
        const requests: docs_v1.Schema$Request[] = operations.map((operation) => {
          switch (operation.type) {
            case "insert_row":
              return {
                insertTableRow: {
                  tableCellLocation: {
                    tableStartLocation,
                    rowIndex: operation.row_index,
                    columnIndex: 0,
                  },
                  insertBelow: operation.position === "after",
                },
              };
            case "delete_row":
              return {
                deleteTableRow: {
                  tableCellLocation: {
                    tableStartLocation,
                    rowIndex: operation.row_index,
                    columnIndex: 0,
                  },
                },
              };
            case "insert_column":
              return {
                insertTableColumn: {
                  tableCellLocation: {
                    tableStartLocation,
                    rowIndex: 0,
                    columnIndex: operation.column_index,
                  },
                  insertRight: operation.position === "after",
                },
              };
            case "delete_column":
              return {
                deleteTableColumn: {
                  tableCellLocation: {
                    tableStartLocation,
                    rowIndex: 0,
                    columnIndex: operation.column_index,
                  },
                },
              };
            case "merge":
            case "unmerge": {
              const tableRange = {
                tableCellLocation: {
                  tableStartLocation,
                  rowIndex: operation.start_row,
                  columnIndex: operation.start_column,
                },
                rowSpan: operation.row_span,
                columnSpan: operation.column_span,
              };
              return operation.type === "merge"
                ? { mergeTableCells: { tableRange } }
                : { unmergeTableCells: { tableRange } };
            }
            case "delete_table":
              return {
                deleteContentRange: {
                  range: {
                    tabId: tab_id,
                    startIndex: table.startIndex,
                    endIndex: table.endIndex,
                  },
                },
              };
          }
        });
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const result = await docsClient.batchUpdateRequests(
          document_id,
          requests,
          revisionId,
          conflict_mode as DocsConflictMode,
          false,
        );
        rememberDocumentRead(document_id, result.revisionId);
        const deleted = operations.some((operation) => operation.type === "delete_table");
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: tab_id,
          revisionId: result.revisionId,
          deleted,
          table: deleted ? undefined : await refreshTable(document_id, tab_id, table_path),
        });
      } catch (err) {
        return errorResult(err);
      }
    },
  );

  server.tool(
    "gdrive_format_doc_table",
    "Format native table cells, columns, and rows using a table path and zero-based coordinates.",
    {
      document_id: z.string(),
      tab_id: z.string(),
      table_path: docTablePathSchema,
      cell_range: z.object({
        start_row: z.number().int().min(0),
        start_column: z.number().int().min(0),
        row_span: z.number().int().min(1),
        column_span: z.number().int().min(1),
      }).optional(),
      background_color: docColorSchema.optional(),
      content_alignment: z.enum(["TOP", "MIDDLE", "BOTTOM"]).optional(),
      padding_top_points: z.number().nonnegative().optional(),
      padding_bottom_points: z.number().nonnegative().optional(),
      padding_left_points: z.number().nonnegative().optional(),
      padding_right_points: z.number().nonnegative().optional(),
      border_color: docColorSchema.optional(),
      border_width_points: z.number().nonnegative().optional(),
      border_dash_style: z.enum(["SOLID", "DOT", "DASH"]).optional(),
      column_widths: z.array(z.object({
        column_index: z.number().int().min(0),
        width_points: z.number().min(5),
      })).max(100).optional(),
      row_styles: z.array(z.object({
        row_index: z.number().int().min(0),
        min_height_points: z.number().nonnegative().optional(),
        prevent_overflow: z.boolean().optional(),
      })).max(100).optional(),
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
      table_path,
      cell_range,
      background_color,
      content_alignment,
      padding_top_points,
      padding_bottom_points,
      padding_left_points,
      padding_right_points,
      border_color,
      border_width_points,
      border_dash_style,
      column_widths,
      row_styles,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) return errorResult(unreadDocumentError());
      try {
        const snapshot = await getStructuredDocumentTab(document_id, tab_id);
        const table = resolveTableByPath(snapshot, table_path);
        const tableStartLocation = { tabId: tab_id, index: table.startIndex };
        const requests: docs_v1.Schema$Request[] = [];
        const cellStyle: docs_v1.Schema$TableCellStyle = {};
        const cellFields: string[] = [];
        if (background_color) {
          cellStyle.backgroundColor = DocsClient.hexToOptionalColor(background_color);
          cellFields.push("backgroundColor");
        }
        if (content_alignment) {
          cellStyle.contentAlignment = content_alignment;
          cellFields.push("contentAlignment");
        }
        for (const [value, key, field] of [
          [padding_top_points, "paddingTop", "paddingTop"],
          [padding_bottom_points, "paddingBottom", "paddingBottom"],
          [padding_left_points, "paddingLeft", "paddingLeft"],
          [padding_right_points, "paddingRight", "paddingRight"],
        ] as const) {
          if (value !== undefined) {
            (cellStyle as Record<string, unknown>)[key] = { magnitude: value, unit: "PT" };
            cellFields.push(field);
          }
        }
        if (border_color || border_width_points !== undefined || border_dash_style) {
          const border: docs_v1.Schema$TableCellBorder = {
            color: border_color ? DocsClient.hexToOptionalColor(border_color) : undefined,
            width: border_width_points !== undefined
              ? { magnitude: border_width_points, unit: "PT" }
              : undefined,
            dashStyle: border_dash_style,
          };
          cellStyle.borderTop = border;
          cellStyle.borderBottom = border;
          cellStyle.borderLeft = border;
          cellStyle.borderRight = border;
          cellFields.push("borderTop", "borderBottom", "borderLeft", "borderRight");
        }
        if (cellFields.length > 0) {
          if (!cell_range) {
            requests.push({
              updateTableCellStyle: {
                tableStartLocation,
                tableCellStyle: cellStyle,
                fields: cellFields.join(","),
              },
            });
          } else {
            requests.push({
              updateTableCellStyle: {
                tableRange: {
                  tableCellLocation: {
                    tableStartLocation,
                    rowIndex: cell_range.start_row,
                    columnIndex: cell_range.start_column,
                  },
                  rowSpan: cell_range.row_span,
                  columnSpan: cell_range.column_span,
                },
                tableCellStyle: cellStyle,
                fields: cellFields.join(","),
              },
            });
          }
        }
        for (const column of column_widths ?? []) {
          requests.push({
            updateTableColumnProperties: {
              tableStartLocation,
              columnIndices: [column.column_index],
              tableColumnProperties: {
                widthType: "FIXED_WIDTH",
                width: { magnitude: column.width_points, unit: "PT" },
              },
              fields: "widthType,width",
            },
          });
        }
        for (const row of row_styles ?? []) {
          const tableRowStyle: docs_v1.Schema$TableRowStyle = {};
          const fields: string[] = [];
          if (row.min_height_points !== undefined) {
            tableRowStyle.minRowHeight = { magnitude: row.min_height_points, unit: "PT" };
            fields.push("minRowHeight");
          }
          if (row.prevent_overflow !== undefined) {
            tableRowStyle.preventOverflow = row.prevent_overflow;
            fields.push("preventOverflow");
          }
          if (fields.length > 0) {
            requests.push({
              updateTableRowStyle: {
                tableStartLocation,
                rowIndices: [row.row_index],
                tableRowStyle,
                fields: fields.join(","),
              },
            });
          }
        }
        if (requests.length === 0) {
          throw new Error("Provide at least one table formatting change.");
        }
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const result = await docsClient.batchUpdateRequests(
          document_id,
          requests,
          revisionId,
          conflict_mode as DocsConflictMode,
          false,
        );
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: tab_id,
          revisionId: result.revisionId,
          table: await refreshTable(document_id, tab_id, table_path),
        });
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
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
      indent_start_points: z.number().nonnegative().nullable().optional().describe("Start indent in points, or null to reset"),
      indent_end_points: z.number().nonnegative().nullable().optional().describe("End indent in points, or null to reset"),
      indent_first_line_points: z.number().nullable().optional().describe("First-line indent in points, or null to reset"),
      space_above_points: z.number().nonnegative().nullable().optional().describe("Space above in points, or null to reset"),
      space_below_points: z.number().nonnegative().nullable().optional().describe("Space below in points, or null to reset"),
      line_spacing: z.number().positive().nullable().optional().describe("Line spacing percentage, or null to reset"),
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
      indent_start_points,
      indent_end_points,
      indent_first_line_points,
      space_above_points,
      space_below_points,
      line_spacing,
      conflict_mode,
    }) => {
      if (!accessedDocs.has(document_id)) {
        return errorResult(unreadDocumentError());
      }

      if (
        named_style_type === undefined &&
        alignment === undefined &&
        indent_start_points === undefined &&
        indent_end_points === undefined &&
        indent_first_line_points === undefined &&
        space_above_points === undefined &&
        space_below_points === undefined &&
        line_spacing === undefined
      ) {
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const result = await docsClient.updateParagraphStyle({
          documentId: document_id,
          tabId: resolved.tabId,
          startIndex: paragraphRange.startIndex,
          endIndex: paragraphRange.endIndex,
          namedStyleType: named_style_type as DocNamedStyleType | undefined,
          alignment: alignment as DocParagraphAlignment | undefined,
          indentStartPoints: indent_start_points,
          indentEndPoints: indent_end_points,
          indentFirstLinePoints: indent_first_line_points,
          spaceAbovePoints: space_above_points,
          spaceBelowPoints: space_below_points,
          lineSpacing: line_spacing,
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
          indentStartPoints: indent_start_points,
          indentEndPoints: indent_end_points,
          indentFirstLinePoints: indent_first_line_points,
          spaceAbovePoints: space_above_points,
          spaceBelowPoints: space_below_points,
          lineSpacing: line_spacing,
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
      nesting_levels: z
        .array(z.number().int().min(0).max(8))
        .optional()
        .describe("One nesting level (0-8) for each selected paragraph"),
      continue_previous: z
        .boolean()
        .default(false)
        .describe("Require this block to continue an immediately adjacent compatible native list"),
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
      nesting_levels,
      continue_previous,
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
        const revisionId = await ensureDocumentRevision(document_id, conflict_mode as DocsConflictMode);
        const paragraphs = (resolved.snapshot.tab.paragraphs ?? []).filter(
          (paragraph) =>
            paragraph.endIndex > paragraphRange.startIndex &&
            paragraph.startIndex < paragraphRange.endIndex,
        );

        if (nesting_levels && nesting_levels.length !== paragraphs.length) {
          throw new Error(
            `nesting_levels must contain one value per selected paragraph (${paragraphs.length}).`,
          );
        }
        if (
          preset === "NUMBERED" &&
          nesting_levels &&
          paragraphs.some((paragraph) => paragraph.displayText.length === 0)
        ) {
          throw new Error(
            "Cannot guarantee native numbered-list continuity across a blank paragraph. Use paragraph spacing instead.",
          );
        }
        if (continue_previous) {
          if (preset === "REMOVE") {
            throw new Error("continue_previous cannot be used with preset REMOVE.");
          }
          const firstIndex = (resolved.snapshot.tab.paragraphs ?? []).indexOf(paragraphs[0]);
          const previous = firstIndex > 0
            ? (resolved.snapshot.tab.paragraphs ?? [])[firstIndex - 1]
            : undefined;
          if (!previous?.list || previous.list.preset !== preset) {
            throw new Error(
              "Cannot continue numbering: the immediately preceding paragraph is not a compatible native list item.",
            );
          }
        }

        let result;
        if (nesting_levels && preset !== "REMOVE") {
          const totalTabs = nesting_levels.reduce((sum, level) => sum + level, 0);
          const requests: docs_v1.Schema$Request[] = [
            {
              deleteParagraphBullets: {
                range: { tabId: resolved.tabId, ...paragraphRange },
              },
            },
            {
              updateParagraphStyle: {
                range: { tabId: resolved.tabId, ...paragraphRange },
                paragraphStyle: {},
                fields: "indentStart,indentFirstLine",
              },
            },
            ...paragraphs
              .map((paragraph, index) => ({ paragraph, level: nesting_levels[index] }))
              .filter(({ level }) => level > 0)
              .sort((left, right) => right.paragraph.startIndex - left.paragraph.startIndex)
              .map(({ paragraph, level }) => ({
                insertText: {
                  location: { tabId: resolved.tabId, index: paragraph.startIndex },
                  text: "\t".repeat(level),
                },
              })),
            {
              createParagraphBullets: {
                range: {
                  tabId: resolved.tabId,
                  startIndex: paragraphRange.startIndex,
                  endIndex: paragraphRange.endIndex + totalTabs,
                },
                bulletPreset: docBulletPreset(preset as Exclude<DocListPreset, "REMOVE">),
              },
            },
          ];
          result = await docsClient.batchUpdateRequests(
            document_id,
            requests,
            revisionId,
            conflict_mode as DocsConflictMode,
            false,
          );
        } else {
          result = await docsClient.updateList({
            documentId: document_id,
            tabId: resolved.tabId,
            startIndex: paragraphRange.startIndex,
            endIndex: paragraphRange.endIndex,
            preset: preset as DocListPreset,
            revisionId,
            conflictMode: conflict_mode as DocsConflictMode,
          });
        }
        rememberDocumentRead(document_id, result.revisionId);
        return jsonResult({
          documentId: document_id,
          documentUrl: documentUrl(document_id),
          tabId: resolved.tabId,
          revisionId: result.revisionId,
          paragraphRange,
          preset,
          nestingLevels: nesting_levels,
          continuedPrevious: continue_previous,
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
