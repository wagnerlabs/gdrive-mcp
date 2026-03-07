import { google, docs_v1, drive_v3 } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import { DriveAPIError, handleApiError } from "./client.js";

const TAB_FIELD_DEPTH = 6;

export type DocsConflictMode = "strict" | "merge";
export type DocListPreset = "BULLETED" | "NUMBERED" | "CHECKBOX" | "REMOVE";
export type DocNamedStyleType =
  | "NORMAL_TEXT"
  | "TITLE"
  | "SUBTITLE"
  | "HEADING_1"
  | "HEADING_2"
  | "HEADING_3"
  | "HEADING_4"
  | "HEADING_5"
  | "HEADING_6";
export type DocParagraphAlignment = "START" | "CENTER" | "END" | "JUSTIFIED";

export interface NormalizedDocTextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSize?: number;
  foregroundColor?: string;
  backgroundColor?: string;
  link?: string;
}

export interface NormalizedDocTextRunElement {
  type: "textRun";
  startIndex: number;
  endIndex: number;
  text: string;
  textStyle: NormalizedDocTextStyle | null;
}

export interface NormalizedDocPlaceholderElement {
  type: "placeholder";
  startIndex: number;
  endIndex: number;
  placeholderKind: "table" | "inlineObject" | "horizontalRule" | "other";
  textStyle?: NormalizedDocTextStyle | null;
}

export type NormalizedDocElement =
  | NormalizedDocTextRunElement
  | NormalizedDocPlaceholderElement;

export interface NormalizedDocParagraph {
  startIndex: number;
  endIndex: number;
  // `displayText` omits the trailing paragraph newline so agents can safely
  // round-trip it into anchor-based write tools.
  displayText: string;
  // `text` preserves the raw Docs paragraph text, including a trailing newline
  // when present, so index arithmetic and structured snapshots stay faithful.
  text: string;
  namedStyleType?: DocNamedStyleType;
  alignment?: DocParagraphAlignment;
  list: {
    preset: "BULLETED" | "NUMBERED" | "CHECKBOX";
    nestingLevel: number;
    listId?: string;
  } | null;
  elements: NormalizedDocElement[];
}

export interface NormalizedDocTab {
  tabId: string;
  title: string;
  parentTabId?: string;
  index: number;
  nestingLevel: number;
  paragraphs?: NormalizedDocParagraph[];
}

export interface NormalizedDocument {
  documentId: string;
  title: string;
  documentUrl: string;
  revisionId?: string;
  tabs: NormalizedDocTab[];
  contentTruncated: boolean;
}

export interface GetDocumentOptions {
  includeContent?: boolean;
  tabId?: string;
  maxChars?: number;
  maxParagraphs?: number;
}

export interface DocsWriteResult {
  documentId: string;
  revisionId?: string;
  replies?: docs_v1.Schema$Response[];
}

export interface InsertTextOptions {
  documentId: string;
  text: string;
  tabId?: string;
  index?: number;
  atEnd?: boolean;
  revisionId?: string;
  conflictMode?: DocsConflictMode;
}

export interface ReplaceTextOptions {
  documentId: string;
  text: string;
  tabId?: string;
  startIndex: number;
  endIndex: number;
  revisionId?: string;
  conflictMode?: DocsConflictMode;
}

export interface DeleteTextOptions {
  documentId: string;
  tabId?: string;
  startIndex: number;
  endIndex: number;
  revisionId?: string;
  conflictMode?: DocsConflictMode;
}

export interface ReplaceAllTextOptions {
  documentId: string;
  searchText: string;
  replaceText: string;
  matchCase?: boolean;
  tabId?: string;
  allTabs?: boolean;
  revisionId?: string;
  conflictMode?: DocsConflictMode;
}

export interface TextStyleUpdateOptions {
  documentId: string;
  tabId?: string;
  startIndex: number;
  endIndex: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSize?: number;
  foregroundColor?: string;
  backgroundColor?: string;
  linkUrl?: string;
  revisionId?: string;
  conflictMode?: DocsConflictMode;
}

export interface ParagraphStyleUpdateOptions {
  documentId: string;
  tabId?: string;
  startIndex: number;
  endIndex: number;
  namedStyleType?: DocNamedStyleType;
  alignment?: DocParagraphAlignment;
  revisionId?: string;
  conflictMode?: DocsConflictMode;
}

export interface ListUpdateOptions {
  documentId: string;
  tabId?: string;
  startIndex: number;
  endIndex: number;
  preset: DocListPreset;
  revisionId?: string;
  conflictMode?: DocsConflictMode;
}

export interface CreateDocumentResult {
  documentId: string;
  title: string;
  documentUrl: string;
  revisionId?: string;
  folderId: string;
}

export interface RenameDocumentResult {
  documentId: string;
  title: string;
  documentUrl: string;
}

export interface DuplicateDocumentResult {
  documentId: string;
  title: string;
  documentUrl: string;
  revisionId?: string;
  sourceDocumentId: string;
  folderId?: string;
}

function documentUrl(documentId: string): string {
  return `https://docs.google.com/document/d/${documentId}/edit`;
}

function paragraphDisplayText(text: string): string {
  return text.endsWith("\n") ? text.slice(0, -1) : text;
}

function buildMetadataTabFields(depth: number): string {
  const current = "tabProperties";
  if (depth <= 0) {
    return current;
  }
  return `${current},childTabs(${buildMetadataTabFields(depth - 1)})`;
}

function buildContentTabFields(depth: number): string {
  // Request the full `lists` object here. Narrower nested list field masks are
  // brittle in tab-aware Docs reads, while the full object keeps list
  // normalization stable across includeTabsContent responses.
  const current = [
    "tabProperties",
    "documentTab(body(content(startIndex,endIndex,paragraph(elements(startIndex,endIndex,textRun(content,textStyle),autoText(type,textStyle),pageBreak(textStyle),columnBreak(textStyle),footnoteReference,horizontalRule,equation,inlineObjectElement,person,richLink),paragraphStyle(namedStyleType,alignment),bullet(listId,nestingLevel,textStyle)),sectionBreak,table,tableOfContents)),lists)",
  ].join(",");

  if (depth <= 0) {
    return current;
  }
  return `${current},childTabs(${buildContentTabFields(depth - 1)})`;
}

function buildDocumentFields(includeContent: boolean): string {
  const base = "documentId,title";
  const tabs = includeContent
    ? buildContentTabFields(TAB_FIELD_DEPTH)
    : buildMetadataTabFields(TAB_FIELD_DEPTH);
  return `${base},tabs(${tabs})`;
}

function isNonNullable<T>(value: T | null | undefined): value is T {
  return value !== undefined && value !== null;
}

export class DocsClient {
  private docs: docs_v1.Docs;
  private drive: drive_v3.Drive;

  constructor(auth: OAuth2Client) {
    this.docs = google.docs({ version: "v1", auth });
    this.drive = google.drive({ version: "v3", auth });
  }

  async getDocument(
    documentId: string,
    options: GetDocumentOptions = {},
  ): Promise<NormalizedDocument> {
    const includeContent = options.includeContent ?? false;
    const maxChars = options.maxChars ?? 20_000;
    const maxParagraphs = options.maxParagraphs ?? 200;

    try {
      // The Docs API rejects tab-aware partial responses that also request
      // revisionId, so fetch the tab tree first and fall back to a lightweight
      // revision-only lookup.
      const res = await this.docs.documents.get({
        documentId,
        includeTabsContent: true,
        suggestionsViewMode: "SUGGESTIONS_INLINE",
        fields: buildDocumentFields(includeContent),
      });

      const flatTabs = this.flattenTabs(res.data.tabs ?? []);
      if (options.tabId && !flatTabs.some((tab) => tab.tabProperties?.tabId === options.tabId)) {
        throw new DriveAPIError(`Tab "${options.tabId}" not found in document.`);
      }

      let remainingChars = maxChars;
      let remainingParagraphs = maxParagraphs;
      let contentTruncated = false;

      const tabs = flatTabs
        .map((tab) => {
          const tabId = tab.tabProperties?.tabId;
          if (!tabId) {
            return null;
          }

          const normalized: NormalizedDocTab = {
            tabId,
            title: tab.tabProperties?.title ?? "Untitled tab",
            parentTabId: tab.tabProperties?.parentTabId ?? undefined,
            index: tab.tabProperties?.index ?? 0,
            nestingLevel: tab.tabProperties?.nestingLevel ?? 0,
          };

          if (includeContent && (!options.tabId || options.tabId === tabId)) {
            const content = this.normalizeTabContent(
              tab.documentTab,
              remainingChars,
              remainingParagraphs,
            );
            normalized.paragraphs = content.paragraphs;
            remainingChars = content.remainingChars;
            remainingParagraphs = content.remainingParagraphs;
            contentTruncated = contentTruncated || content.contentTruncated;
          }

          return normalized;
        })
        .filter(isNonNullable);

      const resolvedDocumentId = res.data.documentId ?? documentId;
      const revisionId = res.data.revisionId ?? (await this.getRevisionId(resolvedDocumentId));

      return {
        documentId: resolvedDocumentId,
        title: res.data.title ?? "Untitled document",
        documentUrl: documentUrl(resolvedDocumentId),
        revisionId,
        tabs,
        contentTruncated,
      };
    } catch (err) {
      handleApiError(err, "Google Docs");
    }
  }

  async createDocument(
    title: string,
    folderId: string = "root",
  ): Promise<CreateDocumentResult> {
    try {
      const res = await this.docs.documents.create({
        requestBody: { title },
      });

      const documentId = res.data.documentId!;

      if (folderId !== "root") {
        const currentFile = await this.drive.files.get({
          fileId: documentId,
          fields: "parents",
          supportsAllDrives: true,
        });
        const removeParents = (currentFile.data.parents ?? []).join(",");
        await this.drive.files.update({
          fileId: documentId,
          addParents: folderId,
          removeParents: removeParents || undefined,
          fields: "id,parents",
          supportsAllDrives: true,
        });
      }

      return {
        documentId,
        title: res.data.title ?? title,
        documentUrl: documentUrl(documentId),
        revisionId: res.data.revisionId ?? (await this.getRevisionId(documentId)),
        folderId,
      };
    } catch (err) {
      handleApiError(err, "Google Docs");
    }
  }

  async getRevisionId(documentId: string): Promise<string | undefined> {
    try {
      const res = await this.docs.documents.get({
        documentId,
        suggestionsViewMode: "SUGGESTIONS_INLINE",
        fields: "revisionId",
      });
      return res.data.revisionId ?? undefined;
    } catch (err) {
      handleApiError(err, "Google Docs");
    }
  }

  async insertText(options: InsertTextOptions): Promise<DocsWriteResult> {
    if (options.index === undefined && !options.atEnd) {
      throw new DriveAPIError("insertText requires either an index or atEnd=true.");
    }

    const request: docs_v1.Schema$Request = options.atEnd
      ? {
          insertText: {
            text: options.text,
            endOfSegmentLocation: {
              tabId: options.tabId,
            },
          },
        }
      : {
          insertText: {
            text: options.text,
            location: {
              tabId: options.tabId,
              index: options.index,
            },
          },
        };

    return this.batchUpdateDocument(
      options.documentId,
      [request],
      options.revisionId,
      options.conflictMode,
    );
  }

  async replaceText(options: ReplaceTextOptions): Promise<DocsWriteResult> {
    const requests: docs_v1.Schema$Request[] = [
      {
        deleteContentRange: {
          range: {
            tabId: options.tabId,
            startIndex: options.startIndex,
            endIndex: options.endIndex,
          },
        },
      },
    ];

    if (options.text.length > 0) {
      requests.push({
        insertText: {
          text: options.text,
          location: {
            tabId: options.tabId,
            index: options.startIndex,
          },
        },
      });
    }

    return this.batchUpdateDocument(
      options.documentId,
      requests,
      options.revisionId,
      options.conflictMode,
    );
  }

  async replaceAllText(options: ReplaceAllTextOptions): Promise<DocsWriteResult> {
    if (!options.allTabs && !options.tabId) {
      throw new DriveAPIError(
        "replaceAllText requires tabId unless allTabs=true is explicitly set.",
      );
    }

    const request: docs_v1.Schema$Request = {
      replaceAllText: {
        replaceText: options.replaceText,
        containsText: {
          text: options.searchText,
          matchCase: options.matchCase ?? true,
        },
        tabsCriteria: options.allTabs
          ? undefined
          : {
              tabIds: [options.tabId!],
            },
      },
    };

    return this.batchUpdateDocument(
      options.documentId,
      [request],
      options.revisionId,
      options.conflictMode,
    );
  }

  async deleteText(options: DeleteTextOptions): Promise<DocsWriteResult> {
    return this.batchUpdateDocument(
      options.documentId,
      [
        {
          deleteContentRange: {
            range: {
              tabId: options.tabId,
              startIndex: options.startIndex,
              endIndex: options.endIndex,
            },
          },
        },
      ],
      options.revisionId,
      options.conflictMode,
    );
  }

  async updateTextStyle(
    options: TextStyleUpdateOptions,
  ): Promise<DocsWriteResult> {
    const { textStyle, fields } = this.buildTextStyleUpdate(options);
    return this.batchUpdateDocument(
      options.documentId,
      [
        {
          updateTextStyle: {
            textStyle,
            fields,
            range: {
              tabId: options.tabId,
              startIndex: options.startIndex,
              endIndex: options.endIndex,
            },
          },
        },
      ],
      options.revisionId,
      options.conflictMode,
    );
  }

  async updateParagraphStyle(
    options: ParagraphStyleUpdateOptions,
  ): Promise<DocsWriteResult> {
    const { paragraphStyle, fields } = this.buildParagraphStyleUpdate(options);
    return this.batchUpdateDocument(
      options.documentId,
      [
        {
          updateParagraphStyle: {
            paragraphStyle,
            fields,
            range: {
              tabId: options.tabId,
              startIndex: options.startIndex,
              endIndex: options.endIndex,
            },
          },
        },
      ],
      options.revisionId,
      options.conflictMode,
    );
  }

  async updateList(options: ListUpdateOptions): Promise<DocsWriteResult> {
    const range = {
      tabId: options.tabId,
      startIndex: options.startIndex,
      endIndex: options.endIndex,
    };

    const request: docs_v1.Schema$Request =
      options.preset === "REMOVE"
        ? {
            deleteParagraphBullets: {
              range,
            },
          }
        : {
            createParagraphBullets: {
              range,
              bulletPreset: this.mapListPreset(options.preset),
            },
          };

    return this.batchUpdateDocument(
      options.documentId,
      [request],
      options.revisionId,
      options.conflictMode,
    );
  }

  async renameDocument(
    documentId: string,
    newTitle: string,
  ): Promise<RenameDocumentResult> {
    try {
      const res = await this.drive.files.update({
        fileId: documentId,
        requestBody: { name: newTitle },
        fields: "id,name",
        supportsAllDrives: true,
      });

      return {
        documentId: res.data.id ?? documentId,
        title: res.data.name ?? newTitle,
        documentUrl: documentUrl(res.data.id ?? documentId),
      };
    } catch (err) {
      handleApiError(err, "Google Drive");
    }
  }

  async duplicateDocument(
    documentId: string,
    newTitle?: string,
    folderId?: string,
  ): Promise<DuplicateDocumentResult> {
    try {
      let destinationParents: string[] | undefined;

      if (folderId) {
        destinationParents = [folderId];
      } else {
        const source = await this.drive.files.get({
          fileId: documentId,
          fields: "parents",
          supportsAllDrives: true,
        });
        destinationParents = source.data.parents ?? undefined;
      }

      const res = await this.drive.files.copy({
        fileId: documentId,
        requestBody: {
          name: newTitle,
          parents: destinationParents,
        },
        fields: "id,name,parents",
        supportsAllDrives: true,
      });

      const copiedDocumentId = res.data.id!;
      return {
        documentId: copiedDocumentId,
        title: res.data.name ?? newTitle ?? "Copy of document",
        documentUrl: documentUrl(copiedDocumentId),
        revisionId: await this.getRevisionId(copiedDocumentId),
        sourceDocumentId: documentId,
        folderId: folderId ?? destinationParents?.[0] ?? undefined,
      };
    } catch (err) {
      handleApiError(err, "Google Drive");
    }
  }

  private async batchUpdateDocument(
    documentId: string,
    requests: docs_v1.Schema$Request[],
    revisionId?: string,
    conflictMode: DocsConflictMode = "strict",
  ): Promise<DocsWriteResult> {
    try {
      const sortedRequests = this.sortRequestsForBatchUpdate(requests);

      const res = await this.docs.documents.batchUpdate({
        documentId,
        requestBody: {
          requests: sortedRequests,
          writeControl: revisionId
            ? conflictMode === "merge"
              ? { targetRevisionId: revisionId }
              : { requiredRevisionId: revisionId }
            : undefined,
        },
      });

      const nextRevisionId =
        res.data.writeControl?.requiredRevisionId ??
        res.data.writeControl?.targetRevisionId ??
        (await this.getRevisionId(documentId));

      return {
        documentId: res.data.documentId ?? documentId,
        revisionId: nextRevisionId,
        replies: res.data.replies ?? undefined,
      };
    } catch (err) {
      handleApiError(err, "Google Docs");
    }
  }

  private getRequestIndex(request: docs_v1.Schema$Request): number {
    return (
      request.deleteContentRange?.range?.startIndex ??
      request.insertText?.location?.index ??
      request.updateTextStyle?.range?.startIndex ??
      request.updateParagraphStyle?.range?.startIndex ??
      request.createParagraphBullets?.range?.startIndex ??
      request.deleteParagraphBullets?.range?.startIndex ??
      Number.NEGATIVE_INFINITY
    );
  }

  private sortRequestsForBatchUpdate(
    requests: docs_v1.Schema$Request[],
  ): docs_v1.Schema$Request[] {
    return requests
      .map((request, originalIndex) => ({ request, originalIndex }))
      .sort((left, right) => {
        const indexDifference =
          this.getRequestIndex(right.request) - this.getRequestIndex(left.request);
        if (indexDifference !== 0) {
          return indexDifference;
        }

        // Preserve the caller's order for equal indices so delete/insert
        // replacements at the same location remain deterministic.
        return left.originalIndex - right.originalIndex;
      })
      .map(({ request }) => request);
  }

  private flattenTabs(tabs: docs_v1.Schema$Tab[]): docs_v1.Schema$Tab[] {
    const result: docs_v1.Schema$Tab[] = [];
    for (const tab of tabs) {
      result.push(tab);
      if (tab.childTabs?.length) {
        result.push(...this.flattenTabs(tab.childTabs));
      }
    }
    return result;
  }

  private normalizeTabContent(
    documentTab: docs_v1.Schema$DocumentTab | undefined,
    remainingChars: number,
    remainingParagraphs: number,
  ): {
    paragraphs: NormalizedDocParagraph[];
    remainingChars: number;
    remainingParagraphs: number;
    contentTruncated: boolean;
  } {
    const paragraphs: NormalizedDocParagraph[] = [];
    let contentTruncated = false;
    const listPresetMap = this.buildListPresetMap(documentTab?.lists ?? undefined);

    for (const element of documentTab?.body?.content ?? []) {
      if (remainingParagraphs <= 0) {
        contentTruncated = true;
        break;
      }

      const paragraph = this.normalizeStructuralElement(element, listPresetMap);
      if (!paragraph) {
        continue;
      }

      if (paragraph.text.length > remainingChars && paragraphs.length > 0) {
        contentTruncated = true;
        break;
      }

      paragraphs.push(paragraph);
      remainingParagraphs -= 1;
      remainingChars = Math.max(0, remainingChars - paragraph.text.length);

      if (remainingChars === 0 && paragraph.text.length > 0) {
        contentTruncated = true;
        break;
      }
    }

    return {
      paragraphs,
      remainingChars,
      remainingParagraphs,
      contentTruncated,
    };
  }

  private normalizeStructuralElement(
    element: docs_v1.Schema$StructuralElement,
    listPresetMap: Map<string, "BULLETED" | "NUMBERED" | "CHECKBOX">,
  ): NormalizedDocParagraph | null {
    const startIndex = element.startIndex ?? 0;
    const endIndex = element.endIndex ?? startIndex;

    if (element.paragraph) {
      return this.normalizeParagraph(
        element.paragraph,
        startIndex,
        endIndex,
        listPresetMap,
      );
    }

    if (element.table || element.tableOfContents || element.sectionBreak) {
      return {
        startIndex,
        endIndex,
        displayText: "",
        text: "",
        list: null,
        elements: [
          {
            type: "placeholder",
            startIndex,
            endIndex,
            placeholderKind: element.table ? "table" : "other",
          },
        ],
      };
    }

    return null;
  }

  private normalizeParagraph(
    paragraph: docs_v1.Schema$Paragraph,
    startIndex: number,
    endIndex: number,
    listPresetMap: Map<string, "BULLETED" | "NUMBERED" | "CHECKBOX">,
  ): NormalizedDocParagraph {
    const elements: NormalizedDocElement[] = [];
    let text = "";

    for (const element of paragraph.elements ?? []) {
      const elementStart = element.startIndex ?? startIndex;
      const elementEnd = element.endIndex ?? elementStart;

      if (element.textRun?.content !== undefined) {
        const runText = element.textRun.content ?? "";
        elements.push({
          type: "textRun",
          startIndex: elementStart,
          endIndex: elementEnd,
          text: runText,
          textStyle: this.normalizeTextStyle(element.textRun.textStyle),
        });
        text += runText;
        continue;
      }

      elements.push({
        type: "placeholder",
        startIndex: elementStart,
        endIndex: elementEnd,
        placeholderKind: this.getParagraphPlaceholderKind(element),
        textStyle: this.normalizeParagraphElementStyle(element),
      });
    }

    const listId = paragraph.bullet?.listId ?? undefined;
    const nestingLevel = paragraph.bullet?.nestingLevel ?? 0;

    return {
      startIndex,
      endIndex,
      displayText: paragraphDisplayText(text),
      text,
      namedStyleType: paragraph.paragraphStyle?.namedStyleType as
        | DocNamedStyleType
        | undefined,
      alignment: paragraph.paragraphStyle?.alignment as
        | DocParagraphAlignment
        | undefined,
      list: listId
        ? {
            preset: listPresetMap.get(listId) ?? "BULLETED",
            nestingLevel,
            listId,
          }
        : null,
      elements,
    };
  }

  private buildListPresetMap(
    lists: Record<string, docs_v1.Schema$List> | undefined,
  ): Map<string, "BULLETED" | "NUMBERED" | "CHECKBOX"> {
    const result = new Map<string, "BULLETED" | "NUMBERED" | "CHECKBOX">();

    for (const [listId, list] of Object.entries(lists ?? {})) {
      const firstLevel = list.listProperties?.nestingLevels?.[0];
      if (firstLevel?.glyphType) {
        result.set(listId, "NUMBERED");
        continue;
      }

      const glyph = `${firstLevel?.glyphSymbol ?? ""}${firstLevel?.glyphFormat ?? ""}`;
      if (glyph.includes("\u274f") || glyph.toLowerCase().includes("checkbox")) {
        result.set(listId, "CHECKBOX");
        continue;
      }

      result.set(listId, "BULLETED");
    }

    return result;
  }

  private getParagraphPlaceholderKind(
    element: docs_v1.Schema$ParagraphElement,
  ): "table" | "inlineObject" | "horizontalRule" | "other" {
    if (element.inlineObjectElement) {
      return "inlineObject";
    }
    if (element.horizontalRule) {
      return "horizontalRule";
    }
    return "other";
  }

  private normalizeParagraphElementStyle(
    element: docs_v1.Schema$ParagraphElement,
  ): NormalizedDocTextStyle | null {
    if (element.autoText?.textStyle) {
      return this.normalizeTextStyle(element.autoText.textStyle);
    }
    if (element.pageBreak?.textStyle) {
      return this.normalizeTextStyle(element.pageBreak.textStyle);
    }
    if (element.columnBreak?.textStyle) {
      return this.normalizeTextStyle(element.columnBreak.textStyle);
    }
    return null;
  }

  private normalizeTextStyle(
    style: docs_v1.Schema$TextStyle | undefined,
  ): NormalizedDocTextStyle | null {
    if (!style) {
      return null;
    }

    const link =
      style.link?.url ??
      style.link?.tabId ??
      undefined;

    const normalized: NormalizedDocTextStyle = {
      bold: style.bold ?? undefined,
      italic: style.italic ?? undefined,
      underline: style.underline ?? undefined,
      strikethrough: style.strikethrough ?? undefined,
      fontFamily: style.weightedFontFamily?.fontFamily ?? undefined,
      fontSize: style.fontSize?.magnitude ?? undefined,
      foregroundColor: this.optionalColorToHex(style.foregroundColor ?? undefined),
      backgroundColor: this.optionalColorToHex(style.backgroundColor ?? undefined),
      link,
    };

    return Object.values(normalized).some((value) => value !== undefined)
      ? normalized
      : null;
  }

  private optionalColorToHex(
    color: docs_v1.Schema$OptionalColor | undefined,
  ): string | undefined {
    if (!color) {
      return undefined;
    }
    const rgb = color.color?.rgbColor;
    if (!rgb) {
      return "transparent";
    }
    return DocsClient.rgbToHex(
      rgb.red ?? 0,
      rgb.green ?? 0,
      rgb.blue ?? 0,
    );
  }

  private buildTextStyleUpdate(
    options: TextStyleUpdateOptions,
  ): { textStyle: docs_v1.Schema$TextStyle; fields: string } {
    const textStyle: docs_v1.Schema$TextStyle = {};
    const fields: string[] = [];

    if (options.bold !== undefined) {
      textStyle.bold = options.bold;
      fields.push("bold");
    }
    if (options.italic !== undefined) {
      textStyle.italic = options.italic;
      fields.push("italic");
    }
    if (options.underline !== undefined) {
      textStyle.underline = options.underline;
      fields.push("underline");
    }
    if (options.strikethrough !== undefined) {
      textStyle.strikethrough = options.strikethrough;
      fields.push("strikethrough");
    }
    if (options.fontFamily !== undefined) {
      textStyle.weightedFontFamily = {
        fontFamily: options.fontFamily,
      };
      fields.push("weightedFontFamily.fontFamily");
    }
    if (options.fontSize !== undefined) {
      textStyle.fontSize = {
        magnitude: options.fontSize,
        unit: "PT",
      };
      fields.push("fontSize");
    }
    if (options.foregroundColor !== undefined) {
      textStyle.foregroundColor = DocsClient.hexToOptionalColor(
        options.foregroundColor,
      );
      fields.push("foregroundColor");
    }
    if (options.backgroundColor !== undefined) {
      textStyle.backgroundColor = DocsClient.hexToOptionalColor(
        options.backgroundColor,
      );
      fields.push("backgroundColor");
    }
    if (options.linkUrl !== undefined) {
      textStyle.link = options.linkUrl ? { url: options.linkUrl } : {};
      fields.push("link");
    }

    if (fields.length === 0) {
      throw new DriveAPIError("At least one text style field must be provided.");
    }

    return {
      textStyle,
      fields: fields.join(","),
    };
  }

  private buildParagraphStyleUpdate(
    options: ParagraphStyleUpdateOptions,
  ): { paragraphStyle: docs_v1.Schema$ParagraphStyle; fields: string } {
    const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
    const fields: string[] = [];

    if (options.namedStyleType !== undefined) {
      paragraphStyle.namedStyleType = options.namedStyleType;
      fields.push("namedStyleType");
    }
    if (options.alignment !== undefined) {
      paragraphStyle.alignment = options.alignment;
      fields.push("alignment");
    }

    if (fields.length === 0) {
      throw new DriveAPIError(
        "At least one paragraph style field must be provided.",
      );
    }

    return {
      paragraphStyle,
      fields: fields.join(","),
    };
  }

  private mapListPreset(
    preset: Exclude<DocListPreset, "REMOVE">,
  ): docs_v1.Schema$CreateParagraphBulletsRequest["bulletPreset"] {
    switch (preset) {
      case "CHECKBOX":
        return "BULLET_CHECKBOX";
      case "NUMBERED":
        return "NUMBERED_DECIMAL_ALPHA_ROMAN";
      case "BULLETED":
      default:
        return "BULLET_DISC_CIRCLE_SQUARE";
    }
  }

  static hexToOptionalColor(hex: string): docs_v1.Schema$OptionalColor {
    const value = hex.replace(/^#/, "");
    return {
      color: {
        rgbColor: {
          red: parseInt(value.slice(0, 2), 16) / 255,
          green: parseInt(value.slice(2, 4), 16) / 255,
          blue: parseInt(value.slice(4, 6), 16) / 255,
        },
      },
    };
  }

  static rgbToHex(red: number, green: number, blue: number): string {
    const toByte = (value: number) =>
      Math.max(0, Math.min(255, Math.round(value * 255)))
        .toString(16)
        .padStart(2, "0");
    return `#${toByte(red)}${toByte(green)}${toByte(blue)}`.toUpperCase();
  }
}
