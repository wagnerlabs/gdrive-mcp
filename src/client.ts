import { google, drive_v3 } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import { GaxiosError } from "gaxios";

const FILE_FIELDS =
  "id, name, mimeType, size, modifiedTime, createdTime, owners, parents, webViewLink, description, starred, trashed";

const LIST_FIELDS = `nextPageToken, files(${FILE_FIELDS})`;

const GOOGLE_EXPORT_MAP: Record<string, string> = {
  "application/vnd.google-apps.document": "text/markdown",
  "application/vnd.google-apps.spreadsheet": "text/csv",
  "application/vnd.google-apps.presentation": "text/plain",
  "application/vnd.google-apps.drawing": "image/png",
};

const TEXT_MIME_PREFIXES = [
  "text/",
  "application/json",
  "application/xml",
  "application/javascript",
  "application/typescript",
  "application/x-yaml",
  "application/toml",
  "application/x-sh",
];

const ERROR_MESSAGES: Record<number, string> = {
  400: "Bad request — check your query parameters.",
  401: "Authentication expired. Re-run 'gdrive-mcp auth'.",
  403: "Insufficient permissions — check your OAuth scopes.",
  404: "Not found — the requested file or folder does not exist.",
  429: "Rate limit exceeded — try again in a moment.",
  500: "Google Drive server error — try again later.",
  503: "Google Drive service temporarily unavailable — try again later.",
};

export class DriveAPIError extends Error {
  statusCode: number | undefined;
  constructor(message: string, statusCode?: number) {
    super(message);
    this.name = "DriveAPIError";
    this.statusCode = statusCode;
  }
}

export interface FileSummary {
  id: string;
  name: string;
  mimeType: string;
  size?: string;
  modifiedTime?: string;
  createdTime?: string;
  webViewLink?: string;
  parents?: string[];
  owners?: Array<{ displayName: string; emailAddress: string }>;
  description?: string;
  starred?: boolean;
  trashed?: boolean;
}

function formatFile(file: drive_v3.Schema$File): FileSummary {
  return {
    id: file.id!,
    name: file.name!,
    mimeType: file.mimeType!,
    size: file.size ?? undefined,
    modifiedTime: file.modifiedTime ?? undefined,
    createdTime: file.createdTime ?? undefined,
    webViewLink: file.webViewLink ?? undefined,
    parents: file.parents ?? undefined,
    owners: file.owners?.map((o) => ({
      displayName: o.displayName!,
      emailAddress: o.emailAddress!,
    })),
    description: file.description ?? undefined,
    starred: file.starred ?? undefined,
    trashed: file.trashed ?? undefined,
  };
}

function isTextMime(mimeType: string): boolean {
  return TEXT_MIME_PREFIXES.some((prefix) => mimeType.startsWith(prefix));
}

function handleApiError(err: unknown): never {
  if (err instanceof GaxiosError) {
    const status = err.response?.status;
    const message =
      status && ERROR_MESSAGES[status]
        ? ERROR_MESSAGES[status]
        : `Google Drive API error: ${err.message}`;
    throw new DriveAPIError(message, status);
  }
  throw err;
}

export class DriveClient {
  private drive: drive_v3.Drive;

  constructor(auth: OAuth2Client) {
    this.drive = google.drive({ version: "v3", auth });
  }

  async search(
    query: string,
    maxResults: number = 20,
    pageToken?: string,
  ): Promise<{ files: FileSummary[]; nextPageToken?: string }> {
    try {
      const res = await this.drive.files.list({
        q: query,
        pageSize: maxResults,
        pageToken: pageToken,
        fields: LIST_FIELDS,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      });
      return {
        files: (res.data.files ?? []).map(formatFile),
        nextPageToken: res.data.nextPageToken ?? undefined,
      };
    } catch (err) {
      handleApiError(err);
    }
  }

  async getFile(fileId: string): Promise<FileSummary> {
    try {
      const res = await this.drive.files.get({
        fileId,
        fields: FILE_FIELDS,
        supportsAllDrives: true,
      });
      return formatFile(res.data);
    } catch (err) {
      handleApiError(err);
    }
  }

  async readFile(
    fileId: string,
    maxChars: number = 100_000,
  ): Promise<{ content: string; mimeType: string; truncated: boolean }> {
    const meta = await this.getFile(fileId);
    const exportMime = GOOGLE_EXPORT_MAP[meta.mimeType];

    if (exportMime) {
      return this.exportGoogleFile(fileId, exportMime, maxChars, meta.mimeType);
    }

    if (isTextMime(meta.mimeType)) {
      return this.downloadTextFile(fileId, meta.mimeType, maxChars);
    }

    throw new DriveAPIError(
      `Cannot read binary file (${meta.mimeType}). ` +
        `Use gdrive_get_file for metadata, or open in browser: ${meta.webViewLink ?? "N/A"}`,
    );
  }

  async listFiles(
    folderId: string = "root",
    maxResults: number = 50,
    pageToken?: string,
    orderBy: string = "folder,modifiedTime desc",
  ): Promise<{ files: FileSummary[]; nextPageToken?: string }> {
    try {
      const res = await this.drive.files.list({
        q: `'${folderId}' in parents and trashed = false`,
        pageSize: maxResults,
        pageToken: pageToken,
        orderBy,
        fields: LIST_FIELDS,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      });
      return {
        files: (res.data.files ?? []).map(formatFile),
        nextPageToken: res.data.nextPageToken ?? undefined,
      };
    } catch (err) {
      handleApiError(err);
    }
  }

  private async exportGoogleFile(
    fileId: string,
    exportMime: string,
    maxChars: number,
    sourceMime: string,
  ): Promise<{ content: string; mimeType: string; truncated: boolean }> {
    try {
      const res = await this.drive.files.export({
        fileId,
        mimeType: exportMime,
      });

      if (exportMime === "image/png") {
        return {
          content: `[Drawing exported as PNG — binary content not shown. Use gdrive_get_file for metadata.]`,
          mimeType: exportMime,
          truncated: false,
        };
      }

      let text = typeof res.data === "string" ? res.data : String(res.data);
      let truncated = false;

      if (sourceMime === "application/vnd.google-apps.spreadsheet") {
        text = `[Note: CSV export contains the first sheet only.]\n\n${text}`;
      }

      if (text.length > maxChars) {
        text = text.slice(0, maxChars) + "\n[truncated]";
        truncated = true;
      }

      return { content: text, mimeType: exportMime, truncated };
    } catch (err) {
      if (
        err instanceof GaxiosError &&
        err.message?.includes("too large for export")
      ) {
        throw new DriveAPIError(
          "File too large for export (Google limits exports to 10 MB). " +
            "Use gdrive_get_file for metadata and open in browser.",
        );
      }
      handleApiError(err);
    }
  }

  private async downloadTextFile(
    fileId: string,
    mimeType: string,
    maxChars: number,
  ): Promise<{ content: string; mimeType: string; truncated: boolean }> {
    try {
      const res = await this.drive.files.get(
        { fileId, alt: "media", supportsAllDrives: true },
        { responseType: "text" },
      );

      let text = typeof res.data === "string" ? res.data : String(res.data);
      let truncated = false;

      if (text.length > maxChars) {
        text = text.slice(0, maxChars) + "\n[truncated]";
        truncated = true;
      }

      return { content: text, mimeType, truncated };
    } catch (err) {
      handleApiError(err);
    }
  }
}
