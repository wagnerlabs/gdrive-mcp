import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import * as fs from "fs/promises";
import * as path from "path";

vi.mock("fs/promises");
vi.mock("@google-cloud/local-auth");
vi.mock("googleapis", () => ({
  google: {
    auth: {
      fromJSON: vi.fn(),
    },
  },
}));

describe("loadCredentials", () => {
  beforeEach(() => {
    vi.resetModules();
    delete process.env.GDRIVE_CREDENTIALS_PATH;
    delete process.env.GDRIVE_OAUTH_PATH;
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("reads from default path relative to project root, not cwd", async () => {
    const mockCreds = {
      type: "authorized_user",
      client_id: "id",
      client_secret: "secret",
      refresh_token: "token",
    };
    vi.mocked(fs.readFile).mockResolvedValue(JSON.stringify(mockCreds));

    const { google } = await import("googleapis");
    const mockClient = { credentials: {} };
    vi.mocked(google.auth.fromJSON).mockReturnValue(mockClient as any);

    const { loadCredentials } = await import("../src/auth.js");
    const result = await loadCredentials();

    const calledPath = vi.mocked(fs.readFile).mock.calls[0][0] as string;
    expect(path.isAbsolute(calledPath)).toBe(true);
    expect(calledPath).toMatch(/credentials[/\\].gdrive-server-credentials\.json$/);
    expect(result).toBe(mockClient);
  });

  it("reads from GDRIVE_CREDENTIALS_PATH env var", async () => {
    process.env.GDRIVE_CREDENTIALS_PATH = "/custom/creds.json";

    const mockCreds = {
      type: "authorized_user",
      client_id: "id",
      client_secret: "secret",
      refresh_token: "token",
    };
    vi.mocked(fs.readFile).mockResolvedValue(JSON.stringify(mockCreds));

    const { google } = await import("googleapis");
    const mockClient = { credentials: {} };
    vi.mocked(google.auth.fromJSON).mockReturnValue(mockClient as any);

    const { loadCredentials } = await import("../src/auth.js");
    const result = await loadCredentials();

    expect(fs.readFile).toHaveBeenCalledWith("/custom/creds.json", "utf-8");
    expect(result).toBe(mockClient);
  });

  it("throws a helpful error when credentials file is missing", async () => {
    vi.mocked(fs.readFile).mockRejectedValue(new Error("ENOENT"));

    const { loadCredentials } = await import("../src/auth.js");

    await expect(loadCredentials()).rejects.toThrow(/No saved credentials/);
    await expect(loadCredentials()).rejects.toThrow(/gdrive-mcp auth/);
  });
});
