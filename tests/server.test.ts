import { describe, it, expect, vi } from "vitest";
import { createServer } from "../src/server.js";
import { DriveClient } from "../src/client.js";

function makeMockClient(): DriveClient {
  return {
    search: vi.fn(),
    getFile: vi.fn(),
    readFile: vi.fn(),
    listFiles: vi.fn(),
  } as unknown as DriveClient;
}

function getTools(server: ReturnType<typeof createServer>) {
  return (server as any)._registeredTools as Record<string, any>;
}

describe("createServer", () => {
  it("registers all four tools", () => {
    const tools = getTools(createServer(makeMockClient()));
    const names = Object.keys(tools);

    expect(names).toContain("gdrive_search");
    expect(names).toContain("gdrive_get_file");
    expect(names).toContain("gdrive_read_file");
    expect(names).toContain("gdrive_list_files");
    expect(names).toHaveLength(4);
  });

  it("does not register any write tools", () => {
    const tools = getTools(createServer(makeMockClient()));
    const names = Object.keys(tools);

    const writeKeywords = ["create", "update", "delete", "move", "copy", "trash", "rename"];
    for (const name of names) {
      for (const keyword of writeKeywords) {
        expect(name).not.toContain(keyword);
      }
    }
  });

  it("all tools have readOnlyHint annotation", () => {
    const tools = getTools(createServer(makeMockClient()));

    for (const [name, tool] of Object.entries(tools)) {
      const annotations = tool.annotations;
      expect(annotations?.readOnlyHint, `${name} should be readOnly`).toBe(true);
      expect(annotations?.destructiveHint, `${name} should not be destructive`).toBe(false);
    }
  });
});
