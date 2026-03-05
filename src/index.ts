#!/usr/bin/env node

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { runAuthFlow, loadCredentials } from "./auth.js";
import { DriveClient } from "./client.js";
import { createServer } from "./server.js";

async function main(): Promise<void> {
  const args = process.argv.slice(2);

  if (args[0] === "auth") {
    try {
      await runAuthFlow();
      console.log("Authentication successful. Credentials saved.");
    } catch (err) {
      console.error(err instanceof Error ? err.message : String(err));
      process.exit(1);
    }
    return;
  }

  let auth;
  try {
    auth = await loadCredentials();
  } catch (err) {
    console.error(err instanceof Error ? err.message : String(err));
    process.exit(1);
  }

  const client = new DriveClient(auth);
  const server = createServer(client);
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
