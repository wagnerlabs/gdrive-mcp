import { authenticate } from "@google-cloud/local-auth";
import { google } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import * as fs from "fs/promises";
import * as path from "path";
import { fileURLToPath } from "url";

const SCOPES = [
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
];

const PROJECT_ROOT = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

const DEFAULT_OAUTH_PATH = path.join(
  PROJECT_ROOT,
  "credentials",
  "gcp-oauth.keys.json",
);
const DEFAULT_CREDENTIALS_PATH = path.join(
  PROJECT_ROOT,
  "credentials",
  ".gdrive-server-credentials.json",
);

function oauthKeyfilePath(): string {
  return process.env.GDRIVE_OAUTH_PATH ?? DEFAULT_OAUTH_PATH;
}

function credentialsPath(): string {
  return process.env.GDRIVE_CREDENTIALS_PATH ?? DEFAULT_CREDENTIALS_PATH;
}

interface StoredCredentials {
  type: string;
  client_id: string;
  client_secret: string;
  refresh_token: string;
}

export async function saveCredentials(client: OAuth2Client): Promise<void> {
  const keyfile = JSON.parse(await fs.readFile(oauthKeyfilePath(), "utf-8"));
  const key = keyfile.installed ?? keyfile.web;
  const payload: StoredCredentials = {
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token!,
  };
  const filePath = credentialsPath();
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(payload, null, 2));
}

export async function loadCredentials(): Promise<OAuth2Client> {
  const filePath = credentialsPath();
  try {
    const content = await fs.readFile(filePath, "utf-8");
    const credentials: StoredCredentials = JSON.parse(content);
    const client = google.auth.fromJSON(credentials) as OAuth2Client;
    return client;
  } catch {
    throw new Error(
      `No saved credentials found at ${filePath}.\n` +
        `Run 'gdrive-mcp auth' to authenticate first.`,
    );
  }
}

export async function runAuthFlow(): Promise<OAuth2Client> {
  const keyfilePath = oauthKeyfilePath();
  try {
    await fs.access(keyfilePath);
  } catch {
    throw new Error(
      `OAuth client secret not found at ${keyfilePath}.\n\n` +
        "To set up Google Drive API access:\n" +
        "  1. Go to https://console.cloud.google.com/\n" +
        "  2. Create a project (or select an existing one)\n" +
        "  3. Enable the Google Drive API (APIs & Services > Library)\n" +
        "  4. Configure the OAuth consent screen (APIs & Services > OAuth consent screen)\n" +
        "  5. Create an OAuth 2.0 Client ID — type 'Desktop app' (APIs & Services > Credentials)\n" +
        `  6. Download the JSON and save it as:\n     ${keyfilePath}`,
    );
  }

  const client = await authenticate({
    scopes: SCOPES,
    keyfilePath,
  });

  if (!client.credentials?.refresh_token) {
    throw new Error(
      "No refresh token received from Google. This can happen when re-authorizing.\n" +
        "Revoke access at https://myaccount.google.com/permissions and try again.",
    );
  }
  await saveCredentials(client);

  return client;
}
