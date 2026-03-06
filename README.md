# Google Drive MCP Server

A read-only [Model Context Protocol](https://modelcontextprotocol.io/) server that gives LLM-powered tools (Claude Code CLI, Cursor, Claude Desktop, etc.) access to your Google Drive.

Search, list, and read files — including automatic export of Google Docs (as Markdown), Sheets (as CSV), and Slides (as plain text). Works with both personal drives and shared drives.

## Quick start

### 1. Install

```bash
git clone https://github.com/wagnerlabs/gdrive-mcp.git
cd gdrive-mcp
npm install
npm run build
```

### 2. Set up Google Cloud credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (e.g. `gdrive-mcp`) or select an existing one
3. **Enable the Google Drive API**
   - Navigate to *APIs & Services > Library*
   - Search for "Google Drive API" and click **Enable**
4. **Configure the OAuth consent screen**
   - Navigate to *APIs & Services > OAuth consent screen* and click **Get started**
   - Enter an app name (e.g. `gdrive-mcp`), select your email as the support email, and click **Next**
   - **Audience**: select *Internal* (Workspace users) or *External* (personal Gmail), then click **Next**
   - **Contact Information**: enter your email and click **Next**
   - **Finish**: check the policy agreement box and click **Create**
5. **Add the Drive scope**
   - In the left sidebar, go to *Data Access*
   - Click **Add or remove scopes**
   - Add: `https://www.googleapis.com/auth/drive.readonly`
   - Save
6. **Create OAuth credentials**
   - In the left sidebar, go to *Clients* and click **Create Client**
   - Application type: **Desktop app**
   - Name: `gdrive-mcp` (this is just a console label to help you identify this client later)
   - Click **Create**
   - Download the JSON file and save it as:
     ```
     credentials/gcp-oauth.keys.json
     ```

### 3. Authenticate

```bash
node dist/index.js auth
```

A browser window will open for Google sign-in. After approval the token is saved to `credentials/.gdrive-server-credentials.json`.

### 4. Add to your MCP client

#### Claude Code CLI

```bash
claude mcp add --scope user wagnerlabs-gdrive -- node /absolute/path/to/gdrive-mcp/dist/index.js
```

The `--scope user` flag installs the server globally, so the MCP server will be available in Claude Code as **wagnerlabs-gdrive** from any directory you run Claude Code in.

To remove:

```bash
claude mcp remove wagnerlabs-gdrive
```

#### Cursor

Add to `.cursor/mcp.json` in any project (or globally):

```json
{
  "mcpServers": {
    "gdrive": {
      "command": "node",
      "args": ["/absolute/path/to/gdrive-mcp/dist/index.js"]
    }
  }
}
```

#### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "gdrive": {
      "command": "node",
      "args": ["/absolute/path/to/gdrive-mcp/dist/index.js"]
    }
  }
}
```

## Tools

All tools are read-only.

| Tool | Description |
|------|-------------|
| `gdrive_search` | Search files using full-text search or Drive query syntax |
| `gdrive_get_file` | Get detailed metadata for a file by ID |
| `gdrive_read_file` | Read file content (Docs → Markdown, Sheets → CSV, Slides → plain text) |
| `gdrive_list_files` | List files in a folder with sorting and pagination |

### File format handling

When reading files, Google Workspace documents are automatically exported:

| Source format | Exported as |
|---------------|-------------|
| Google Docs | Markdown |
| Google Sheets | CSV (first sheet) |
| Google Slides | Plain text |
| Google Drawings | PNG (metadata only) |
| Text files (`.txt`, `.json`, `.js`, etc.) | Read directly as UTF-8 |
| Binary files (images, PDFs, etc.) | Returns metadata with browser link |

## Configuration

Credential paths can be customized via environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `GDRIVE_OAUTH_PATH` | `credentials/gcp-oauth.keys.json` | Path to OAuth client secret |
| `GDRIVE_CREDENTIALS_PATH` | `credentials/.gdrive-server-credentials.json` | Path to saved token |

## Updating

After pulling new changes, rebuild and the MCP server will pick up the update on next launch — no need to re-register it:

```bash
cd /path/to/gdrive-mcp
git pull
npm install
npm run build
```

## Development

```bash
# Install dependencies
npm install

# Run in dev mode (uses tsx, no build step)
npm run dev

# Build
npm run build

# Run tests
npm test

# Run tests in watch mode
npm run test:watch
```

## License

MIT
