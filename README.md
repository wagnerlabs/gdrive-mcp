# Google Drive MCP Server

A [Model Context Protocol](https://modelcontextprotocol.io/) server that gives LLM-powered tools (Claude Code CLI, Cursor, Claude Desktop, etc.) access to your Google Drive and Google Sheets.

Search, list, and read files — including automatic export of Google Docs (as Markdown), Sheets (as CSV), and Slides (as plain text). Create new spreadsheets and edit existing ones: update cells, append rows, format ranges, manage tabs, and insert or delete rows and columns. Works with both personal drives and shared drives.

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
4. **Enable the Google Sheets API**
   - In *APIs & Services > Library*, search for "Google Sheets API" and click **Enable**
5. **Configure the OAuth consent screen**
   - Navigate to *APIs & Services > OAuth consent screen* and click **Get started**
   - Enter an app name (e.g. `gdrive-mcp`), select your email as the support email, and click **Next**
   - **Audience**: select *Internal* (Workspace users) or *External* (personal Gmail), then click **Next**
   - **Contact Information**: enter your email and click **Next**
   - **Finish**: check the policy agreement box and click **Create**
6. **Add the required scopes**
   - In the left sidebar, go to *Data Access*
   - Click **Add or remove scopes**
   - Add both:
     - `https://www.googleapis.com/auth/drive.readonly`
     - `https://www.googleapis.com/auth/spreadsheets`
   - Save
7. **Create OAuth credentials**
   - In the left sidebar, go to *Clients* and click **Create Client**
   - Application type: **Desktop app**
   - Name: `gdrive-mcp` (this is just a console label to help you identify this client later)
   - Click **Create**
   - Download the JSON file and save it to the `credentials/` folder at the root of this repo:
     ```
     gdrive-mcp/credentials/gcp-oauth.keys.json
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

### Read-only tools

| Tool | Description |
|------|-------------|
| `gdrive_search` | Search files using full-text search or Drive query syntax |
| `gdrive_get_file` | Get detailed metadata for a file by ID |
| `gdrive_read_file` | Read file content (Docs → Markdown, Sheets → CSV, Slides → plain text) |
| `gdrive_list_files` | List files in a folder with sorting and pagination |
| `gdrive_get_spreadsheet_info` | Get spreadsheet metadata including all sheet tabs and named ranges |

### Write tools

| Tool | Description | Destructive | Idempotent |
|------|-------------|:-----------:|:----------:|
| `gdrive_create_sheet` | Create a new spreadsheet | No | No |
| `gdrive_update_sheet` | Overwrite values in a cell range | Yes | Yes |
| `gdrive_append_sheet` | Append rows after existing data | No | No |
| `gdrive_clear_values` | Clear values from a cell range (preserves formatting) | Yes | Yes |
| `gdrive_format_cells` | Apply formatting to a cell range | No | Yes |
| `gdrive_add_sheet_tab` | Add a new sheet tab | No | No |
| `gdrive_delete_sheet_tab` | Delete a sheet tab and all its data | Yes | No |
| `gdrive_rename_sheet_tab` | Rename an existing sheet tab | Yes | No |
| `gdrive_insert_rows_columns` | Insert empty rows or columns | No | No |
| `gdrive_delete_rows_columns` | Delete rows or columns and their data | Yes | No |

### Value input options

When writing cell values (`gdrive_update_sheet`, `gdrive_append_sheet`), the `value_input_option` parameter controls how values are interpreted:

- **`USER_ENTERED`** (default) — Values are parsed as if typed into the Google Sheets UI. Formulas are executed (`=SUM(A1:A10)`), numbers and dates are formatted automatically.
- **`RAW`** — Values are stored exactly as provided. A string like `=SUM(A1:A10)` is stored as literal text, not executed as a formula.

### File format handling

When reading files with `gdrive_read_file`, Google Workspace documents are automatically exported:

| Source format | Exported as |
|---------------|-------------|
| Google Docs | Markdown |
| Google Sheets | CSV (first sheet only) |
| Google Slides | Plain text |
| Google Drawings | PNG (metadata only) |
| Text files (`.txt`, `.json`, `.js`, etc.) | Read directly as UTF-8 |
| Binary files (images, PDFs, etc.) | Returns metadata with browser link |

For full spreadsheet access (all tabs, structured data, editing), use `gdrive_get_spreadsheet_info` and the write tools instead of `gdrive_read_file`.

## Safety model

Three layers of protection ensure safe spreadsheet editing, none of which require configuration.

### 1. Per-tool MCP annotations

Each tool declares its safety characteristics via [MCP tool annotations](https://modelcontextprotocol.io/specification/2025-03-26/server/tools#annotations), so MCP clients can prompt the user appropriately before executing destructive operations. See the annotations in the tools table above.

### 2. Read-before-write guard

The server tracks which spreadsheets the agent has actually looked at during the current session. A spreadsheet is marked as "read" when the agent uses:

- `gdrive_read_file` (shows cell data as CSV)
- `gdrive_get_spreadsheet_info` (shows sheet structure and tabs)
- `gdrive_create_sheet` (the agent just created it, so it knows what's there)

Every write tool checks this set before executing. If the agent hasn't read the target spreadsheet, the call is rejected:

> *"You must read this spreadsheet before writing to it. Use gdrive_read_file or gdrive_get_spreadsheet_info first."*

This prevents the agent from accidentally targeting the wrong spreadsheet. The set resets when the server process restarts (every MCP session).

`gdrive_get_file` is deliberately excluded — it only returns basic Drive metadata and does not show the agent anything about the sheet's content or structure.

### 3. Optional precondition check

`gdrive_update_sheet` accepts an optional `expected_current_values` parameter — a 2D array the same shape as `values`. When provided, the server reads the current cell contents and compares them before writing. If they don't match, the write is refused with an error showing what the cells actually contain.

- **For small, targeted edits** (changing one cell, fixing a formula): include `expected_current_values` as a safety net.
- **For bulk operations** (reformatting dates across 1,000 rows): skip it to avoid doubling API calls and hitting rate limits.

Set `include_previous_values: true` to include the old values in the response for auditing. When `expected_current_values` is provided, previous values are always included automatically.

### Recovery

Edits made via the Sheets API appear in Google Sheets' version history (*File > Version history*), so users can revert changes if something goes wrong.

**No tool can delete an entire spreadsheet file from Google Drive.** Deletion is limited to individual sheet tabs within a spreadsheet (`gdrive_delete_sheet_tab`). The spreadsheet file itself can only be deleted through the Google Drive UI.

## Configuration

Credential paths can be customized via environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `GDRIVE_OAUTH_PATH` | `credentials/gcp-oauth.keys.json` | Path to OAuth client secret |
| `GDRIVE_CREDENTIALS_PATH` | `credentials/.gdrive-server-credentials.json` | Path to saved token |

## Upgrading from read-only

If you previously used this server before write support was added, your saved token only has the `drive.readonly` scope and won't be able to edit spreadsheets. To grant the new `spreadsheets` scope:

```bash
rm credentials/.gdrive-server-credentials.json
node dist/index.js auth
```

This opens a new browser consent screen that includes the Sheets scope. Existing read-only functionality is unaffected.

> **Note:** The `spreadsheets` scope grants read/write access to all spreadsheets the signed-in user can access. This is necessary for editing existing sheets. The `drive.readonly` scope remains for reading non-spreadsheet files.

## Updating

After pulling new changes, rebuild and the MCP server will pick up the update on next launch — no need to re-register it:

```bash
cd /path/to/gdrive-mcp
git pull
npm install
npm run build
```

## Notes

- Spreadsheet creation (`gdrive_create_sheet`) places the new spreadsheet in the user's root Drive folder. Creating in a specific folder is not supported.
- `gdrive_read_file` exports spreadsheets as CSV from the first sheet only. Use `gdrive_get_spreadsheet_info` to discover all tabs.

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
