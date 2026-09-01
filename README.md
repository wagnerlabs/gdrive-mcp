# Google Drive MCP Server

A [Model Context Protocol](https://modelcontextprotocol.io/) server that gives LLM-powered tools (Claude Code CLI, Cursor, Claude Desktop, etc.) access to Google Drive, Google Docs, and Google Sheets.

Search, list, and read files, including automatic export of Google Docs as Markdown, Sheets as CSV, and Slides as plain text. Create and edit Google Docs with revision-protected text, rich content, native lists, and native tables. Create and edit Google Sheets with bounded range reads, value preconditions, formatting, tab management, and row or column operations. Works with both personal drives and shared drives.

## Quick start

### 1. Guided setup (recommended)

```bash
git clone https://github.com/wagnerlabs/gdrive-mcp.git
cd gdrive-mcp
./scripts/install.sh
```

The setup script installs dependencies, builds the project, and walks you through creating a Google Cloud project, enabling APIs, configuring OAuth, and authenticating. It prints ready-to-copy MCP client config at the end. Run with `--dry-run` to preview without side effects:

```bash
./scripts/install.sh --dry-run
```

### 2. Add to your MCP client

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
| `gdrive_get_sheet_values` | Read a bounded A1 range (up to 10,000 cells) for inspection or a later value precondition |
| `gdrive_get_document_info` | Get Google Docs metadata and optional tab-scoped structured content |
| `gdrive_get_document_content` | Read a Docs tab as paginated, revision-bound paragraphs and native table blocks |

### Write tools

#### Sheets

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

#### Docs

| Tool | Description | Destructive | Idempotent |
|------|-------------|:-----------:|:----------:|
| `gdrive_create_doc` | Create a blank Google Doc, optionally in a specific folder | No | No |
| `gdrive_insert_doc_content` | Insert styled paragraph blocks and nested native lists in one revision-controlled batch | No | No |
| `gdrive_batch_update_doc` | Apply up to 100 non-overlapping anchor-based text, style, paragraph, or list changes atomically | Yes | No |
| `gdrive_insert_doc_text` | Insert text at a position, explicit index, or text anchor | No | No |
| `gdrive_replace_doc_text` | Replace a targeted text range or anchored text match | Yes | No |
| `gdrive_replace_all_doc_text` | Replace every exact text match in a tab or across all tabs | Yes | Yes |
| `gdrive_delete_doc_text` | Delete a targeted text range or anchored text match | Yes | No |
| `gdrive_update_doc_text_style` | Apply character-level formatting such as bold, colors, fonts, and links | No | Yes |
| `gdrive_update_doc_paragraph_style` | Apply headings, alignment, indentation, and paragraph spacing | No | Yes |
| `gdrive_update_doc_list` | Create, continue, nest, change, or remove native list formatting | Yes | Yes |
| `gdrive_insert_doc_table` | Insert a native table and optionally populate its cells | No | No |
| `gdrive_update_doc_table_cells` | Replace multiple table cells with optional expected-text checks | Yes | No |
| `gdrive_modify_doc_table` | Insert/delete rows or columns, merge/unmerge cells, or delete a table | Yes | No |
| `gdrive_format_doc_table` | Format table cells, borders, columns, and rows | No | Yes |
| `gdrive_rename_doc` | Rename an existing Google Doc file | Yes | No |
| `gdrive_duplicate_doc` | Duplicate a Google Doc, optionally into a specific folder | No | No |

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

For full spreadsheet access, use `gdrive_get_spreadsheet_info` to discover tabs and `gdrive_get_sheet_values` to read an exact bounded range. The returned rectangular `values` array can be passed directly as `expected_current_values` to `gdrive_update_sheet`.

For structured Google Docs reads, use `gdrive_get_document_info` for compact metadata and tab discovery, then `gdrive_get_document_content` for bounded pages of paragraphs and native tables. Page tokens contain the document revision and fail with `STALE_PAGE_TOKEN` if a collaborator changes the document between pages. Structured paragraph responses include both raw `text` and `displayText` without the trailing paragraph newline, which is usually the safer anchor to round-trip into Docs write tools. `gdrive_read_file` remains the fastest way to read a Doc as Markdown.

Table blocks expose a one-based `tablePath` such as `[2]`; nested tables extend that path, such as `[2,1]`. Table cell coordinates are zero-based. Covered coordinates inside merged cells are rejected instead of being silently redirected to another cell.

## Safety model

The server applies several layers of protection for both Docs and Sheets, with Docs adding revision-aware write controls and anchor validation on top of the existing Sheets safeguards.

### 1. Per-tool MCP annotations

Each tool declares its safety characteristics via [MCP tool annotations](https://modelcontextprotocol.io/specification/2025-03-26/server/tools#annotations), so MCP clients can prompt the user appropriately before executing destructive operations. See the annotations in the tools tables above.

### 2. Read-before-write guard

The server tracks which spreadsheets and Docs the agent has actually looked at during the current session.

A spreadsheet is marked as "read" when the agent uses:

- `gdrive_read_file` (shows cell data as CSV)
- `gdrive_get_spreadsheet_info` (shows sheet structure and tabs)
- `gdrive_get_sheet_values` (shows the exact requested cell range)
- `gdrive_create_sheet` (the agent just created it, so it knows what's there)

A Google Doc is marked as "read" when the agent uses:

- `gdrive_read_file` (reads the Doc as Markdown and caches the current Docs revision when available)
- `gdrive_get_document_info` (reads tab metadata or structured paragraph content)
- `gdrive_get_document_content` (reads a revision-bound page of structured blocks)
- `gdrive_create_doc` (the agent just created it)

Every write tool checks this session state before executing. If the agent hasn't read the target resource, the call is rejected:

> *"You must read this spreadsheet before writing to it. Use gdrive_read_file, gdrive_get_spreadsheet_info, or gdrive_get_sheet_values first."*

> *"You must read this document before writing to it. Use gdrive_read_file, gdrive_get_document_info, or gdrive_get_document_content first."*

This is intentionally retained: it prevents blind writes and helps catch the wrong target. It is an authorization/awareness guard, not a concurrency guarantee. Preventing stale overwrites requires a revision/version check at write time, like the spreadsheet `expected_current_values` protection described below.

The read set resets when the server process restarts (every MCP session).

`gdrive_get_file` is deliberately excluded because it only returns Drive metadata, not sheet structure or document content.

### 3. Docs revision-aware writes

Docs edits are tied to the revision the agent most recently read explicitly:

- `conflict_mode: "strict"` is the default. The server checks the current revision immediately before the edit and also sends Docs `requiredRevisionId`, closing the race between that check and the write. A mismatch fails as `STALE_DOCUMENT` and requires a reread.
- `conflict_mode: "merge"` uses Docs `targetRevisionId`, which lets Google merge the edit with collaborator changes when possible

Internal metadata or anchor-resolution fetches do not silently advance the explicitly observed revision. This matters when a user edits a document after the agent reads it: an internal refresh may help resolve an anchor, but strict mode still refuses to write against the newer unseen revision.

The server also maintains a small session-scoped structured-content cache. Anchor-based tools reuse it when the document revision still matches; otherwise the server fetches a fresh structured snapshot before resolving anchors. Prefer text anchors, table paths, and cell coordinates over raw Docs indices. Explicit indices remain available as an advanced fallback and are validated against editable paragraph ranges where possible.

For targeted Docs text edits, you can also pass `expected_text` as an optimistic safety check. This verifies the exact text in the resolved range before the write is sent.

For anchor-based `gdrive_delete_doc_text` and `gdrive_replace_doc_text`, the server automatically trims only the final paragraph newline when a match reaches the end of the current tab, because the Docs API rejects delete ranges that include the segment-terminal newline. Explicit `start_index` / `end_index` edits stay strict and must exclude that trailing newline themselves.

### 4. Sheets precondition check

`gdrive_update_sheet` accepts an optional `expected_current_values` parameter — a 2D array the same shape as `values`. When provided, the server reads the current cell contents and compares them before writing. If they don't match, the write is refused with `STALE_SHEET_VALUES`, the range, and the actual values. Read the range again with `gdrive_get_sheet_values` before retrying.

- **For small, targeted edits** (changing one cell, fixing a formula): include `expected_current_values` as a safety net.
- **For bulk operations** (reformatting dates across 1,000 rows): skip it to avoid doubling API calls and hitting rate limits.

Set `include_previous_values: true` to include the old values in the response for auditing. When `expected_current_values` is provided, previous values are always included automatically.

### Recovery

Edits made via the Docs and Sheets APIs appear in Google Workspace version history, so users can revert changes if something goes wrong.

**No tool can delete an entire Google Doc or spreadsheet file from Google Drive.** Destructive operations are limited to structured edits inside a Doc and sheet-level operations inside a spreadsheet. Whole-file deletion still requires the Google Drive UI.

## Configuration

Credential paths can be customized via environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `GDRIVE_OAUTH_PATH` | `credentials/gcp-oauth.keys.json` | Path to OAuth client secret |
| `GDRIVE_CREDENTIALS_PATH` | `credentials/.gdrive-server-credentials.json` | Path to saved token |

## Upgrading

After pulling new changes, run the upgrade script. It rebuilds the project and walks you through any new setup steps (new APIs, scope changes, re-authentication) based on [`setup-manifest.json`](setup-manifest.json):

```bash
cd /path/to/gdrive-mcp
git pull
./scripts/upgrade.sh
```

If no setup changes are needed, `./scripts/upgrade.sh` just rebuilds and confirms you're up to date. The MCP server picks up changes on next launch — no need to re-register it.

> **Scope tradeoff:** This server requests full `drive` rather than `drive.file`. That is broader than Google's narrowest best practice, but it is required to preserve the current "read any accessible Drive file" behavior and to support rename, duplicate, and write operations on arbitrary existing Docs. The `documents` scope is required for structured Docs reads and Docs `batchUpdate` writes.

## Notes

- Spreadsheet creation (`gdrive_create_sheet`) places the new spreadsheet in the user's root Drive folder. Creating in a specific folder is not supported.
- `gdrive_read_file` exports spreadsheets as CSV from the first sheet only. Use `gdrive_get_spreadsheet_info` to discover all tabs.
- `gdrive_read_file` continues to export Google Docs as Markdown. Use `gdrive_get_document_info` when you need tab metadata, paragraph boundaries, headings, list state, or anchor-friendly ranges.
- For Docs formatting and paragraph structure changes, prefer `gdrive_get_document_content` so the agent has complete, paginated, revision-aware block and table data.
- `gdrive_replace_all_doc_text` defaults to the first tab for safety. To replace across every tab, you must set `all_tabs: true` explicitly.
- Numbered lists use native Google Docs list state. The server will continue a compatible preceding list when requested or fail clearly; it does not simulate numbering by inserting literal digits.
- Live visual inspection and screenshot-based post-write verification are deliberately deferred to a later release. This server verifies structure and revisions through the Docs API, but it does not claim pixel-level layout fidelity.

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

# Run the live Google Docs smoke test (requires saved credentials)
npm run test:live

# Run tests in watch mode
npm run test:watch
```

`npm run test:live` creates a temporary Google Doc, exercises revision-controlled text and structured reads through the MCP server flow, verifies the result, and then trashes the temporary file during cleanup. Set `RUN_LIVE_GOOGLE_TESTS=1` if invoking Vitest directly.

## License

MIT
