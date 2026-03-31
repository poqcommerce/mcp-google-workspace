# Google Workspace MCP Server

An MCP server that gives Claude (or any MCP-compatible AI) full read/write access to Google Sheets, Docs, Drive, and Slides. 56 tools, batch operations throughout, and a template workflow for branded presentations.

**Version:** 2.3.0 | **Last Updated:** 2026-03-30 | **Tools:** 56

---

## Setup

There are two parts: a **one-time Google Cloud setup** (needs someone with Google Workspace admin access), and **per-user setup** (done by each person who wants to use the MCP). If your org has already completed Part A, skip straight to Part B — you just need the Client ID and Client Secret from whoever set it up.

### Part A: Google Cloud project (once per org)

This creates the "app" that lets Claude talk to Google Workspace. It only needs to be done once — after that, everyone on the team uses the same Client ID and Client Secret.

**Who does this?** Someone who can create projects in your organisation's [Google Cloud Console](https://console.cloud.google.com/). This is usually an IT admin or whoever manages your Google Workspace. If you're not sure, ask your IT team to follow these steps.

1. Go to [console.cloud.google.com](https://console.cloud.google.com/) and sign in with your work Google account
2. Click the project dropdown (top-left) and select **New Project**. Name it something like "Claude MCP" and click **Create**
3. Make sure the new project is selected, then go to **APIs & Services > Library** (left sidebar)
4. Search for and **enable** each of these five APIs (click each one, then click "Enable"):
   - **Google Drive API**
   - **Google Docs API**
   - **Google Sheets API**
   - **Google Slides API**
   - **Drive Activity API** (required for suggestion attribution — who made which tracked change)
5. Go to **APIs & Services > Credentials** (left sidebar)
6. Click **+ Create Credentials > OAuth client ID**
   - If prompted to configure the OAuth consent screen first, choose **Internal** (restricts to your org's Google accounts) or **External** (for testing with personal accounts), fill in the required fields (app name, support email), and save
   - For Application type, select **Desktop app**
   - Give it a name (e.g. "Claude MCP") and click **Create**
7. You'll see a **Client ID** and **Client Secret**. These are shared across your team — every user needs the same pair. Distribute them securely (e.g. via a password manager or private internal channel, not a public wiki). They authenticate the app itself and should be kept confidential to your organisation

### Part B: Per-user setup

You need two things before starting:
- The **Client ID** and **Client Secret** from Part A (ask whoever set up the Google Cloud project)
- [Node.js](https://nodejs.org/) v18 or later installed on your machine

#### 1. Install

Open a terminal and run:

```bash
git clone https://github.com/poqcommerce/mcp-google-workspace
cd mcp-google-workspace
npm install
npm run build
```

#### 2. Create your credentials file

In the `mcp-google-workspace` folder, create a file called `.env` and paste in the Client ID and Client Secret you were given:

```bash
GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your-client-secret
GOOGLE_REFRESH_TOKEN=
```

Leave `GOOGLE_REFRESH_TOKEN` blank for now — the next step fills it in.

#### 3. Authenticate with your Google account

**macOS/Linux:**
```bash
npm run auth
```

**Windows (PowerShell):**
```powershell
npm run auth
```

This opens your browser and asks you to sign in with **your own** Google account and grant permission. Each user authenticates individually — the Client ID/Secret identify the app, but this step links it to *your* Google account and files.

After you approve, it prints a **refresh token** in the terminal. Copy the token and paste it into your `.env` file after `GOOGLE_REFRESH_TOKEN=`.

> **Note:** The refresh token is personal to you. Don't share it — it grants access to your Google Drive, Docs, Sheets, and Slides.

#### 4. Connect to Claude

The easiest way is to run the setup script — it reads your `.env` file and writes the Claude Desktop config automatically:

```bash
npm run setup
```

This detects your platform, finds the config file, merges with any existing MCP servers, and writes valid JSON. No manual editing needed.

**If you prefer to do it manually**, add the MCP server to your Claude config file:

| Platform | Config file location |
|----------|---------------------|
| Claude Code (CLI) | `~/.claude/mcp.json` |
| Claude Desktop (macOS) | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| Claude Desktop (Windows) | `%APPDATA%\Claude\claude_desktop_config.json` |

Add this to the file (create it if it doesn't exist), replacing the paths and credentials with your own:

**macOS/Linux:**
```json
{
  "mcpServers": {
    "google-workspace": {
      "command": "/full/path/to/mcp-google-workspace/start-mcp.sh",
      "env": {
        "GOOGLE_CLIENT_ID": "your-client-id",
        "GOOGLE_CLIENT_SECRET": "your-client-secret",
        "GOOGLE_REFRESH_TOKEN": "your-personal-refresh-token"
      }
    }
  }
}
```

**Windows:**
```json
{
  "mcpServers": {
    "google-workspace": {
      "command": "C:/Users/yourname/mcp-google-workspace/start-mcp.cmd",
      "env": {
        "GOOGLE_CLIENT_ID": "your-client-id",
        "GOOGLE_CLIENT_SECRET": "your-client-secret",
        "GOOGLE_REFRESH_TOKEN": "your-personal-refresh-token"
      }
    }
  }
}
```

> **Tips:**
> - Use full absolute paths. Relative paths won't work.
> - On Windows, use forward slashes (`C:/Users/...`) in JSON — they work fine and avoid escaping issues.
> - The `%APPDATA%\Claude` folder may be hidden. Paste `%APPDATA%\Claude` into the File Explorer address bar to navigate there directly.

#### 5. Verify it works

Restart Claude, then try asking: "Search my Google Drive for recent documents." If Claude lists your files, you're connected.

---

## Tools Reference

### Google Sheets — 12 tools

| Tool | Description |
|------|-------------|
| `gsheets_create_and_populate` | Create a spreadsheet with data in one call |
| `gsheets_read_data` | Read data from a range |
| `gsheets_batch_update` | Update multiple ranges in a single API call |
| `gsheets_append_rows` | Add rows to the end of a sheet |
| `gsheets_format_cells` | Apply styling (bold, colours, borders, number formats) |
| `gsheets_get_spreadsheet_info` | Get sheet names, row/column counts, metadata |
| `gsheets_add_sheet` | Add a new sheet tab |
| `gsheets_delete_sheet` | Remove a sheet tab |
| `gsheets_rename_sheet` | Rename a sheet tab |
| `gsheets_duplicate_sheet` | Copy a sheet within or between spreadsheets |
| `gsheets_insert_delete_dimensions` | Insert or delete rows/columns |
| `gsheets_sort_range` | Sort data by one or more columns |

### Google Docs — 7 tools

| Tool | Description |
|------|-------------|
| `gdocs_create_document` | Create a new document (optionally in a folder with initial content) |
| `gdocs_get_document` | Read document content including tables (rendered as markdown) with tracked changes detected and section context |
| `gdocs_insert_text` | Insert text at a specific index |
| `gdocs_append_text` | Append text to the end |
| `gdocs_replace_text` | Find and replace throughout the document |
| `gdocs_format_text` | Apply bold, italic, underline, font size |
| `gdocs_set_heading` | Convert text to heading (H1–H6) |

### Google Drive — 19 tools

| Tool | Description |
|------|-------------|
| `gdrive_search` | Search by name, type, folder, or combined queries |
| `gdrive_search_content` | Full-text search within file contents |
| `gdrive_read_file` | Read file contents (Docs, Sheets, text files) |
| `gdrive_get_file_info` | Get file metadata (size, dates, parents) |
| `gdrive_copy_file` | Copy any file (preserves themes/layouts for Slides) |
| `gdrive_move_file` | Move a file to a different folder |
| `gdrive_batch_move` | Move multiple files at once |
| `gdrive_create_folder` | Create a new folder |
| `gdrive_copy_folder` | Recursively copy a folder and all contents |
| `gdrive_list_folder_tree` | List all files recursively with metadata |
| `gdrive_get_revisions` | Get version history |
| `gdrive_export_file` | Export to PDF, DOCX, XLSX, PPTX |
| `gdrive_batch_export` | Bulk export multiple files |
| `gdrive_list_permissions` | See who has access to a file/folder |
| `gdrive_list_comments` | List all comments with replies, resolved status, and anchored text |
| `gdrive_create_comment` | Add a comment to any file (Docs, Sheets, Slides) |
| `gdrive_reply_to_comment` | Reply to a comment, or resolve/reopen it |
| `gdrive_delete_comment` | Delete a comment (author only) |
| `gdrive_suggestion_activity` | Get who made suggestions and when (for redline attribution) |

### Google Slides — 17 tools

| Tool | Description |
|------|-------------|
| `gslides_create_presentation` | Create a new presentation |
| `gslides_get_presentation` | Get full structure: slides, elements, text, notes |
| `gslides_add_slide` | Add a slide with a predefined layout |
| `gslides_delete_slide` | Remove a slide |
| `gslides_insert_text` | Insert text into a shape or placeholder |
| `gslides_replace_text` | Find and replace across slides |
| `gslides_speaker_notes` | Read or write speaker notes |
| `gslides_add_shape` | Add rectangles, ellipses, text boxes, etc. |
| `gslides_add_image` | Insert an image from a URL |
| `gslides_add_table` | Create a table on a slide |
| `gslides_insert_table_text` | Populate a specific table cell |
| `gslides_format_text` | Format text in shapes or individual table cells |
| `gslides_batch_format_text` | Format many shapes/cells in a single API call |
| `gslides_format_table` | Set column widths, row heights, alignment, header colours |
| `gslides_create_bullets` | Convert text to bulleted or numbered lists |
| `gslides_update_page_properties` | Set slide background colour |
| `gslides_get_thumbnail` | Get a PNG thumbnail URL for a slide |

### Authentication — 2 tools

| Tool | Description |
|------|-------------|
| `gsheets_get_auth_url` | Generate the OAuth consent URL |
| `gsheets_set_auth_code` | Exchange auth code for a refresh token |

---

## Best Practices

### Use batch operations

The single biggest performance lever. Instead of 30 individual `gslides_format_text` calls, use `gslides_batch_format_text` with an array of operations — one API call instead of 30.

Batch tools available:
- **Sheets:** `gsheets_batch_update` (multi-range writes)
- **Drive:** `gdrive_batch_move`, `gdrive_batch_export`
- **Slides:** `gslides_batch_format_text` (text styling), `gslides_format_table` (column widths, alignment, backgrounds)

### Work with IDs, not names

Google Workspace uses file/object IDs everywhere. When creating files, capture the returned ID for subsequent operations. When working with slides, use `gslides_get_presentation` to discover element object IDs before trying to modify them.

### Branded presentations via template copy

The most reliable way to create branded presentations:

1. `gdrive_copy_file` — copy a branded template (preserves theme, master slides, layouts, logos)
2. `gslides_delete_slide` — remove template placeholder slides
3. `gslides_add_slide` — add slides with branded layouts
4. Populate and format content

This is better than trying to set colours/fonts from scratch — the copy inherits the full theme.

### Table creation workflow

For well-formatted tables:

1. `gslides_add_table` — create with position and overall dimensions
2. `gslides_insert_table_text` — populate cells
3. `gslides_format_table` — set column widths proportional to content (e.g. 25/75 for label/description), alignment, header background
4. `gslides_batch_format_text` — apply font size, colour, bold to all cells in one call

Key tips:
- Set explicit column widths — don't rely on equal-width defaults
- Choose font size based on row count (10–11pt for <12 rows, 9pt for 13–18 rows)
- Use `contentAlignment: "TOP"` for multi-line cells
- Header rows work well with a contrasting background + white bold text, 2–3pt larger than body

### Re-authenticate after scope changes

If you add a new Google API (e.g. enabling Slides for the first time), run `npm run auth` again to grant the new scope. The OAuth scopes are: `spreadsheets`, `drive`, `documents`, `presentations`, `drive.activity.readonly`.

### Style consistency

When reading documents (`gdocs_get_document`) or presentations (`gslides_get_presentation`), the MCP automatically detects the dominant text style — font family, font size, colour, bold/italic — and includes it in the response. This means Claude can see what formatting already exists and match it when adding new content.

For example, when you say "add a section to this doc", Claude will see:

```
Dominant body style: {"fontFamily":"Roboto","fontSize":11,"foregroundColor":{"red":0.2,"green":0.2,"blue":0.2}}
Heading style: {"fontFamily":"Roboto","fontSize":16,"bold":true}
```

...and use those same properties when formatting the new text. No manual specification needed.

---

## Style Guides

Style guides tell Claude *how* to format content — brand colours, font choices, table layout rules, slide conventions. They work alongside the automatic style detection: detection handles "match what's already there", while style guides handle "here's what it *should* look like".

### How Claude reads style guides

Claude reads instructions from `CLAUDE.md` files automatically at the start of every session. The key question is **where** you put the file, because that determines who sees it and when.

| Location | Scope | Shared via git? | Use case |
|----------|-------|-----------------|----------|
| `CLAUDE.md` in a repo root | Anyone working in that repo | Yes | Project-specific conventions |
| `~/.claude/CLAUDE.md` | One user, all their projects | No | Personal preferences |
| `~/.claude/projects/<path>/CLAUDE.md` | One user, one project | No | Personal project overrides |

All three levels are loaded together when they apply. If there's a conflict, be explicit about priority in the file itself.

#### In this repo (MCP tool developers)

If you're working *inside* this repo (developing the MCP tools themselves), add a `CLAUDE.md` to the repo root. It will apply to everyone who clones it.

#### In other projects (MCP tool users)

This is the more common case: you're working in a different repo (e.g. a board reports project, a data pipeline) and want Claude to follow your style guide when creating Google Workspace content.

**Option 1: User-level (recommended starting point)**

Add your style rules to `~/.claude/CLAUDE.md`. They'll apply to every Claude session you run, regardless of which project you're in. This is the simplest way to get consistent formatting across all your work.

```markdown
# ~/.claude/CLAUDE.md

## Google Workspace Style Guide
- Presentations: copy branded template (ID: your-template-id) before creating slides
- Brand colours: dark purple rgb(0.106, 0.078, 0.392), white backgrounds
- Tables: 25/75 column split, 9pt for dense tables, dark purple header with white text
- Documents: use existing font/size (Claude detects this automatically)
- Always use British English spelling
```

**Option 2: Team-wide (share via a dotfiles repo or onboarding doc)**

Since `~/.claude/CLAUDE.md` isn't shared via git, the best way to standardise across a team is to publish the style guide content somewhere accessible (e.g. a shared doc, internal wiki, or dotfiles repo) and have each team member copy it into their own `~/.claude/CLAUDE.md`.

Include a ready-to-paste block in your onboarding docs:

```
To set up Claude with our brand guidelines, create or edit ~/.claude/CLAUDE.md
and paste the following:

[your style guide content here]
```

**Option 3: Per-project**

If different projects need different styles, add a `CLAUDE.md` to each project repo. This is version-controlled and shared with the team via git, but only applies when someone is working in that specific repo.

### What to include in a style guide

A good style guide for MCP tools covers:

- **Brand colours** as RGB values (0–1 range for Google APIs)
- **Font choices** and sizes for different contexts (body, headings, tables)
- **Table layout rules** — column width ratios, alignment, header styling
- **Template IDs** — which branded template to copy for new presentations
- **Naming conventions** — file/folder naming patterns
- **Workflow preferences** — e.g. "always create documents in the Projects folder"

Keep it concise. Claude reads the entire file at session start, so shorter guides are more likely to be followed consistently. See [`docs/example-style-guide.md`](docs/example-style-guide.md) for a full example.

---

## Project Structure

```
mcp-google-workspace/
├── src/
│   ├── index.ts          # MCP server, handler dispatch
│   ├── types.ts          # Shared interfaces
│   ├── utils.ts          # Helpers (a1ToGridRange, response builders)
│   ├── auth.ts           # Standalone OAuth flow (npm run auth)
│   └── tools/
│       ├── sheets.ts     # 12 tools
│       ├── docs.ts       # 7 tools
│       ├── drive.ts      # 19 tools
│       ├── slides.ts     # 17 tools
│       └── auth.ts       # 2 tools
├── dist/                 # Compiled JS
├── start-mcp.sh          # Launch script (macOS/Linux)
├── start-mcp.cmd         # Launch script (Windows)
├── install-windows.bat   # One-click Windows installer
├── .env                  # Credentials (not committed)
├── package.json
└── tsconfig.json
```

Each `tools/*.ts` exports a `get*ToolDefinitions()` function and a handler class. The server iterates handlers in order — first non-null response wins.

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Insufficient permissions" | Re-run `npm run auth` to refresh token with current scopes |
| "File not found" | Verify file ID; check the file is accessible to your Google account |
| "Rate limit exceeded" | Use batch operations; reduce request frequency |
| OAuth redirect fails | Ensure port 3000 is free (`lsof -i :3000`) |
| MCP tools not appearing | Check `start-mcp.sh` (or `start-mcp.cmd` on Windows) path in config; run `npm run build` |
| Config JSON parse error (Windows) | Run `npm run setup` to regenerate the config. If editing manually, use forward slashes in paths (`C:/Users/...`) and ensure no trailing backslashes |
| "Bad control character" in config | The file has invisible characters (BOM or line breaks). Re-run `npm run setup` or re-save in Notepad as UTF-8 (not UTF-8 with BOM) |
| Slides layout not found | Use `gslides_get_presentation` to check available layouts in the theme |

---

## Development

```bash
npm run build     # Compile TypeScript
npm run dev       # Watch mode
npm run auth      # OAuth flow
npm run setup     # Write Claude Desktop config from .env
npm start         # Start server directly
```

---

## Version History

### v2.3.0 — 2026-03-30
- `gdrive_suggestion_activity` — new tool to surface who made suggestions and when, using the Drive Activity API. Useful for attributing redline changes in contract negotiations. Note: only captures suggestions made natively in Google Docs, not tracked changes imported from Word
- `gdocs_get_document` now includes section heading context on each suggestion (e.g. `[Service Credits] DELETE: "sole and exclusive"`)
- Removed dead author-extraction code from suggestion parsing
- Added `drive.activity.readonly` OAuth scope (requires re-auth: `npm run auth`)
- Fixed dotenv/MCP env var conflict — `.env` now loads with `override: true` to prevent stale cached tokens
- 56 total tools (was 55)

### v2.2.0 — 2026-03-23
- `gdocs_get_document` now renders tables as pipe-delimited markdown — previously table content was silently dropped
- Suggestion/redline detection now works inside table cells (insertions, deletions, format changes)
- Style detection includes text inside tables
- Added test suite (vitest) with table rendering tests

### v2.1.0 — 2026-03-19
- Added comment tools: list, create, reply/resolve, delete (works on Docs, Sheets, Slides)
- `gdocs_get_document` now detects and surfaces suggested changes (tracked changes/redlines)
- Added `dotenv` — `.env` files load natively on all platforms
- Added Windows support: `install-windows.bat` (one-click installer), `start-mcp.cmd`, `npm run setup`
- 55 total tools (was 51)

### v2.0.0 — 2026-03-09
- Added Google Slides support (17 tools)
- Added `gdrive_copy_file` for template workflows
- Added batch text formatting for Slides
- Added table structure formatting (column widths, row heights, alignment, backgrounds)
- 51 total tools (was 27)

### v1.2.0 — 2026-02-04
- Added 8 new Drive tools (batch move, folder ops, export, permissions)
- Fixed OAuth flow (deprecated OOB to localhost redirect)

### v1.0.0 — 2025
- Initial release: Sheets, Docs, Drive basics

---

**Built with:** TypeScript, Google APIs, Model Context Protocol SDK

**License:** MIT
