# Google Workspace MCP Server

A comprehensive Model Context Protocol (MCP) server providing programmatic access to Google Workspace (Drive, Docs, Sheets) for automation, document management, and workflow integration.

**Version:** 1.2.0
**Last Updated:** 2026-02-04
**Total Tools:** 27

---

## Features Overview

### 📄 Google Docs - 7 Tools
Create, read, edit documents with full formatting support

### 📊 Google Sheets - 5 Tools
Create, populate, update spreadsheets with batch operations

### 📁 Google Drive - 13 Tools
Complete file and folder management with advanced search

### 🔐 Authentication - 2 Tools
Modern OAuth 2.0 with secure token management

---

## Quick Start

### 1. Installation

```bash
git clone <repository-url>
cd mcp-google-workspace
npm install
npm run build
```

### 2. Google Cloud Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable these APIs:
   - Google Drive API
   - Google Docs API
   - Google Sheets API
4. Create OAuth 2.0 credentials (Desktop app type)
5. Download credentials

### 3. Configure Environment

Create `.env` file:

```bash
GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your-client-secret
GOOGLE_REFRESH_TOKEN=your-refresh-token
```

### 4. Authenticate

```bash
npm run auth
```

This opens your browser, requests permissions, and generates a refresh token. Copy the token to your `.env` file.

### 5. Start Server

```bash
npm start
```

Or for development:
```bash
npm run dev
```

---

## Complete Tool Reference

### Google Sheets (5 tools)

#### `gsheets_create_and_populate`
Create a spreadsheet with data in one operation.

```javascript
{
  title: "Q1 Sales Report",
  data: [
    ["Product", "Sales", "Region"],
    ["Widget A", "1250", "North"]
  ],
  parentFolderId: "optional-folder-id"
}
```

#### `gsheets_batch_update`
Update multiple ranges in a single API call (40x faster).

```javascript
{
  spreadsheetId: "sheet-id",
  updates: [
    {
      range: "A1:B10",
      values: [["Header", "Value"]]
    }
  ]
}
```

#### `gsheets_append_rows`
Add rows to the end of a sheet.

```javascript
{
  spreadsheetId: "sheet-id",
  range: "Sheet1!A:Z",
  values: [["New", "Row", "Data"]]
}
```

#### `gsheets_format_cells`
Apply styling to cell ranges.

```javascript
{
  spreadsheetId: "sheet-id",
  requests: [{
    range: "A1:B1",
    format: { textFormat: { bold: true } }
  }]
}
```

#### `gsheets_get_auth_url`
Get OAuth authorization URL for setup.

---

### Google Docs (7 tools)

#### `gdocs_create_document`
Create a new document.

```javascript
{
  title: "Project Proposal",
  content: "# Summary\n\nThis proposal...",
  parentFolderId: "optional-folder-id"
}
```

#### `gdocs_get_document`
Read document content.

```javascript
{
  documentId: "doc-id"
}
```

#### `gdocs_insert_text`
Insert text at specific position.

```javascript
{
  documentId: "doc-id",
  text: "New paragraph\n",
  index: 100  // optional, defaults to end
}
```

#### `gdocs_append_text`
Add text to document end.

```javascript
{
  documentId: "doc-id",
  text: "Additional content"
}
```

#### `gdocs_replace_text`
Find and replace throughout document.

```javascript
{
  documentId: "doc-id",
  find: "old text",
  replace: "new text"
}
```

#### `gdocs_format_text`
Apply text formatting.

```javascript
{
  documentId: "doc-id",
  startIndex: 0,
  endIndex: 20,
  format: {
    bold: true,
    italic: false,
    underline: true,
    fontSize: 16
  }
}
```

#### `gdocs_set_heading`
Convert text to heading (H1-H6).

```javascript
{
  documentId: "doc-id",
  startIndex: 0,
  endIndex: 15,
  headingLevel: 1  // 1-6
}
```

---

### Google Drive (13 tools)

#### File Operations

##### `gdrive_search`
Search for files with powerful queries.

```javascript
// By name
{ query: "name contains 'budget'" }

// By type
{ query: "mimeType='application/pdf'" }

// In folder
{ query: "'folder-id' in parents" }

// Combined
{ query: "name contains 'report' and mimeType='application/vnd.google-apps.document'" }
```

##### `gdrive_read_file`
Read file contents (Docs, Sheets, text files).

```javascript
{
  fileId: "file-id"
}
```

##### `gdrive_get_file_info`
Get file metadata.

```javascript
{
  fileId: "file-id"
}
```

##### `gdrive_move_file`
Move file to different folder.

```javascript
{
  fileId: "file-id",
  targetFolderId: "target-folder-id"
}
```

##### `gdrive_batch_move`
Move multiple files at once.

```javascript
{
  fileIds: ["id1", "id2", "id3"],
  targetFolderId: "target-folder-id"
}
```

#### Folder Operations

##### `gdrive_create_folder`
Create new folder.

```javascript
{
  name: "Project Alpha",
  parentFolderId: "optional-parent-id"
}
```

##### `gdrive_copy_folder`
Recursively copy folder and contents.

```javascript
{
  sourceFolderId: "source-id",
  targetParentFolderId: "optional-parent-id",
  newName: "optional-new-name"
}
```

##### `gdrive_list_folder_tree`
List all files in folder (recursive).

```javascript
{
  folderId: "folder-id",
  recursive: true,
  includeMetadata: true
}
```

Returns complete file tree with paths.

#### Advanced Operations

##### `gdrive_search_content`
Search within file contents (fullText).

```javascript
{
  query: "machine learning",
  folderId: "optional-folder-id",
  maxResults: 100
}
```

##### `gdrive_get_revisions`
Get version history for file.

```javascript
{
  fileId: "file-id",
  maxResults: 20
}
```

##### `gdrive_export_file`
Export file to specific format.

```javascript
{
  fileId: "file-id",
  mimeType: "application/pdf"
}
```

**Common MIME types:**
- PDF: `application/pdf`
- Word: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- Excel: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- PowerPoint: `application/vnd.openxmlformats-officedocument.presentationml.presentation`

##### `gdrive_batch_export`
Export multiple files at once.

```javascript
{
  fileIds: ["id1", "id2", "id3"],
  format: "pdf"  // pdf, docx, xlsx, pptx
}
```

##### `gdrive_list_permissions`
See who has access to file/folder.

```javascript
{
  fileId: "file-id"
}
```

Returns: type, role, email for each permission.

---

### Authentication (2 tools)

#### `gsheets_get_auth_url`
Get OAuth URL for initial setup.

#### `gsheets_set_auth_code`
Exchange auth code for tokens.

```javascript
{
  code: "authorization-code-from-google"
}
```

---

## Common Workflows

### Project Setup

```javascript
// Create folder structure
const root = gdrive_create_folder({
  name: "Project Alpha"
});

["Documents", "Data", "Reports"].forEach(name => {
  gdrive_create_folder({
    name,
    parentFolderId: root.folderId
  });
});
```

### Bulk File Organization

```javascript
// Search for files
const files = gdrive_search({
  query: "name contains 'Q1' and mimeType='application/pdf'"
});

// Move all at once
gdrive_batch_move({
  fileIds: files.map(f => f.id),
  targetFolderId: "archive-folder-id"
});
```

### Document Automation

```javascript
// Create report
const doc = gdocs_create_document({
  title: "Weekly Report",
  content: "# Weekly Report\n\n## Summary\n\n",
  parentFolderId: "reports-folder-id"
});

// Add content
gdocs_append_text({
  documentId: doc.documentId,
  text: "Key metrics:\n- Revenue: $50K\n- Growth: 15%\n"
});
```

### Export and Backup

```javascript
// Get all docs in folder
const files = gdrive_list_folder_tree({
  folderId: "project-folder-id",
  recursive: true
});

// Export all to PDF
const docs = files.filter(f =>
  f.type === 'application/vnd.google-apps.document'
);

gdrive_batch_export({
  fileIds: docs.map(d => d.id),
  format: 'pdf'
});
```

### Content Discovery

```javascript
// Search within file contents
const results = gdrive_search_content({
  query: "budget proposal",
  folderId: "search-folder-id"
});

// Check permissions for each
results.forEach(file => {
  const perms = gdrive_list_permissions({
    fileId: file.id
  });
});
```

### Template Processing

```javascript
// Create from template
const doc = gdocs_create_document({
  title: "Customer Proposal - Acme Corp",
  content: templateContent
});

// Customize
gdocs_replace_text({
  documentId: doc.documentId,
  find: "{{CLIENT_NAME}}",
  replace: "Acme Corporation"
});
```

---

## API Rate Limits

### Google Drive API
- Queries: 1,000 per 100 seconds
- Requests: 20,000 per 100 seconds per user
- Downloads: 10,000 per 100 seconds

### Best Practices
1. Use batch operations for >10 files
2. Add delays between bulk operations
3. Cache folder IDs to reduce queries
4. Use pagination for large result sets

---

## OAuth Scopes

Required scopes:
- `https://www.googleapis.com/auth/spreadsheets` - Sheets read/write
- `https://www.googleapis.com/auth/drive.file` - Drive file operations
- `https://www.googleapis.com/auth/documents` - Docs read/write

**Note:** `drive.file` scope provides access to files created or opened by this app.

---

## Troubleshooting

### "Insufficient permissions" error
**Solution:** Re-run `npm run auth` to refresh OAuth token with updated scopes.

### "File not found" error
**Solution:** Verify file ID and ensure file was created/opened by this app (drive.file scope limitation).

### "Rate limit exceeded" error
**Solution:** Add delays between API calls or reduce batch sizes.

### OAuth redirect not working
**Solution:** Ensure port 3000 is available. Uses `http://localhost:3000/oauth/callback`.

---

## Development

### Commands

```bash
npm run build     # Compile TypeScript
npm run dev       # Watch mode
npm run auth      # Run OAuth flow
npm start         # Start server
```

### Project Structure

```
mcp-google-workspace/
├── src/
│   ├── index.ts       # Main MCP server
│   └── auth.ts        # OAuth flow
├── dist/              # Compiled JS
├── .env               # Credentials
├── package.json
└── tsconfig.json
```

---

## Version History

### v1.2.0 - 2026-02-04
- ✨ Added 8 new tools (batch move, folder ops, export, search, permissions)
- 🔧 Fixed OAuth flow (deprecated OOB → modern localhost)
- 📈 Now 27 total tools (was 19)

### v1.1.0 - 2026-02-04
- ✨ Added `parentFolderId` to create files in folders
- ✨ Added `gdrive_move_file`
- 🔒 Updated OAuth scope to `drive.file`

### v1.0.0 - 2025
- 🎉 Initial release
- Sheets, Docs, Drive basic operations

---

## Performance

- **Batch operations:** 40x faster than individual API calls
- **Recursive operations:** Process entire folder trees
- **Export operations:** Bulk export hundreds of files
- **Search operations:** Content search across thousands of files

---

## Security

- ✅ Modern OAuth 2.0 with localhost redirect
- ✅ Secure token storage in .env
- ✅ Scoped permissions (only files created by app)
- ❌ Never commit `.env` or credentials
- 🔄 Rotate refresh tokens periodically

---

## Support

- **Documentation:** See `MCP_UPDATES.md` for detailed changes
- **Tool Reference:** See `TOOLS_REFERENCE.md` for comprehensive docs
- **Issues:** Report bugs via GitHub Issues
- **API Docs:**
  - [Google Sheets API](https://developers.google.com/sheets/api)
  - [Google Drive API](https://developers.google.com/drive/api)
  - [Google Docs API](https://developers.google.com/docs/api)

---

## Roadmap

### Phase 1: Collaboration & Sharing 🎯 **NEXT**

Complete the collaboration workflow with permission management:

- [ ] **`gdrive_share_file`** - Share files/folders with specific users or groups
  - Grant reader, writer, or commenter access by email
  - Essential for team collaboration workflows

- [ ] **`gdrive_create_share_link`** - Generate shareable links
  - Create public or domain-restricted links
  - Set expiration dates and access levels

- [ ] **`gdrive_update_permission`** - Modify existing permissions
  - Change user roles (reader → writer)
  - Update access settings

- [ ] **`gdrive_remove_permission`** - Revoke access
  - Remove specific users from files/folders
  - Cleanup after project completion

**Impact:** Completes permission management (we can read, now we can write)
**Use Cases:** Team collaboration, controlled distribution, access management

---

### Phase 1.5: Upload & File Transfer 📤 **ESSENTIAL**

Enable local file upload to complete the Drive workflow:

- [ ] **`gdrive_upload_file`** - Upload local file to Drive
  - Upload any file type (PDF, images, documents, etc.)
  - Specify target folder and optional new name
  - Automatic MIME type detection

- [ ] **`gdrive_upload_folder`** - Upload local folder recursively
  - Preserve folder structure
  - Batch upload multiple files
  - Progress tracking for large uploads

- [ ] **`gdrive_update_file_content`** - Replace existing file
  - Update file content from local file
  - Maintain file ID and metadata
  - Version control friendly

- [ ] **`gdrive_batch_upload`** - Upload multiple files at once
  - Bulk upload operations
  - Error handling per file
  - Resume capability for failed uploads

**Impact:** Completes the full Drive workflow (local → Drive → organize → export)
**Use Cases:** Backup automation, content migration, bulk file uploads, automated deployments
**Technical Notes:** Requires file streaming, chunking for large files, resumable uploads

---

### Phase 2: Organization & Workflow

Practical tools for file management at scale:

- [ ] **`gdrive_batch_rename`** - Rename multiple files with patterns
  - Sequential numbering (Doc_001, Doc_002...)
  - Pattern-based naming with variables
  - Standardize file naming conventions

- [ ] **`gdrive_set_properties`** - Add custom metadata/tags
  - Tag files with custom properties (status, priority, project)
  - Enable advanced filtering and organization
  - Track workflow states

- [ ] **`gdrive_batch_set_properties`** - Tag multiple files at once
  - Bulk metadata operations
  - Workflow automation

- [ ] **`gdrive_get_folder_stats`** - Analyze folder contents
  - Total size, file counts, type breakdown
  - Storage planning and monitoring
  - Recursive statistics

- [ ] **`gdrive_find_duplicates`** - Detect duplicate files
  - Find duplicates by name, size, or content
  - Storage cleanup and optimization
  - Deduplication workflows

**Impact:** Saves hours of manual file organization
**Use Cases:** File standardization, workflow tracking, storage management

---

### Phase 3: Advanced Features

Power user features for sophisticated workflows:

- [ ] **`gdocs_add_comment`** - Collaborative review workflow
  - Add comments to specific text ranges
  - Suggest edits and changes
  - Track review notes

- [ ] **`gdocs_list_comments`** - Track review progress
  - Get all comments on a document
  - Filter by resolved/unresolved
  - Export review notes

- [ ] **`gdocs_resolve_comment`** - Mark comments complete
  - Close review items
  - Track completion

- [ ] **`gdocs_create_from_template`** - Template automation
  - Create documents from templates
  - Variable substitution ({{NAME}}, {{DATE}}, etc.)
  - Automated document generation

- [ ] **`gdrive_get_activity`** - File activity logs
  - Track who accessed/modified files
  - Compliance and audit trails
  - Activity timeline

**Impact:** Enables advanced automation and compliance
**Use Cases:** Document review, templated content, audit trails

---

### Future Considerations

Additional capabilities under consideration:

- **OCR & Text Extraction** - Extract text from images/scans
- **Advanced Search Filters** - Search by date range, owner, multiple criteria
- **Webhook Support** - Event notifications for file changes
- **Quota Management** - Monitor API usage and limits
- **Batch Operations** - Additional bulk operations (delete, copy, etc.)
- **Drive Shortcuts** - Create and manage shortcuts
- **Starred Files** - Manage starred/favorite files

---

### Contributing to Roadmap

Have ideas for new features? We welcome:
- Feature requests via GitHub Issues
- Use case descriptions
- Pull requests with new tools
- Documentation improvements

---

## License

MIT License - See LICENSE file

---

## Contributing

1. Fork the repository
2. Create feature branch
3. Add tests for new functionality
4. Submit pull request

---

**Built with:** TypeScript • Google APIs • Model Context Protocol SDK

**Maintained by:** Your team name here

**Status:** Production ready ✅
