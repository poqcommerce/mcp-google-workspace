# Google Workspace MCP Server - Recent Updates

**Last Updated:** 2026-02-04
**Version:** 1.2.0

---

## Latest Updates (v1.2.0)

### 5 New Power Tools for DSAR Workflows 🚀

1. **`gdrive_batch_move`** - Move multiple files at once (bulk operations)
2. **`gdrive_create_folder`** - Create folders programmatically
3. **`gdrive_copy_folder`** - Recursively copy entire folder structures
4. **`gdrive_search_content`** - Search within file contents (fullText search)
5. **`gdrive_get_revisions`** - Access version history for audit trails

### Tool Details

#### 1. `gdrive_batch_move` - Bulk File Operations
```typescript
{
  fileIds: string[],      // Array of file IDs to move
  targetFolderId: string  // Destination folder
}
// Returns: { movedCount, failedCount, details }
```

**Use Case:** Move all reviewed DSAR documents to "Completed" folder in one operation
```typescript
gdrive_batch_move({
  fileIds: ["id1", "id2", "id3", ...],
  targetFolderId: "completed-folder-id"
})
```

#### 2. `gdrive_create_folder` - Programmatic Folder Creation
```typescript
{
  name: string,
  parentFolderId?: string  // Optional parent
}
// Returns: { folderId, name, url }
```

**Use Case:** Set up entire DSAR folder structure programmatically
```typescript
const root = gdrive_create_folder({ name: "DSAR_Project_2026" });
const platforms = ["Gmail", "Slack", "Jira", "Review"];
platforms.forEach(p => gdrive_create_folder({
  name: p,
  parentFolderId: root.folderId
}));
```

#### 3. `gdrive_copy_folder` - Recursive Folder Copying
```typescript
{
  sourceFolderId: string,
  targetParentFolderId?: string,  // Optional destination
  newName?: string                // Optional new name
}
// Returns: { newFolderId, copiedFiles, copiedFolders }
```

**Use Case:** Duplicate DSAR template structure for new requests
```typescript
gdrive_copy_folder({
  sourceFolderId: "template-folder-id",
  newName: "DSAR_NewSubject_2026"
})
```

#### 4. `gdrive_search_content` - Full-Text Search
```typescript
{
  query: string,           // Text to search for
  folderId?: string,       // Limit to specific folder
  maxResults?: number      // Default: 100
}
// Returns: { resultCount, results[] }
```

**Use Case:** Find all documents mentioning "redundancy" or "maternity leave"
```typescript
gdrive_search_content({
  query: "redundancy OR maternity",
  folderId: "dsar-folder-id",
  maxResults: 200
})
```

#### 5. `gdrive_get_revisions` - Version History
```typescript
{
  fileId: string,
  maxResults?: number  // Default: 20
}
// Returns: { revisionCount, revisions[] }
```

**Use Case:** Audit trail - track who modified redacted documents and when
```typescript
gdrive_get_revisions({
  fileId: "sensitive-doc-id",
  maxResults: 50
})
```

---

## Previous Updates (v1.1.0)

### 1. Create Files in Specific Folders ✨

Documents and spreadsheets can now be created directly in any Google Drive folder by specifying a `parentFolderId` parameter.

#### Updated Tools

**`gdocs_create_document`**
```typescript
{
  title: string,
  content?: string,
  parentFolderId?: string  // NEW - Folder ID where document should be created
}
```

**`gsheets_create_and_populate`**
```typescript
{
  title: string,
  data: string[][],
  sheetTitle?: string,
  parentFolderId?: string  // NEW - Folder ID where spreadsheet should be created
}
```

#### Example Usage

```typescript
// Create a document in a specific folder
gdocs_create_document({
  title: "DSAR Review Notes",
  content: "Review notes for Natasha Driver DSAR...",
  parentFolderId: "1rtnTkKhpazLMQNW5-PGvsY_7ULXBii7S"
})

// Create a spreadsheet in a specific folder
gsheets_create_and_populate({
  title: "DSAR Document Index",
  data: [
    ["Document", "Source", "Date", "Status"],
    ["Email_001.pdf", "Gmail", "2024-06-15", "Reviewed"]
  ],
  parentFolderId: "1rtnTkKhpazLMQNW5-PGvsY_7ULXBii7S"
})
```

---

### 2. Move Files Between Folders 🚀

A new tool allows you to move any file from one folder to another in Google Drive.

#### New Tool

**`gdrive_move_file`**
```typescript
{
  fileId: string,           // ID of the file to move
  targetFolderId: string    // ID of the destination folder
}
```

#### Example Usage

```typescript
// Move a document to the Review_Redaction folder
gdrive_move_file({
  fileId: "1Hyfu6Ox_cLgCzitE6rYytYhjPt0GNEPsaEwyIGpjxU0",
  targetFolderId: "1abc123_ReviewFolder_xyz"
})
```

**Note:** This removes the file from all previous parent folders and places it in the new folder.

---

### 3. Updated OAuth Scopes 🔐

The Drive API scope has been upgraded to support write operations:

**Before:** `https://www.googleapis.com/auth/drive.readonly`
**After:** `https://www.googleapis.com/auth/drive.file`

#### Re-authentication Required

If you get permission errors, you'll need to re-authorize:

```bash
# Get new authorization URL
gsheets_get_auth_url()

# Follow the URL, authorize, and paste the code
gsheets_set_auth_code("your-authorization-code-here")

# Update your .env file with the new refresh token
```

---

## Implementation Details

### File Creation in Folders

When you specify a `parentFolderId`, the MCP server:
1. Creates the document/spreadsheet using the native API
2. Retrieves the current parent folder(s) of the new file
3. Moves the file to the specified folder using the Drive API
4. Returns success with the file ID and URL

### File Moving

The `gdrive_move_file` tool:
1. Gets the current parent folder(s) of the file
2. Removes the file from all current parents
3. Adds the file to the new target folder
4. Returns details about the move operation

---

## Breaking Changes

None. All existing tools continue to work exactly as before. The `parentFolderId` parameter is optional.

---

## Use Cases for DSAR Projects

### 1. Organized Document Collection
```typescript
// Create folder structure
const dsarFolderId = "1rtnTkKhpazLMQNW5-PGvsY_7ULXBii7S";

// Create review tracker in DSAR folder
gsheets_create_and_populate({
  title: "DSAR Review Tracker",
  data: [
    ["Platform", "Files Collected", "Reviewed", "Redacted", "Status"],
    ["Gmail", "150", "45", "12", "In Progress"],
    ["Slack", "230", "0", "0", "Not Started"]
  ],
  parentFolderId: dsarFolderId
});

// Create redaction notes doc in DSAR folder
gdocs_create_document({
  title: "Redaction Log - Natasha Driver",
  content: "# Redaction Log\n\n## Third-Party Data Redactions\n\n",
  parentFolderId: dsarFolderId
});
```

### 2. Workflow Management
```typescript
// Move completed reviews to different folders
gdrive_move_file({
  fileId: "doc-being-reviewed-123",
  targetFolderId: "completed-reviews-folder-456"
});
```

### 3. Automated Organization
```typescript
// After processing exports, create summary docs in the project folder
gdocs_create_document({
  title: "Gmail Export Summary",
  content: generateSummary(gmailData),
  parentFolderId: dsarFolderId
});
```

---

## Error Handling

### Common Errors

**"Insufficient permissions"**
- Solution: Re-authorize with the updated scopes using `gsheets_get_auth_url`

**"File not found"**
- Solution: Verify the `fileId` is correct using `gdrive_search` or `gdrive_get_file_info`

**"Folder not found"**
- Solution: Verify the `parentFolderId` exists using `gdrive_get_file_info`

**"User rate limit exceeded"**
- Solution: Add delays between API calls, or reduce batch sizes

---

## Testing the New Features

### 1. Test File Creation in Folder
```typescript
// Create a test document
const result = gdocs_create_document({
  title: "Test Document",
  content: "This is a test",
  parentFolderId: "your-folder-id"
});

// Verify it's in the folder
gdrive_search({
  query: "'your-folder-id' in parents"
});
```

### 2. Test File Moving
```typescript
// Get a file ID
const files = gdrive_search({ query: "name contains 'test'" });

// Move it
gdrive_move_file({
  fileId: "file-id-from-search",
  targetFolderId: "target-folder-id"
});

// Verify the move
gdrive_get_file_info({ fileId: "file-id-from-search" });
```

---

## Performance Considerations

### API Rate Limits
- Google Drive API: 20,000 requests per 100 seconds per user
- Individual file operations are single API calls
- Batch operations recommended for large-scale moves

### Best Practices
1. **Batch operations:** Group similar operations together
2. **Error handling:** Always check operation results
3. **Folder verification:** Verify folder IDs before bulk operations
4. **Audit logging:** Track all file moves for compliance

---

## Migration Guide

### From Root Creation to Folder Creation

**Old approach:**
```typescript
// Create in root, manually move later
const doc = gdocs_create_document({
  title: "My Document",
  content: "Content"
});
// Manual move in UI
```

**New approach:**
```typescript
// Create directly in target folder
const doc = gdocs_create_document({
  title: "My Document",
  content: "Content",
  parentFolderId: "folder-id"
});
```

---

## Roadmap

### Potential Future Enhancements
- [ ] Bulk file moving (move multiple files at once)
- [ ] Folder creation tool
- [ ] Recursive folder copying
- [ ] File permissions management
- [ ] Sharing link generation
- [ ] Export documents to PDF/Word
- [ ] Batch file metadata updates
- [ ] Drive activity logs
- [ ] Search within file contents
- [ ] Version history access

**See "Suggested Additional Tools" section below for prioritized recommendations.**

---

## Support

### Documentation
- Google Drive API: https://developers.google.com/drive/api/v3/reference
- Google Docs API: https://developers.google.com/docs/api
- Google Sheets API: https://developers.google.com/sheets/api

### Troubleshooting
1. Check MCP server logs: Look for error messages in stderr
2. Verify authentication: Test with `gdrive_search` first
3. Check permissions: Ensure Drive scope is `drive.file` not `drive.readonly`
4. Test with simple operations: Create a test doc before bulk operations

---

## Changelog

### v1.2.0 - 2026-02-04 (Latest)
- Added `gdrive_batch_move` - Bulk file moving operations
- Added `gdrive_create_folder` - Programmatic folder creation
- Added `gdrive_copy_folder` - Recursive folder copying with all contents
- Added `gdrive_search_content` - Full-text search within file contents
- Added `gdrive_get_revisions` - Version history access for audit trails
- Fixed deprecated OAuth flow - migrated from OOB to localhost redirect
- Updated authentication to use modern OAuth 2.0 protocol

### v1.1.0 - 2026-02-04
- Added `parentFolderId` parameter to `gdocs_create_document`
- Added `parentFolderId` parameter to `gsheets_create_and_populate`
- Added `gdrive_move_file` tool for moving files between folders
- Updated Drive API scope from `drive.readonly` to `drive.file`
- Added comprehensive error handling for file operations

### v1.0.0 - 2025
- Initial release with basic Google Workspace tools
- Sheets batch update, create, append, format
- Docs create, read, insert, append, replace, format
- Drive search, read, get file info
- OAuth authentication flow

---

**Questions or issues?** Check the main README.md or open an issue in the repository.
