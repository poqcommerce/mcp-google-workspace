# Google Workspace MCP Tools - Quick Reference

**Last Updated:** 2026-02-04

---

## Currently Available Tools ✅

### Google Sheets (5 tools)

| Tool | Purpose | Key Parameters |
|------|---------|----------------|
| `gsheets_batch_update` | Update multiple ranges in one call | spreadsheetId, updates[] |
| `gsheets_create_and_populate` | Create new sheet with data | title, data[][], parentFolderId? |
| `gsheets_append_rows` | Add rows to end of sheet | spreadsheetId, range, values[][] |
| `gsheets_format_cells` | Apply formatting to ranges | spreadsheetId, requests[] |
| `gsheets_get_auth_url` | Get OAuth authorization URL | - |

### Google Docs (7 tools)

| Tool | Purpose | Key Parameters |
|------|---------|----------------|
| `gdocs_create_document` | Create new document | title, content?, parentFolderId? |
| `gdocs_get_document` | Read document content | documentId |
| `gdocs_insert_text` | Insert text at position | documentId, text, index? |
| `gdocs_append_text` | Add text to end | documentId, text |
| `gdocs_replace_text` | Find and replace | documentId, find, replace |
| `gdocs_format_text` | Apply text formatting | documentId, startIndex, endIndex, format |
| `gdocs_set_heading` | Convert text to heading | documentId, startIndex, endIndex, headingLevel |

### Google Drive (4 tools)

| Tool | Purpose | Key Parameters |
|------|---------|----------------|
| `gdrive_search` | Search for files | query, pageSize?, pageToken? |
| `gdrive_read_file` | Read file contents | fileId |
| `gdrive_get_file_info` | Get file metadata | fileId |
| `gdrive_move_file` | Move file to folder | fileId, targetFolderId |

### Authentication (1 tool)

| Tool | Purpose | Key Parameters |
|------|---------|----------------|
| `gsheets_set_auth_code` | Exchange auth code for tokens | code |

**Total: 17 tools currently available**

---

## Proposed Tools for DSAR Workflows 📋

### Priority 1: Critical (Implement First) 🔴

| Tool | Purpose | Impact | Complexity |
|------|---------|--------|------------|
| `gdrive_create_folder` | Create folder programmatically | Very High | Low |
| `gdrive_list_folder_contents` | Get all files in folder (recursive) | Very High | Low |
| `gdrive_batch_move` | Move multiple files at once | High | Medium |
| `gdrive_batch_get_metadata` | Get metadata for multiple files | High | Medium |
| `gdrive_export_document` | Export to PDF/Word/Excel | Very High | Medium |

**Estimated time savings: 60-70%**

### Priority 2: High Value 🟡

| Tool | Purpose | Impact | Complexity |
|------|---------|--------|------------|
| `gdrive_advanced_search` | Search with multiple criteria | High | Medium |
| `gdocs_add_comment` | Add review comments | Medium | Medium |
| `gdocs_list_comments` | List all comments | Medium | Low |
| `gdrive_batch_download` | Download multiple files | High | High |
| `gdrive_add_custom_properties` | Tag files with metadata | Medium | Medium |

**Estimated additional savings: 15-20%**

### Priority 3: Nice to Have 🟢

| Tool | Purpose | Impact | Complexity |
|------|---------|--------|------------|
| `gdrive_get_permissions` | List file permissions | Low | Low |
| `gdrive_share_file` | Share file with user | Medium | Low |
| `gdrive_get_revisions` | Get version history | Low | Medium |
| `gdrive_find_duplicates` | Detect duplicate files | Medium | High |
| `gdrive_batch_rename` | Rename multiple files | Low | Low |
| `gdrive_get_folder_size` | Calculate folder size | Low | Medium |

**Estimated additional savings: 5-10%**

---

## DSAR Workflow Mapping

### Phase 1: Project Setup
**Current tools:**
- ✅ `gsheets_create_and_populate` - Create tracker
- ✅ `gdocs_create_document` - Create plan docs

**Needed tools:**
- 🔴 `gdrive_create_folder` - Set up folder structure
- 🔴 `gdrive_list_folder_contents` - Verify setup

---

### Phase 2: Document Collection
**Current tools:**
- ✅ `gdrive_search` - Find uploaded files
- ✅ `gdrive_get_file_info` - Check file details

**Needed tools:**
- 🔴 `gdrive_batch_get_metadata` - Generate inventory
- 🟡 `gdrive_add_custom_properties` - Tag files

---

### Phase 3: Review & Organization
**Current tools:**
- ✅ `gdrive_move_file` - Move individual files

**Needed tools:**
- 🔴 `gdrive_batch_move` - Organize by status
- 🟡 `gdocs_add_comment` - Add review notes
- 🟡 `gdrive_advanced_search` - Filter by criteria

---

### Phase 4: Redaction
**Current tools:**
- ✅ `gdocs_get_document` - Read content
- ✅ `gdocs_replace_text` - Basic redaction

**Needed tools:**
- 🟡 `gdocs_list_comments` - Track redactions
- 🟢 `gdrive_get_revisions` - Audit changes

---

### Phase 5: Delivery Preparation
**Current tools:**
- ✅ `gsheets_create_and_populate` - Create index

**Needed tools:**
- 🔴 `gdrive_export_document` - Export to PDF
- 🔴 `gdrive_batch_download` - Download all
- 🟢 `gdrive_get_folder_size` - Check package size

---

## Example Usage Patterns

### Pattern 1: Set Up DSAR Project
```typescript
// Create folder structure
const root = gdrive_create_folder({ name: "DSAR_Project" });
const folders = ["Gmail", "Slack", "Review", "Final"];

folders.forEach(name => {
  gdrive_create_folder({
    name,
    parentFolderId: root.folderId
  });
});

// Create tracker
gsheets_create_and_populate({
  title: "DSAR Tracker",
  data: [["Platform", "Status"], ...folders.map(f => [f, "Not Started"])],
  parentFolderId: root.folderId
});
```

### Pattern 2: Generate Document Inventory
```typescript
// List all collected files
const files = gdrive_list_folder_contents({
  folderId: dsarFolderId,
  recursive: true
});

// Get detailed metadata
const metadata = gdrive_batch_get_metadata({
  fileIds: files.map(f => f.id),
  fields: ["name", "size", "createdTime", "owners"]
});

// Create inventory spreadsheet
gsheets_create_and_populate({
  title: "Document Inventory",
  data: [
    ["ID", "Filename", "Size", "Date", "Owner"],
    ...metadata.map((f, i) => [i+1, f.name, f.size, f.createdTime, f.owners[0]])
  ],
  parentFolderId: dsarFolderId
});
```

### Pattern 3: Organize by Review Status
```typescript
// Search for completed reviews
const completed = gdrive_advanced_search({
  query: "properties has { key='review_status' and value='completed' }",
  folderIds: [reviewFolderId]
});

// Move to completed folder
gdrive_batch_move({
  fileIds: completed.map(f => f.id),
  targetFolderId: completedFolderId
});
```

### Pattern 4: Export for Delivery
```typescript
// Get all final documents
const finals = gdrive_list_folder_contents({
  folderId: finalFolderId,
  recursive: false
});

// Export each to PDF
finals.forEach(doc => {
  gdrive_export_document({
    fileId: doc.id,
    mimeType: "application/pdf",
    outputPath: `./delivery/${doc.name}.pdf`
  });
});

// Create delivery index
const index = gsheets_create_and_populate({
  title: "Delivery Index",
  data: [
    ["Doc ID", "Filename", "Format"],
    ...finals.map((f, i) => [`DOC_${i+1}`, f.name, "PDF"])
  ],
  parentFolderId: finalFolderId
});
```

---

## Tool Combinations for Common Tasks

### Task: Set up new DSAR project
**Tools needed:**
1. `gdrive_create_folder` × 8-10 (folder structure)
2. `gsheets_create_and_populate` × 2-3 (trackers, checklists)
3. `gdocs_create_document` × 2-3 (plans, notes)

### Task: Generate complete inventory
**Tools needed:**
1. `gdrive_list_folder_contents` × 1 (get all files)
2. `gdrive_batch_get_metadata` × 1 (get details)
3. `gsheets_create_and_populate` × 1 (create spreadsheet)

### Task: Review workflow
**Tools needed:**
1. `gdrive_advanced_search` (find pending reviews)
2. `gdocs_add_comment` (add review notes)
3. `gdrive_batch_move` (move to completed folder)
4. `gsheets_batch_update` (update tracker)

### Task: Prepare delivery package
**Tools needed:**
1. `gdrive_list_folder_contents` (get final docs)
2. `gdrive_export_document` × N (export each doc)
3. `gsheets_create_and_populate` (create index)
4. `gdrive_get_folder_size` (check package size)

---

## API Rate Limits & Best Practices

### Google Drive API Limits
- **Requests:** 20,000 per 100 seconds per user
- **Queries:** 1,000 per 100 seconds per user
- **Downloads:** 10,000 requests per 100 seconds

### Best Practices
1. **Batch operations:** Use batch tools for >10 files
2. **Pagination:** Use pageToken for large result sets
3. **Rate limiting:** Add delays between bulk operations
4. **Error handling:** Implement exponential backoff
5. **Caching:** Cache folder IDs and metadata when possible

### Example: Safe Batch Processing
```typescript
// Process files in batches of 50
const files = [...]; // Large array of file IDs
const batchSize = 50;

for (let i = 0; i < files.length; i += batchSize) {
  const batch = files.slice(i, i + batchSize);

  await gdrive_batch_move({
    fileIds: batch,
    targetFolderId: targetId
  });

  // Rate limit: Wait 2 seconds between batches
  await sleep(2000);
}
```

---

## OAuth Scopes Required

### Current Scopes
```
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/drive.file
https://www.googleapis.com/auth/documents
```

### Additional Scopes Needed for Proposed Tools
```
https://www.googleapis.com/auth/drive.metadata.readonly  // For advanced metadata
https://www.googleapis.com/auth/drive.activity.readonly  // For activity logs (Priority 3)
```

**Note:** Current `drive.file` scope covers most proposed tools!

---

## Implementation Checklist

### Phase 1: Core DSAR Tools (Week 1)
- [ ] Implement `gdrive_create_folder`
- [ ] Implement `gdrive_list_folder_contents`
- [ ] Implement `gdrive_batch_move`
- [ ] Implement `gdrive_batch_get_metadata`
- [ ] Implement `gdrive_export_document`
- [ ] Test with real DSAR folder structure
- [ ] Document examples and patterns

### Phase 2: Enhanced Workflow (Week 2)
- [ ] Implement `gdrive_advanced_search`
- [ ] Implement `gdocs_add_comment`
- [ ] Implement `gdocs_list_comments`
- [ ] Implement `gdrive_batch_download`
- [ ] Implement `gdrive_add_custom_properties`
- [ ] Create end-to-end DSAR workflow example

### Phase 3: Advanced Features (Week 3)
- [ ] Implement remaining Priority 3 tools
- [ ] Add comprehensive error handling
- [ ] Create automation scripts for common workflows
- [ ] Performance optimization
- [ ] Documentation and guides

---

## Resources

### Documentation
- **MCP Updates:** `MCP_UPDATES.md` - Recent changes and new features
- **DSAR Tools:** `SUGGESTED_DSAR_TOOLS.md` - Detailed tool proposals
- **This Document:** Quick reference for all tools

### API Documentation
- Google Drive API: https://developers.google.com/drive/api/v3
- Google Docs API: https://developers.google.com/docs/api
- Google Sheets API: https://developers.google.com/sheets/api

### Testing
- Test with small datasets first
- Use test folders, not production DSAR data
- Verify each tool individually before combining

---

## Quick Decision Matrix

**Choose `gdrive_move_file` when:**
- Moving single files
- Interactive workflow
- Immediate feedback needed

**Choose `gdrive_batch_move` when:**
- Moving >10 files
- Automated workflow
- Organizing after bulk operations

**Choose `gdrive_search` when:**
- Simple name/type queries
- Quick lookups
- Small result sets

**Choose `gdrive_advanced_search` when:**
- Complex multi-criteria searches
- Filtering by metadata
- Large result sets with ordering

**Choose `gdocs_create_document` when:**
- Creating individual docs
- Interactive document creation
- Custom content per doc

**Choose batch operations when:**
- Processing >10 items
- Automated workflows
- Performance critical

---

## Support

For questions or issues:
1. Check tool documentation in code comments
2. Review examples in `SUGGESTED_DSAR_TOOLS.md`
3. Test with simple cases first
4. Verify OAuth scopes and permissions

---

**Summary:** 17 tools available now, 15+ proposed for DSAR workflows. Phase 1 priorities would provide 60-70% time savings on DSAR projects.
