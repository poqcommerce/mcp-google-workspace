# Suggested Google Workspace Tools for DSAR Projects

**Purpose:** Enhance the MCP server with tools specifically designed for Data Subject Access Request (DSAR) workflows

**Context:** DSAR projects involve collecting, organizing, reviewing, redacting, and delivering large volumes of documents from multiple sources. These suggested tools would streamline the entire DSAR lifecycle.

---

## Priority 1: Critical for DSAR Workflow 🔴

### 1. Folder Management

**`gdrive_create_folder`**
```typescript
{
  name: string,
  parentFolderId?: string
}
// Returns: { folderId, name, url }
```

**Use Case:**
- Programmatically create folder structure for DSAR project
- Create folders for each platform (Gmail, Slack, Jira, etc.)
- Create Review_Redaction subfolders automatically

**Example:**
```typescript
// Set up DSAR folder structure
const dsarRoot = gdrive_create_folder({
  name: "DSAR_Natasha_Driver_2026"
});

const folders = ["Gmail_Drive", "Slack", "Jira", "Confluence",
                 "Review_Redaction", "Final_Deliverables"];

folders.forEach(name => {
  gdrive_create_folder({
    name,
    parentFolderId: dsarRoot.folderId
  });
});
```

---

**`gdrive_list_folder_contents`**
```typescript
{
  folderId: string,
  recursive?: boolean,
  includeMetadata?: boolean
}
// Returns: Array of files with metadata
```

**Use Case:**
- Get complete inventory of collected documents
- Generate file lists for review tracking
- Verify completeness of exports

**Example:**
```typescript
// Get all files in DSAR folder for inventory
const files = gdrive_list_folder_contents({
  folderId: "1rtnTkKhpazLMQNW5-PGvsY_7ULXBii7S",
  recursive: true,
  includeMetadata: true
});

// Generate inventory spreadsheet
const inventory = files.map(f => [
  f.name, f.mimeType, f.size, f.modifiedTime, f.owners
]);
```

---

### 2. Batch File Operations

**`gdrive_batch_move`**
```typescript
{
  fileIds: string[],
  targetFolderId: string
}
// Returns: { successCount, failures: [] }
```

**Use Case:**
- Move all reviewed documents to "Completed" folder
- Organize files by status (To Review, In Progress, Completed)
- Bulk organization after filtering

**Example:**
```typescript
// Move all reviewed Gmail exports to completed folder
const reviewedFiles = ["id1", "id2", "id3", ...];
gdrive_batch_move({
  fileIds: reviewedFiles,
  targetFolderId: completedFolderId
});
```

---

**`gdrive_batch_get_metadata`**
```typescript
{
  fileIds: string[],
  fields?: string[]  // e.g., ["name", "size", "owners", "modifiedTime"]
}
// Returns: Array of file metadata objects
```

**Use Case:**
- Generate document index for DSAR delivery
- Create audit trail of all collected documents
- Track file sizes for package planning

**Example:**
```typescript
// Get metadata for all files to create index
const allFiles = getAllFileIds();
const metadata = gdrive_batch_get_metadata({
  fileIds: allFiles,
  fields: ["name", "size", "createdTime", "owners"]
});

// Create index spreadsheet
gsheets_create_and_populate({
  title: "DSAR Document Index",
  data: [
    ["Filename", "Size", "Created", "Owner"],
    ...metadata.map(f => [f.name, f.size, f.createdTime, f.owners[0]])
  ]
});
```

---

### 3. Export & Download Operations

**`gdrive_export_document`**
```typescript
{
  fileId: string,
  mimeType: string,  // "application/pdf", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", etc.
  outputPath?: string  // Where to save locally
}
// Returns: { success, filePath, size }
```

**Use Case:**
- Export all Google Docs to PDF for DSAR delivery
- Convert Sheets to Excel for universal compatibility
- Create final deliverable package

**Example:**
```typescript
// Export all docs in folder to PDF
const docs = gdrive_search({
  query: "'folder-id' in parents and mimeType='application/vnd.google-apps.document'"
});

docs.forEach(doc => {
  gdrive_export_document({
    fileId: doc.id,
    mimeType: "application/pdf",
    outputPath: `./exports/${doc.name}.pdf`
  });
});
```

---

**`gdrive_batch_download`**
```typescript
{
  fileIds: string[],
  outputDirectory: string,
  exportFormat?: string  // For Google Docs/Sheets
}
// Returns: { successCount, failures: [], totalSize }
```

**Use Case:**
- Download all collected documents for offline review
- Create local backup before delivery
- Prepare files for redaction software

---

### 4. Advanced Search & Filtering

**`gdrive_advanced_search`**
```typescript
{
  query: string,
  owner?: string,
  mimeType?: string,
  modifiedAfter?: Date,
  modifiedBefore?: Date,
  folderIds?: string[],  // Search within specific folders
  orderBy?: string,       // e.g., "modifiedTime desc"
  maxResults?: number
}
// Returns: Array of matching files with full metadata
```

**Use Case:**
- Find all documents modified during DSAR collection period
- Search for specific document types across folders
- Filter by owner (custodian) for DSAR scope

**Example:**
```typescript
// Find all PDFs created in June 2024 by specific custodian
const results = gdrive_advanced_search({
  mimeType: "application/pdf",
  modifiedAfter: new Date("2024-06-01"),
  modifiedBefore: new Date("2024-06-30"),
  owner: "james.peake@poqcommerce.com",
  folderIds: [dsarFolderId]
});
```

---

## Priority 2: Highly Valuable for Review Workflow 🟡

### 5. Comments & Collaboration

**`gdocs_add_comment`**
```typescript
{
  documentId: string,
  text: string,
  position?: { startIndex: number, endIndex: number }
}
// Returns: { commentId, created, author }
```

**Use Case:**
- Add review notes to documents
- Flag items for redaction
- Collaborate on privilege review

**Example:**
```typescript
// Flag document for redaction review
gdocs_add_comment({
  documentId: "doc-123",
  text: "REDACTION NEEDED: Third-party personal data on page 2"
});
```

---

**`gdocs_list_comments`**
```typescript
{
  documentId: string,
  resolved?: boolean  // Filter by resolution status
}
// Returns: Array of comments with metadata
```

**Use Case:**
- Track review progress
- Export review notes
- Generate QA checklist

---

**`gdocs_resolve_comment`**
```typescript
{
  documentId: string,
  commentId: string
}
```

**Use Case:**
- Mark redactions as completed
- Track review workflow progress

---

### 6. Permissions & Sharing

**`gdrive_get_permissions`**
```typescript
{
  fileId: string
}
// Returns: Array of { email, role, type } for each permission
```

**Use Case:**
- Audit trail: Document who had access to DSAR files
- Compliance: Verify restricted access
- Security review

---

**`gdrive_share_file`**
```typescript
{
  fileId: string,
  email: string,
  role: "reader" | "commenter" | "writer",
  sendNotification?: boolean
}
// Returns: { permissionId, success }
```

**Use Case:**
- Share review folders with legal team
- Grant access to HR for specific documents
- Controlled distribution of DSAR materials

---

**`gdrive_batch_set_permissions`**
```typescript
{
  fileIds: string[],
  permissions: { email: string, role: string }[]
}
```

**Use Case:**
- Grant legal team access to all review documents
- Set up review workflow permissions
- Restrict access after completion

---

### 7. Metadata & Organization

**`gdrive_set_description`**
```typescript
{
  fileId: string,
  description: string
}
```

**Use Case:**
- Add metadata about document source
- Tag documents with review status
- Add redaction notes

**Example:**
```typescript
gdrive_set_description({
  fileId: "doc-123",
  description: "Source: Gmail | Custodian: James Peake | Date: 2024-06-15 | Status: Redacted | Reviewer: Legal"
});
```

---

**`gdrive_add_custom_properties`**
```typescript
{
  fileId: string,
  properties: Record<string, string>
}
```

**Use Case:**
- Tag documents with structured metadata
- Track review workflow status
- Filter by custom properties

**Example:**
```typescript
gdrive_add_custom_properties({
  fileId: "doc-123",
  properties: {
    "dsar_project": "Natasha_Driver_2026",
    "platform": "Gmail",
    "custodian": "james.peake@poqcommerce.com",
    "review_status": "completed",
    "redaction_required": "yes",
    "privilege_claim": "no"
  }
});
```

---

### 8. Version History & Audit

**`gdrive_get_revisions`**
```typescript
{
  fileId: string,
  maxResults?: number
}
// Returns: Array of { id, modifiedTime, modifyingUser, size }
```

**Use Case:**
- Audit trail of document changes
- Track redaction history
- Compliance documentation

---

**`gdrive_restore_revision`**
```typescript
{
  fileId: string,
  revisionId: string
}
```

**Use Case:**
- Undo incorrect redactions
- Restore original documents
- Quality control

---

## Priority 3: Nice to Have 🟢

### 9. Duplicate Detection

**`gdrive_find_duplicates`**
```typescript
{
  folderId: string,
  criteria: "name" | "content" | "md5"
}
// Returns: Array of duplicate file groups
```

**Use Case:**
- Deduplicate collected documents
- Reduce review effort
- Minimize delivery package size

---

### 10. Bulk Rename

**`gdrive_batch_rename`**
```typescript
{
  renames: Array<{ fileId: string, newName: string }>
}
```

**Use Case:**
- Standardize file naming conventions
- Add platform prefixes (e.g., "Gmail_", "Slack_")
- Prepare files for delivery

**Example:**
```typescript
// Add platform prefix to all files
const files = gdrive_list_folder_contents({ folderId: gmailFolderId });
const renames = files.map(f => ({
  fileId: f.id,
  newName: `Gmail_${f.name}`
}));
gdrive_batch_rename({ renames });
```

---

### 11. Size & Statistics

**`gdrive_get_folder_size`**
```typescript
{
  folderId: string,
  recursive: boolean
}
// Returns: { totalSize, fileCount, breakdown: {} }
```

**Use Case:**
- Plan delivery method (email vs. cloud link)
- Track collection progress
- Estimate storage needs

---

### 12. Activity Logs

**`gdrive_get_activity`**
```typescript
{
  fileId: string,
  startTime?: Date,
  endTime?: Date
}
// Returns: Array of { actor, action, timestamp }
```

**Use Case:**
- Audit trail for compliance
- Track who accessed DSAR files
- Document chain of custody

---

## Suggested Implementation Priority

### Phase 1: Core DSAR Tools (Implement First)
1. `gdrive_create_folder` - Essential for project setup
2. `gdrive_list_folder_contents` - Essential for inventory
3. `gdrive_batch_move` - Essential for workflow
4. `gdrive_batch_get_metadata` - Essential for document index
5. `gdrive_export_document` - Essential for delivery

**Estimated Impact:** Reduces manual work by 60-70%

### Phase 2: Review Workflow Tools
6. `gdrive_advanced_search` - High value for filtering
7. `gdocs_add_comment` - Valuable for collaboration
8. `gdocs_list_comments` - Valuable for tracking
9. `gdrive_batch_download` - Valuable for offline work
10. `gdrive_add_custom_properties` - Valuable for organization

**Estimated Impact:** Additional 15-20% efficiency gain

### Phase 3: Advanced Features
11. `gdrive_get_permissions` - Nice to have for audit
12. `gdrive_get_revisions` - Nice to have for compliance
13. `gdrive_find_duplicates` - Nice to have for efficiency
14. `gdrive_batch_rename` - Nice to have for consistency
15. `gdrive_get_folder_size` - Nice to have for planning

**Estimated Impact:** Additional 5-10% efficiency gain

---

## Example: Complete DSAR Workflow with Proposed Tools

```typescript
// 1. SET UP PROJECT STRUCTURE
const dsarRoot = gdrive_create_folder({
  name: "DSAR_Natasha_Driver_2026"
});

const platforms = ["Gmail", "Slack", "Jira", "Confluence", "Figma",
                   "CharlieHR", "Review", "Final_Deliverables"];

const folderIds = {};
platforms.forEach(name => {
  const folder = gdrive_create_folder({
    name,
    parentFolderId: dsarRoot.folderId
  });
  folderIds[name] = folder.folderId;
});

// 2. CREATE TRACKING SPREADSHEET
const tracker = gsheets_create_and_populate({
  title: "DSAR Review Tracker",
  data: [
    ["Platform", "Files", "Reviewed", "Redacted", "Completed"],
    ...platforms.map(p => [p, "0", "0", "0", "No"])
  ],
  parentFolderId: dsarRoot.folderId
});

// 3. COLLECT DOCUMENTS (manual exports, then upload to folders)
// ... user uploads exports to respective platform folders ...

// 4. GENERATE INVENTORY
const allFiles = gdrive_list_folder_contents({
  folderId: dsarRoot.folderId,
  recursive: true,
  includeMetadata: true
});

const metadata = gdrive_batch_get_metadata({
  fileIds: allFiles.map(f => f.id),
  fields: ["name", "size", "createdTime", "owners", "mimeType"]
});

const inventory = gsheets_create_and_populate({
  title: "Document Inventory",
  data: [
    ["Filename", "Platform", "Size", "Type", "Created"],
    ...metadata.map(f => [
      f.name,
      getPlatformFromPath(f),
      formatSize(f.size),
      f.mimeType,
      f.createdTime
    ])
  ],
  parentFolderId: dsarRoot.folderId
});

// 5. TAG FILES FOR REVIEW
allFiles.forEach(file => {
  gdrive_add_custom_properties({
    fileId: file.id,
    properties: {
      "dsar_project": "Natasha_Driver_2026",
      "platform": getPlatformFromPath(file),
      "review_status": "pending",
      "reviewer": "",
      "redaction_required": "unknown"
    }
  });
});

// 6. SEARCH FOR SENSITIVE CONTENT
const possiblePII = gdrive_advanced_search({
  query: "fullText contains 'passport' or fullText contains 'SSN' or fullText contains 'bank'",
  folderIds: [dsarRoot.folderId],
  maxResults: 1000
});

// Flag these for priority review
possiblePII.forEach(file => {
  gdocs_add_comment({
    documentId: file.id,
    text: "⚠️ PRIORITY REVIEW: Potential third-party PII detected"
  });
});

// 7. ORGANIZE BY REVIEW STATUS
const reviewed = gdrive_advanced_search({
  query: "properties has { key='review_status' and value='completed' }",
  folderIds: [dsarRoot.folderId]
});

gdrive_batch_move({
  fileIds: reviewed.map(f => f.id),
  targetFolderId: folderIds["Review"]
});

// 8. EXPORT FOR DELIVERY
const finalDocs = gdrive_list_folder_contents({
  folderId: folderIds["Review"],
  recursive: false
});

finalDocs.forEach(doc => {
  gdrive_export_document({
    fileId: doc.id,
    mimeType: "application/pdf",
    outputPath: `./final_deliverables/${doc.name}.pdf`
  });
});

// 9. CREATE FINAL INDEX
const deliveryIndex = gsheets_create_and_populate({
  title: "DSAR Delivery Index",
  data: [
    ["Document ID", "Filename", "Platform", "Date Range", "Pages"],
    ...finalDocs.map((doc, i) => [
      `DOC_${String(i+1).padStart(4, '0')}`,
      doc.name,
      getPlatformFromPath(doc),
      "2024-06-01 to 2026-01-14",
      getPageCount(doc)
    ])
  ],
  parentFolderId: folderIds["Final_Deliverables"]
});

// 10. AUDIT TRAIL
const auditLog = gdocs_create_document({
  title: "DSAR Audit Trail",
  content: generateAuditLog(allFiles),
  parentFolderId: dsarRoot.folderId
});
```

---

## Estimated Time Savings

### Without Proposed Tools (Manual Process)
- Folder creation: 15 minutes
- File organization: 2-3 hours
- Inventory generation: 1-2 hours
- Search & filter: 1-2 hours
- Export preparation: 2-3 hours
- **Total: 6-10 hours**

### With Proposed Tools (Automated Process)
- All operations scripted: 15-30 minutes
- Human time: Review & verify only
- **Total: 30 minutes - 1 hour**

**Time Savings: 85-90% reduction in manual work**

---

## Implementation Recommendations

### Quick Wins (Implement in 1-2 days)
1. `gdrive_create_folder` - Simple Drive API call
2. `gdrive_list_folder_contents` - Basic Drive API listing
3. `gdrive_batch_move` - Iterate existing move function

### Medium Complexity (Implement in 3-5 days)
4. `gdrive_batch_get_metadata` - Batch API calls
5. `gdrive_export_document` - Export + download logic
6. `gdrive_advanced_search` - Complex query builder

### Advanced Features (Implement in 1-2 weeks)
7. `gdocs_add_comment` - Docs API integration
8. `gdrive_add_custom_properties` - Metadata API
9. `gdrive_get_revisions` - Version history API

---

## Questions to Consider

1. **Scope:** Should these tools be DSAR-specific or general-purpose?
2. **Batch sizes:** What's the maximum batch size for operations?
3. **Rate limiting:** How should we handle Drive API rate limits?
4. **Error handling:** Retry logic for failed operations?
5. **Logging:** Should we create audit logs automatically?
6. **Permissions:** What Drive scopes are needed for each tool?

---

## Next Steps

1. Review this proposal and prioritize tools
2. Validate use cases against your DSAR workflow
3. Implement Phase 1 tools first (core DSAR functionality)
4. Test with real DSAR project structure
5. Iterate based on feedback

---

**Would these tools be valuable for your DSAR workflow?** Let's discuss which ones to prioritize for implementation!
