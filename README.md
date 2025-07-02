# Google Workspace MCP

A comprehensive Model Context Protocol (MCP) server that provides efficient operations for Google Sheets, full Google Drive access, and complete Google Docs functionality, reducing API calls by up to 40x compared to individual operations.

## Features

### Google Sheets
- 🚀 **Batch Operations**: Update multiple ranges in a single API call
- 📊 **Create & Populate**: Create new sheets with data in one operation
- ➕ **Append Rows**: Efficiently add new data to existing sheets
- 🎨 **Format Cells**: Apply formatting to multiple ranges at once

### Google Docs
- 📝 **Document Creation**: Create new Google Docs with initial content
- ✏️ **Text Operations**: Insert, append, and replace text efficiently
- 🎨 **Rich Formatting**: Apply bold, italic, underline, and font sizing
- 📑 **Heading Management**: Convert text to headings (H1-H6)

### Google Drive
- 🔍 **Advanced Search**: Find files with powerful query syntax
- 📖 **File Reading**: Read Google Docs, Sheets, and text files
- 📋 **Metadata Access**: Get detailed information about Drive files

### General
- 🔐 **OAuth2 Authentication**: Secure user-based access to Google Workspace
- ⚡ **Performance Optimized**: Batch operations for maximum efficiency

## Quick Start

### Prerequisites

- Node.js 18+ 
- Google Cloud Project with Sheets API, Drive API, and Docs API enabled
- Claude Desktop or compatible MCP client

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/poqcommerce/mcp-google-workspace.git
   cd mcp-google-workspace
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Build the project:**
   ```bash
   npm run build
   ```

### Authentication Setup

1. **Get Google OAuth credentials:**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Enable Google Sheets API, Google Drive API, and Google Docs API
   - Create OAuth 2.0 Client ID (Desktop Application)
   - Copy Client ID and Secret

2. **Run the authentication flow:**
   ```bash
   GOOGLE_CLIENT_ID="your-id" GOOGLE_CLIENT_SECRET="your-secret" npm run auth
   ```
   
   This will authorize access to:
   - ✓ Google Sheets (read/write)
   - ✓ Google Drive (read-only)
   - ✓ Google Docs (read/write)

3. **Save the refresh token** for use in your MCP client configuration.

### Claude Desktop Configuration

You have two options for configuration:

#### **Configuration with Shell Script (Recommended)**

Your setup uses a shell script approach for better security:

1. **Clone and setup:**
   ```bash
   git clone https://github.com/poqcommerce/mcp-google-workspace.git
   cd mcp-google-workspace
   npm install
   npm run build
   ```

2. **Create your environment file:**
   ```bash
   cp .env.example .env
   # Edit .env with your actual Google OAuth credentials
   ```

3. **Update Claude Desktop config:**
   ```json
   {
     "mcpServers": {
       "google-workspace": {
         "command": "/path/to/mcp-google-workspace/start-mcp.sh"
       }
     }
   }
   ```

The included `start-mcp.sh` script handles loading environment variables and starting the server.

#### **Alternative: Direct Configuration**

If you prefer to put credentials directly in Claude config:

```json
{
  "mcpServers": {
    "google-workspace": {
      "command": "node",
      "args": ["/path/to/mcp-google-workspace/dist/index.js"],
      "env": {
        "GOOGLE_CLIENT_ID": "your-client-id",
        "GOOGLE_CLIENT_SECRET": "your-client-secret", 
        "GOOGLE_REFRESH_TOKEN": "your-refresh-token"
      }
    }
  }
}
```

## Available Tools

### Google Sheets Operations

#### `gsheets_batch_update`
Update multiple ranges in a single API call (40x faster than individual updates).

```typescript
await mcp.call('gsheets_batch_update', {
  spreadsheetId: 'your-sheet-id',
  updates: [
    {
      range: 'A1:B10',
      values: [['Header 1', 'Header 2'], ['Data 1', 'Data 2']]
    }
  ]
});
```

#### `gsheets_create_and_populate`
Create a new sheet with initial data.

```typescript
await mcp.call('gsheets_create_and_populate', {
  title: 'My New Sheet',
  data: [['Name', 'Value'], ['Item 1', '100']]
});
```

#### `gsheets_append_rows`
Add new rows to an existing sheet.

```typescript
await mcp.call('gsheets_append_rows', {
  spreadsheetId: 'your-sheet-id',
  range: 'Sheet1!A:Z',
  values: [['New', 'Data', 'Row']]
});
```

#### `gsheets_format_cells`
Apply formatting to cell ranges.

```typescript
await mcp.call('gsheets_format_cells', {
  spreadsheetId: 'your-sheet-id',
  requests: [
    {
      range: 'A1:B1',
      format: { textFormat: { bold: true } }
    }
  ]
});
```

### Google Docs Operations

#### `gdocs_create_document`
Create a new Google Document with optional initial content.

```typescript
await mcp.call('gdocs_create_document', {
  title: 'My New Document',
  content: 'This is the initial content of my document.'
});
```

#### `gdocs_get_document`
Read the full content of a Google Document.

```typescript
await mcp.call('gdocs_get_document', {
  documentId: 'your-document-id'
});
```

#### `gdocs_insert_text`
Insert text at a specific position in the document.

```typescript
await mcp.call('gdocs_insert_text', {
  documentId: 'your-document-id',
  text: 'New paragraph content\n',
  index: 100  // Optional: position to insert (defaults to end)
});
```

#### `gdocs_append_text`
Add text to the end of a document.

```typescript
await mcp.call('gdocs_append_text', {
  documentId: 'your-document-id',
  text: 'This text will be added to the end of the document.'
});
```

#### `gdocs_replace_text`
Find and replace text throughout the document.

```typescript
await mcp.call('gdocs_replace_text', {
  documentId: 'your-document-id',
  find: 'old text',
  replace: 'new text'
});
```

#### `gdocs_format_text`
Apply formatting to specific text ranges.

```typescript
await mcp.call('gdocs_format_text', {
  documentId: 'your-document-id',
  startIndex: 0,
  endIndex: 20,
  format: {
    bold: true,
    italic: false,
    underline: true,
    fontSize: 16
  }
});
```

#### `gdocs_set_heading`
Convert text to a heading (H1-H6).

```typescript
await mcp.call('gdocs_set_heading', {
  documentId: 'your-document-id',
  startIndex: 0,
  endIndex: 15,
  headingLevel: 1  // 1-6 for H1-H6
});
```

### Google Drive Operations

#### `gdrive_search`
Search for files in Google Drive with powerful query syntax.

```typescript
await mcp.call('gdrive_search', {
  query: "name contains 'budget' and type = 'spreadsheet'",
  pageSize: 20
});
```

**Common search queries:**
- `name contains 'KPI'` - Files with "KPI" in the name
- `type = 'document'` - Google Docs only
- `type = 'spreadsheet'` - Google Sheets only
- `modifiedTime > '2024-01-01'` - Recently modified files
- `'your-folder-id' in parents` - Files in specific folder

#### `gdrive_read_file`
Read contents of Google Drive files (Docs, Sheets as CSV, text files).

```typescript
await mcp.call('gdrive_read_file', {
  fileId: 'your-file-id'
});
```

**Supported file types:**
- Google Docs (exported as plain text)
- Google Sheets (exported as CSV)
- Text files (.txt, .md, .json, etc.)

#### `gdrive_get_file_info`
Get detailed metadata about a file.

```typescript
await mcp.call('gdrive_get_file_info', {
  fileId: 'your-file-id'
});
```

### Authentication Tools

#### `gsheets_get_auth_url`
Get the OAuth authorization URL for initial setup.

#### `gsheets_set_auth_code`
Exchange authorization code for access tokens.

## Performance Benefits

- **Before**: 40+ individual API calls for complex dashboards and documents
- **After**: 1-3 batch API calls
- **Result**: 20x faster execution, more reliable, within API limits

## Comprehensive Workflow Examples

### Create KPI Dashboard from Drive Data

```typescript
// 1. Search for source data in Drive
const files = await mcp.call('gdrive_search', {
  query: "name contains 'analytics' and type = 'spreadsheet'"
});

// 2. Read the data
const sourceData = await mcp.call('gdrive_read_file', {
  fileId: files[0].id
});

// 3. Create comprehensive dashboard in one batch operation
await mcp.call('gsheets_create_and_populate', {
  title: 'Executive KPI Dashboard',
  data: processedAnalyticsData // All rows in single API call
});
```

### Automated Report Generation

```typescript
// Create a new report document
const report = await mcp.call('gdocs_create_document', {
  title: 'Monthly Business Report',
  content: 'MONTHLY BUSINESS REPORT\n\n'
});

// Add formatted sections
await mcp.call('gdocs_set_heading', {
  documentId: report.documentId,
  startIndex: 0,
  endIndex: 23,
  headingLevel: 1
});

// Append analytics data
await mcp.call('gdocs_append_text', {
  documentId: report.documentId,
  text: '\n\nEXECUTIVE SUMMARY\n\nKey metrics for this month:\n- Revenue: $X\n- Growth: Y%\n- Customers: Z'
});

// Format the summary section
await mcp.call('gdocs_format_text', {
  documentId: report.documentId,
  startIndex: 25,
  endIndex: 41,
  format: { bold: true, fontSize: 14 }
});
```

### Business Intelligence Automation

```typescript
// Find all customer deal documents
const dealDocs = await mcp.call('gdrive_search', {
  query: "name contains 'customer deals' and type = 'document'"
});

// Read and consolidate data from multiple sources
const consolidatedData = [];
for (const doc of dealDocs) {
  const content = await mcp.call('gdrive_read_file', {fileId: doc.id});
  consolidatedData.push(...parseDealsData(content));
}

// Create master dashboard with batch operations
await mcp.call('gsheets_batch_update', {
  spreadsheetId: 'master-dashboard-id',
  updates: [
    { range: 'A1:Z100', values: consolidatedData },
    { range: 'Summary!A1:E20', values: summaryMetrics }
  ]
});
```

### Content Management Workflows

```typescript
// Search for template documents
const templates = await mcp.call('gdrive_search', {
  query: "name contains 'template' and type = 'document'"
});

// Read template content
const templateContent = await mcp.call('gdrive_read_file', {
  fileId: templates[0].id
});

// Create customized documents from template
const newDoc = await mcp.call('gdocs_create_document', {
  title: 'Customer Proposal - Acme Corp',
  content: templateContent
});

// Customize with client-specific information
await mcp.call('gdocs_replace_text', {
  documentId: newDoc.documentId,
  find: '{{CLIENT_NAME}}',
  replace: 'Acme Corporation'
});

await mcp.call('gdocs_replace_text', {
  documentId: newDoc.documentId,
  find: '{{DATE}}',
  replace: new Date().toLocaleDateString()
});
```

## Example Use Cases

### Document Management
- **Report Generation**: Create formatted reports with data from multiple sources
- **Template Processing**: Automate document creation from templates
- **Content Analysis**: Process and analyze documents at scale
- **Proposal Creation**: Generate customized proposals and contracts

### Data Operations
- **Executive Dashboards**: Combine data from multiple Drive sources into comprehensive Sheets dashboards
- **Business Intelligence**: Automate report generation from scattered documents
- **Data Consolidation**: Merge information from various Google Docs and Sheets
- **Automated Reporting**: Create recurring reports from dynamic data sources

### Workflow Automation
- **Project Management**: Aggregate status from multiple project documents
- **Content Publishing**: Coordinate content across Docs and Sheets
- **Document Workflows**: Automate document creation, formatting, and distribution
- **Cross-Platform Integration**: Bridge data between Sheets, Docs, and Drive

## Development

```bash
# Watch mode for development
npm run dev

# Build for production
npm run build

# Run authentication flow
npm run auth
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Security

- Never commit OAuth credentials to the repository
- Use environment variables or secure credential storage
- Regularly rotate your OAuth tokens
- Follow Google's API usage guidelines
- **Drive permissions**: This MCP uses read-only access to Google Drive for security
- **Docs permissions**: Full read/write access to Google Docs for document management

## Support

- [Issues](https://github.com/poqcommerce/mcp-google-workspace/issues)
- [Model Context Protocol Documentation](https://modelcontextprotocol.io/)
- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- [Google Drive API Documentation](https://developers.google.com/drive/api)
- [Google Docs API Documentation](https://developers.google.com/docs/api)