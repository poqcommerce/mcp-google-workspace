# Google Workspace Batch MCP

A comprehensive Model Context Protocol (MCP) server that provides efficient batch operations for Google Sheets and full Google Drive access, reducing API calls by up to 40x compared to individual cell updates.

## Features

- 🚀 **Batch Sheets Operations**: Update multiple ranges in a single API call
- 📊 **Create & Populate**: Create new sheets with data in one operation
- ➕ **Append Rows**: Efficiently add new data to existing sheets
- 🎨 **Format Cells**: Apply formatting to multiple ranges at once
- 🔍 **Google Drive Search**: Find files with powerful query syntax
- 📖 **Drive File Reading**: Read Google Docs, Sheets, and text files
- 📋 **File Metadata**: Get detailed information about Drive files
- 🔐 **OAuth2 Authentication**: Secure user-based access to Google Workspace

## Quick Start

### Prerequisites

- Node.js 18+ 
- Google Cloud Project with Sheets API and Drive API enabled
- Claude Desktop or compatible MCP client

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/poqcommerce/mcp-google-drive-batch.git
   cd mcp-google-drive-batch
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
   - Enable Google Sheets API and Google Drive API
   - Create OAuth 2.0 Client ID (Desktop Application)
   - Copy Client ID and Secret

2. **Run the authentication flow:**
   ```bash
   GOOGLE_CLIENT_ID="your-id" GOOGLE_CLIENT_SECRET="your-secret" npm run auth
   ```

3. **Save the refresh token** for use in your MCP client configuration.

### Claude Desktop Configuration

You have two options for configuration:

#### **Configuration with Shell Script (Recommended)**

Your setup uses a shell script approach for better security:

1. **Clone and setup:**
   ```bash
   git clone https://github.com/poqcommerce/mcp-google-drive-batch.git
   cd mcp-google-drive-batch
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
       "google-workspace-batch": {
         "command": "/path/to/mcp-google-drive-batch/start-mcp.sh"
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
    "google-workspace-batch": {
      "command": "node",
      "args": ["/path/to/mcp-google-drive-batch/dist/index.js"],
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

## Performance Benefits

- **Before**: 40+ individual API calls for complex dashboards
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

## Example Use Cases

- **Executive Dashboards**: Combine data from multiple Drive sources into comprehensive Sheets dashboards
- **Business Intelligence**: Automate report generation from scattered documents
- **Data Consolidation**: Merge information from various Google Docs and Sheets
- **Content Analysis**: Process and analyze documents at scale
- **Automated Reporting**: Create recurring reports from dynamic data sources
- **Project Management**: Aggregate status from multiple project documents

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

## Support

- [Issues](https://github.com/poqcommerce/mcp-google-drive-batch/issues)
- [Model Context Protocol Documentation](https://modelcontextprotocol.io/)
- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- [Google Drive API Documentation](https://developers.google.com/drive/api)