# Google Sheets Batch MCP

A Model Context Protocol (MCP) server that provides efficient batch operations for Google Sheets, reducing API calls by up to 40x compared to individual cell updates.

## Features

- 🚀 **Batch Updates**: Update multiple ranges in a single API call
- 📊 **Create & Populate**: Create new sheets with data in one operation
- ➕ **Append Rows**: Efficiently add new data to existing sheets
- 🎨 **Format Cells**: Apply formatting to multiple ranges at once
- 🔐 **OAuth2 Authentication**: Secure user-based access to Google Sheets

## Quick Start

### Prerequisites

- Node.js 18+ 
- Google Cloud Project with Sheets API enabled
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
   - Enable Google Sheets API
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
   #Edit .env with your actual Google OAuth credentials
   ```

      
3. **Update Claude Desktop config:**

   ```json
   {
     "mcpServers": {
       "google-sheets-batch": {
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
    "google-sheets-batch": {
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

### `gsheets_batch_update`
Update multiple ranges in a single API call.

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

### `gsheets_create_and_populate`
Create a new sheet with initial data.

```typescript
await mcp.call('gsheets_create_and_populate', {
  title: 'My New Sheet',
  data: [['Name', 'Value'], ['Item 1', '100']]
});
```

### `gsheets_append_rows`
Add new rows to an existing sheet.

```typescript
await mcp.call('gsheets_append_rows', {
  spreadsheetId: 'your-sheet-id',
  range: 'Sheet1!A:Z',
  values: [['New', 'Data', 'Row']]
});
```

### `gsheets_format_cells`
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

## Performance Benefits

- **Before**: 40+ individual API calls for complex dashboards
- **After**: 1-3 batch API calls
- **Result**: 20x faster execution, more reliable, within API limits

## Example Use Cases

- Creating KPI dashboards from analytics data
- Bulk data imports and exports
- Automated report generation
- Data synchronization between systems

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

## Support

- [Issues](https://github.com/poqcommerce/mcp-google-drive-batch/issues)
- [Model Context Protocol Documentation](https://modelcontextprotocol.io/)
- [Google Sheets API Documentation](https://developers.google.com/sheets/api)