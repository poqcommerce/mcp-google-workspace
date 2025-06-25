import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { google, sheets_v4 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';

interface BatchUpdateRequest {
  spreadsheetId: string;
  updates: {
    range: string;
    values: any[][];
  }[];
}

interface CreateSheetRequest {
  title: string;
  sheetTitle?: string;
}

class GoogleSheetsBatchMCP {
  private server: Server;
  private auth: OAuth2Client;
  private sheets: sheets_v4.Sheets;

  constructor() {
    // Add debugging
    console.error('Starting GoogleSheetsBatchMCP...');
    console.error('Environment check:');
    console.error('GOOGLE_CLIENT_ID:', process.env.GOOGLE_CLIENT_ID ? 'Set' : 'Missing');
    console.error('GOOGLE_CLIENT_SECRET:', process.env.GOOGLE_CLIENT_SECRET ? 'Set' : 'Missing');
    console.error('GOOGLE_REFRESH_TOKEN:', process.env.GOOGLE_REFRESH_TOKEN ? 'Set' : 'Missing');

    this.server = new Server(
      {
        name: 'google-sheets-batch',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    try {
      // Initialize OAuth2 Client with error handling
      const clientId = process.env.GOOGLE_CLIENT_ID;
      const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
      
      if (!clientId || !clientSecret) {
        console.error('Warning: Google OAuth credentials not set. Some features may not work.');
        // Create a minimal auth client that will fail gracefully
        this.auth = new OAuth2Client();
      } else {
        this.auth = new OAuth2Client({
          clientId,
          clientSecret,
          redirectUri: 'urn:ietf:wg:oauth:2.0:oob', // For desktop apps
        });

        // Set the refresh token if available
        if (process.env.GOOGLE_REFRESH_TOKEN) {
          this.auth.setCredentials({
            refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
          });
        }
      }

      this.sheets = google.sheets({ version: 'v4', auth: this.auth });
      console.error('Google Sheets client initialized successfully');

      this.setupToolHandlers();
      console.error('Tool handlers set up successfully');
      
    } catch (error) {
      console.error('Error during initialization:', error);
      throw error;
    }
  }

  // Add a method to get the authorization URL for initial setup
  private getAuthUrl(): string {
    const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
    
    return this.auth.generateAuthUrl({
      access_type: 'offline',
      scope: scopes,
      prompt: 'consent', // Forces refresh token generation
    });
  }

  // Add a method to exchange authorization code for tokens
  private async getTokensFromCode(code: string) {
    const { tokens } = await this.auth.getToken(code);
    this.auth.setCredentials(tokens);
    return tokens;
  }

  private validateBatchUpdateArgs(args: any): BatchUpdateRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.spreadsheetId || typeof args.spreadsheetId !== 'string') {
      throw new Error('Invalid spreadsheetId: expected non-empty string');
    }
    if (!Array.isArray(args.updates)) {
      throw new Error('Invalid updates: expected array');
    }
    
    const updates = args.updates.map((update: any, index: number) => {
      if (!update || typeof update !== 'object') {
        throw new Error(`Invalid update at index ${index}: expected object`);
      }
      if (!update.range || typeof update.range !== 'string') {
        throw new Error(`Invalid range at index ${index}: expected non-empty string`);
      }
      if (!Array.isArray(update.values)) {
        throw new Error(`Invalid values at index ${index}: expected 2D array`);
      }
      return {
        range: update.range,
        values: update.values
      };
    });

    return {
      spreadsheetId: args.spreadsheetId,
      updates
    };
  }

  private validateCreateArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.title || typeof args.title !== 'string') {
      throw new Error('Invalid title: expected non-empty string');
    }
    return {
      title: args.title,
      sheetTitle: args.sheetTitle || 'Sheet1',
      data: args.data || []
    };
  }

  private validateAppendArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.spreadsheetId || typeof args.spreadsheetId !== 'string') {
      throw new Error('Invalid spreadsheetId: expected non-empty string');
    }
    if (!args.range || typeof args.range !== 'string') {
      throw new Error('Invalid range: expected non-empty string');
    }
    if (!Array.isArray(args.values)) {
      throw new Error('Invalid values: expected 2D array');
    }
    return {
      spreadsheetId: args.spreadsheetId,
      range: args.range,
      values: args.values
    };
  }

  private validateFormatArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.spreadsheetId || typeof args.spreadsheetId !== 'string') {
      throw new Error('Invalid spreadsheetId: expected non-empty string');
    }
    if (!Array.isArray(args.requests)) {
      throw new Error('Invalid requests: expected array');
    }
    return {
      spreadsheetId: args.spreadsheetId,
      requests: args.requests
    };
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'gsheets_batch_update',
          description: 'Update multiple ranges in a Google Sheet in a single API call',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: {
                type: 'string',
                description: 'The ID of the spreadsheet to update',
              },
              updates: {
                type: 'array',
                description: 'Array of range updates to perform',
                items: {
                  type: 'object',
                  properties: {
                    range: {
                      type: 'string',
                      description: 'A1 notation range (e.g., "Sheet1!A1:C3")',
                    },
                    values: {
                      type: 'array',
                      description: '2D array of values to insert',
                      items: {
                        type: 'array',
                        items: { type: 'string' }
                      }
                    },
                  },
                  required: ['range', 'values'],
                },
              },
            },
            required: ['spreadsheetId', 'updates'],
          },
        },
        {
          name: 'gsheets_create_and_populate',
          description: 'Create a new Google Sheet and populate it with data',
          inputSchema: {
            type: 'object',
            properties: {
              title: {
                type: 'string',
                description: 'Title for the new spreadsheet',
              },
              sheetTitle: {
                type: 'string',
                description: 'Title for the first sheet tab (optional)',
              },
              data: {
                type: 'array',
                description: '2D array of data to populate the sheet',
                items: {
                  type: 'array',
                  items: { type: 'string' }
                }
              },
            },
            required: ['title', 'data'],
          },
        },
        {
          name: 'gsheets_append_rows',
          description: 'Append rows to the end of a sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: {
                type: 'string',
                description: 'The ID of the spreadsheet',
              },
              range: {
                type: 'string',
                description: 'Range to append to (e.g., "Sheet1!A:Z")',
              },
              values: {
                type: 'array',
                description: '2D array of values to append',
                items: {
                  type: 'array',
                  items: { type: 'string' }
                }
              },
            },
            required: ['spreadsheetId', 'range', 'values'],
          },
        },
        {
          name: 'gsheets_format_cells',
          description: 'Apply formatting to cell ranges',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: {
                type: 'string',
                description: 'The ID of the spreadsheet',
              },
              requests: {
                type: 'array',
                description: 'Array of formatting requests',
                items: {
                  type: 'object',
                  properties: {
                    range: {
                      type: 'string',
                      description: 'A1 notation range to format',
                    },
                    format: {
                      type: 'object',
                      description: 'Formatting options (bold, backgroundColor, etc.)',
                    },
                  },
                },
              },
            },
            required: ['spreadsheetId', 'requests'],
          },
        },
        {
          name: 'gsheets_get_auth_url',
          description: 'Get Google OAuth authorization URL for initial setup',
          inputSchema: {
            type: 'object',
            properties: {},
          },
        },
        {
          name: 'gsheets_set_auth_code',
          description: 'Exchange authorization code for access tokens',
          inputSchema: {
            type: 'object',
            properties: {
              code: {
                type: 'string',
                description: 'Authorization code from Google OAuth flow',
              },
            },
            required: ['code'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        switch (request.params.name) {
          case 'gsheets_batch_update':
            const batchArgs = this.validateBatchUpdateArgs(request.params.arguments);
            return this.handleBatchUpdate(batchArgs);
          
          case 'gsheets_create_and_populate':
            const createArgs = this.validateCreateArgs(request.params.arguments);
            return this.handleCreateAndPopulate(createArgs);
          
          case 'gsheets_append_rows':
            const appendArgs = this.validateAppendArgs(request.params.arguments);
            return this.handleAppendRows(appendArgs);
          
          case 'gsheets_format_cells':
            const formatArgs = this.validateFormatArgs(request.params.arguments);
            return this.handleFormatCells(formatArgs);
          
          case 'gsheets_get_auth_url':
            return this.handleGetAuthUrl();
          
          case 'gsheets_set_auth_code':
            const code = (request.params.arguments as any)?.code;
            if (!code) {
              throw new Error('Authorization code is required');
            }
            return this.handleSetAuthCode(code);
          
          default:
            throw new Error(`Unknown tool: ${request.params.name}`);
        }
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  private async handleGetAuthUrl() {
    try {
      const url = this.getAuthUrl();
      return {
        content: [
          {
            type: 'text',
            text: `Please visit this URL to authorize the application:\n\n${url}\n\nAfter authorizing, copy the authorization code and use the gsheets_set_auth_code tool.`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error generating auth URL: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleSetAuthCode(code: string) {
    try {
      const tokens = await this.getTokensFromCode(code);
      return {
        content: [
          {
            type: 'text',
            text: `Authorization successful! Please save this refresh token in your environment:\n\nGOOGLE_REFRESH_TOKEN=${tokens.refresh_token}\n\nRestart the MCP server with this environment variable set.`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error exchanging code for tokens: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleBatchUpdate(args: BatchUpdateRequest) {
    try {
      const batchUpdateRequest: sheets_v4.Schema$BatchUpdateValuesRequest = {
        valueInputOption: 'USER_ENTERED',
        data: args.updates.map(update => ({
          range: update.range,
          values: update.values,
        })),
      };

      const response = await this.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: batchUpdateRequest,
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              updatedCells: response.data.totalUpdatedCells,
              updatedRows: response.data.totalUpdatedRows,
              updatedColumns: response.data.totalUpdatedColumns,
              updatedSheets: response.data.totalUpdatedSheets,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error updating spreadsheet: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleCreateAndPopulate(args: {
    title: string;
    sheetTitle: string;
    data: any[][];
  }) {
    try {
      // Create new spreadsheet
      const createResponse = await this.sheets.spreadsheets.create({
        requestBody: {
          properties: {
            title: args.title,
          },
          sheets: [
            {
              properties: {
                title: args.sheetTitle || 'Sheet1',
              },
            },
          ],
        },
      });

      const spreadsheetId = createResponse.data.spreadsheetId!;
      const sheetName = args.sheetTitle || 'Sheet1';

      // Populate with data
      if (args.data && args.data.length > 0) {
        await this.sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: args.data,
          },
        });
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              spreadsheetId,
              url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
              rowsAdded: args.data?.length || 0,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error creating spreadsheet: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleAppendRows(args: {
    spreadsheetId: string;
    range: string;
    values: any[][];
  }) {
    try {
      const response = await this.sheets.spreadsheets.values.append({
        spreadsheetId: args.spreadsheetId,
        range: args.range,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: args.values,
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              updatedRange: response.data.updates?.updatedRange,
              updatedRows: response.data.updates?.updatedRows,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error appending rows: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleFormatCells(args: {
    spreadsheetId: string;
    requests: any[];
  }) {
    try {
      const requests = args.requests.map((req: any) => {
        // Convert A1 notation to GridRange
        const range = this.a1ToGridRange(req.range);
        
        return {
          repeatCell: {
            range,
            cell: {
              userEnteredFormat: req.format,
            },
            fields: Object.keys(req.format).map(key => `userEnteredFormat.${key}`).join(','),
          },
        };
      });

      const response = await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: {
          requests,
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              appliedFormats: requests.length,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error formatting cells: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private a1ToGridRange(a1Range: string): sheets_v4.Schema$GridRange {
    // Simple A1 to GridRange conversion
    // This is a simplified version - you'd want a more robust parser
    const match = a1Range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`Invalid A1 range: ${a1Range}`);
    }

    const [, startCol, startRow, endCol, endRow] = match;
    
    return {
      startRowIndex: parseInt(startRow) - 1,
      endRowIndex: parseInt(endRow),
      startColumnIndex: this.columnToIndex(startCol),
      endColumnIndex: this.columnToIndex(endCol) + 1,
    };
  }

  private columnToIndex(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1;
  }

  async run() {
    try {
      console.error('Starting server transport...');
      const transport = new StdioServerTransport();
      await this.server.connect(transport);
      console.error('Server connected successfully');
    } catch (error) {
      console.error('Error starting server:', error);
      process.exit(1);
    }
  }
}

// Usage example for your KPI dashboard
const exampleUsage = `
// Instead of 40+ individual calls, you could do:

await mcp.call('gsheets_batch_update', {
  spreadsheetId: '1D-FzyQrGLwOSAzcL9z-Sk3KCuWGLhZRSqfMNK-xmTi0',
  updates: [
    {
      range: 'A1:B13',
      values: [
        ['Mama\\'s and Papas - Top KPIs Dashboard'],
        ['Data Period: May 25, 2025 - June 23, 2025 (30 days)'],
        [''],
        ['EXECUTIVE SUMMARY'],
        ['Metric', 'Value'],
        ['Total Sessions', '2,287,839'],
        ['Total Transactions', '27,565'],
        ['Total Revenue (£)', '£4,834,482'],
        ['Conversion Rate', '1.20%'],
        ['Average Order Value (£)', '£207.24'],
        ['New Users', '827,212'],
        ['Active Users', '1,753,621'],
        ['Sessions per User', '1.23']
      ]
    },
    {
      range: 'A15:F25',
      values: [
        ['PERFORMANCE BY DEVICE & OS'],
        ['Device', 'OS', 'Sessions', 'Transactions', 'Revenue (£)', 'Conversion Rate'],
        ['Mobile', 'iOS', '1,495,172', '19,046', '£3,115,943', '1.27%'],
        ['Mobile', 'Android', '421,794', '4,386', '£665,205', '1.04%'],
        ['Desktop', 'Windows', '164,594', '2,139', '£591,662', '1.30%'],
        // ... etc
      ]
    }
  ]
});

// This would reduce 40+ calls to just 1-3 calls!
`;

// Start the server with error handling
// ES module way to check if this is the main module
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

if (process.argv[1] === __filename) {
  console.error('Starting MCP server...');
  try {
    const server = new GoogleSheetsBatchMCP();
    server.run().catch((error) => {
      console.error('Server run error:', error);
      process.exit(1);
    });
  } catch (error) {
    console.error('Server initialization error:', error);
    process.exit(1);
  }
}

export { GoogleSheetsBatchMCP };