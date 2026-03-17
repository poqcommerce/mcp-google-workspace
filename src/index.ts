import 'dotenv/config';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

import { SheetsHandler, getSheetsToolDefinitions } from './tools/sheets.js';
import { DocsHandler, getDocsToolDefinitions } from './tools/docs.js';
import { DriveHandler, getDriveToolDefinitions } from './tools/drive.js';
import { SlidesHandler, getSlidesToolDefinitions } from './tools/slides.js';
import { AuthHandler, getAuthToolDefinitions } from './tools/auth.js';
import type { ToolResponse } from './types.js';

class GoogleWorkspaceMCP {
  private server: Server;
  private auth: OAuth2Client;

  // Domain handlers
  private sheetsHandler: SheetsHandler;
  private docsHandler: DocsHandler;
  private driveHandler: DriveHandler;
  private slidesHandler: SlidesHandler;
  private authHandler: AuthHandler;

  constructor() {
    console.error('Starting GoogleWorkspaceMCP...');
    console.error('Environment check:');
    console.error('GOOGLE_CLIENT_ID:', process.env.GOOGLE_CLIENT_ID ? 'Set' : 'Missing');
    console.error('GOOGLE_CLIENT_SECRET:', process.env.GOOGLE_CLIENT_SECRET ? 'Set' : 'Missing');
    console.error('GOOGLE_REFRESH_TOKEN:', process.env.GOOGLE_REFRESH_TOKEN ? 'Set' : 'Missing');

    this.server = new Server(
      { name: 'google-workspace', version: '1.0.0' },
      { capabilities: { tools: {} } },
    );

    try {
      const clientId = process.env.GOOGLE_CLIENT_ID;
      const clientSecret = process.env.GOOGLE_CLIENT_SECRET;

      if (!clientId || !clientSecret) {
        console.error('Warning: Google OAuth credentials not set. Some features may not work.');
        this.auth = new OAuth2Client();
      } else {
        this.auth = new OAuth2Client({
          clientId,
          clientSecret,
          redirectUri: 'http://localhost:3000/oauth/callback',
        });

        if (process.env.GOOGLE_REFRESH_TOKEN) {
          this.auth.setCredentials({
            refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
          });
        }
      }

      // Initialize Google API clients
      const sheets = google.sheets({ version: 'v4', auth: this.auth });
      const drive = google.drive({ version: 'v3', auth: this.auth });
      const docs = google.docs({ version: 'v1', auth: this.auth });
      const slidesApi = google.slides({ version: 'v1', auth: this.auth });

      // Initialize domain handlers
      this.sheetsHandler = new SheetsHandler(sheets, drive);
      this.docsHandler = new DocsHandler(docs, drive);
      this.driveHandler = new DriveHandler(drive);
      this.slidesHandler = new SlidesHandler(slidesApi, drive);
      this.authHandler = new AuthHandler(this.auth);

      console.error('Google Workspace clients initialized successfully');

      this.setupToolHandlers();
      console.error('Tool handlers set up successfully');
    } catch (error) {
      console.error('Error during initialization:', error);
      throw error;
    }
  }

  private setupToolHandlers() {
    // Aggregate all tool definitions
    const allTools = [
      ...getSheetsToolDefinitions(),
      ...getDocsToolDefinitions(),
      ...getSlidesToolDefinitions(),
      ...getAuthToolDefinitions(),
      ...getDriveToolDefinitions(),
    ];

    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: allTools,
    }));

    // Ordered list of handlers — first match wins
    const handlers = [
      this.sheetsHandler,
      this.docsHandler,
      this.slidesHandler,
      this.authHandler,
      this.driveHandler,
    ];

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        const { name } = request.params;
        const args = request.params.arguments;

        for (const handler of handlers) {
          const result: ToolResponse | null = await handler.handleTool(name, args);
          if (result !== null) return result;
        }

        throw new Error(`Unknown tool: ${name}`);
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

// Start the server
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

if (process.argv[1] === __filename) {
  console.error('Starting MCP server...');
  try {
    const server = new GoogleWorkspaceMCP();
    server.run().catch((error) => {
      console.error('Server run error:', error);
      process.exit(1);
    });
  } catch (error) {
    console.error('Server initialization error:', error);
    process.exit(1);
  }
}

export { GoogleWorkspaceMCP };
