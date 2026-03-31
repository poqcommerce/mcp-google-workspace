import { OAuth2Client } from 'google-auth-library';
import type { ToolDefinition, ToolResponse } from '../types.js';
import { textResponse, errorResponse } from '../utils.js';

export const OAUTH_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/presentations',
  'https://www.googleapis.com/auth/drive.activity.readonly',
];

// ── Tool definitions ───────────────────────────────────────────────────────────

export function getAuthToolDefinitions(): ToolDefinition[] {
  return [
    {
      name: 'gsheets_get_auth_url',
      description: 'Get Google OAuth authorization URL for initial setup',
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'gsheets_set_auth_code',
      description: 'Exchange authorization code for access tokens',
      inputSchema: {
        type: 'object',
        properties: {
          code: { type: 'string', description: 'Authorization code from Google OAuth flow' },
        },
        required: ['code'],
      },
    },
  ];
}

// ── Handler class ──────────────────────────────────────────────────────────────

export class AuthHandler {
  constructor(private auth: OAuth2Client) {}

  /** Route a tool call to the appropriate handler. Returns null if not handled. */
  async handleTool(name: string, args: any): Promise<ToolResponse | null> {
    switch (name) {
      case 'gsheets_get_auth_url':
        return this.handleGetAuthUrl();
      case 'gsheets_set_auth_code': {
        const code = (args as any)?.code;
        if (!code) throw new Error('Authorization code is required');
        return this.handleSetAuthCode(code);
      }
      default:
        return null;
    }
  }

  // ── Handlers ───────────────────────────────────────────────────────────────

  private async handleGetAuthUrl(): Promise<ToolResponse> {
    try {
      const url = this.auth.generateAuthUrl({
        access_type: 'offline',
        scope: OAUTH_SCOPES,
        prompt: 'consent',
      });

      return textResponse(
        `Please visit this URL to authorize the application:\n\n${url}\n\nAfter authorizing, copy the authorization code and use the gsheets_set_auth_code tool.`,
      );
    } catch (error) {
      return errorResponse('generating auth URL', error);
    }
  }

  private async handleSetAuthCode(code: string): Promise<ToolResponse> {
    try {
      const { tokens } = await this.auth.getToken(code);
      this.auth.setCredentials(tokens);

      return textResponse(
        `Authorization successful! Please save this refresh token in your environment:\n\nGOOGLE_REFRESH_TOKEN=${tokens.refresh_token}\n\nRestart the MCP server with this environment variable set.`,
      );
    } catch (error) {
      return errorResponse('exchanging code for tokens', error);
    }
  }
}
