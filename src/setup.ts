#!/usr/bin/env node

/**
 * Interactive setup script that writes the Claude Desktop config file.
 * Reads credentials from .env and generates the correct JSON config
 * for the user's platform — no manual JSON editing needed.
 *
 * Usage: npm run setup
 */

import 'dotenv/config';
import { createInterface } from 'readline';
import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'fs';
import { join, resolve } from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const projectRoot = resolve(__dirname, '..');

const rl = createInterface({ input: process.stdin, output: process.stdout });
const ask = (q: string): Promise<string> =>
  new Promise((res) => rl.question(q, (a) => res(a.trim())));

async function setup() {
  console.log('\n=== Google Workspace MCP — Setup ===\n');

  // Detect platform
  const platform = process.platform; // 'win32', 'darwin', 'linux'
  const isWindows = platform === 'win32';

  // Determine config path
  let configDir: string;
  let configPath: string;

  if (isWindows) {
    configDir = join(process.env.APPDATA || '', 'Claude');
    configPath = join(configDir, 'claude_desktop_config.json');
  } else {
    const home = process.env.HOME || '';
    configDir = join(home, 'Library', 'Application Support', 'Claude');
    configPath = join(configDir, 'claude_desktop_config.json');
  }

  console.log(`Platform: ${isWindows ? 'Windows' : 'macOS/Linux'}`);
  console.log(`Config file: ${configPath}\n`);

  // Read credentials from .env or environment
  let clientId = process.env.GOOGLE_CLIENT_ID || '';
  let clientSecret = process.env.GOOGLE_CLIENT_SECRET || '';
  let refreshToken = process.env.GOOGLE_REFRESH_TOKEN || '';

  // Prompt for any missing values
  if (!clientId) {
    clientId = await ask('GOOGLE_CLIENT_ID: ');
  } else {
    console.log(`GOOGLE_CLIENT_ID: ${clientId.slice(0, 20)}... (from .env)`);
  }

  if (!clientSecret) {
    clientSecret = await ask('GOOGLE_CLIENT_SECRET: ');
  } else {
    console.log(`GOOGLE_CLIENT_SECRET: ${clientSecret.slice(0, 10)}... (from .env)`);
  }

  if (!refreshToken) {
    refreshToken = await ask('GOOGLE_REFRESH_TOKEN (run "npm run auth" first if blank): ');
  } else {
    console.log(`GOOGLE_REFRESH_TOKEN: ${refreshToken.slice(0, 15)}... (from .env)`);
  }

  if (!clientId || !clientSecret) {
    console.error('\nError: Client ID and Secret are required. Add them to .env or enter them above.');
    process.exit(1);
  }

  // Build the MCP server entry
  const startScript = isWindows ? 'start-mcp.cmd' : 'start-mcp.sh';
  const command = resolve(projectRoot, startScript).replace(/\\/g, '/');

  const mcpEntry: Record<string, unknown> = {
    command: command,
    env: {
      GOOGLE_CLIENT_ID: clientId,
      GOOGLE_CLIENT_SECRET: clientSecret,
      ...(refreshToken ? { GOOGLE_REFRESH_TOKEN: refreshToken } : {}),
    },
  };

  // Read existing config or start fresh
  let config: Record<string, unknown> = {};
  if (existsSync(configPath)) {
    try {
      config = JSON.parse(readFileSync(configPath, 'utf-8'));
      console.log('\nExisting config found — adding google-workspace server.\n');
    } catch {
      console.log('\nExisting config file is invalid JSON — will back up and replace.\n');
      const backupPath = configPath + '.backup';
      writeFileSync(backupPath, readFileSync(configPath));
      console.log(`Backed up to: ${backupPath}`);
    }
  }

  // Merge
  if (!config.mcpServers || typeof config.mcpServers !== 'object') {
    config.mcpServers = {};
  }
  (config.mcpServers as Record<string, unknown>)['google-workspace'] = mcpEntry;

  // Write
  mkdirSync(configDir, { recursive: true });
  writeFileSync(configPath, JSON.stringify(config, null, 2), 'utf-8');

  console.log(`Config written to: ${configPath}\n`);
  console.log('Next steps:');
  if (!refreshToken) {
    console.log('  1. Run "npm run auth" to get your refresh token');
    console.log('  2. Run "npm run setup" again to update the config');
    console.log('  3. Restart Claude Desktop');
  } else {
    console.log('  1. Restart Claude Desktop');
    console.log('  2. Try asking Claude: "Search my Google Drive for recent documents"');
  }
  console.log('');

  rl.close();
}

setup().catch((err) => {
  console.error('Setup failed:', err);
  process.exit(1);
});
