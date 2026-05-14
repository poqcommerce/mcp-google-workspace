# Updating the Google Workspace MCP — Windows

Steps to pull the latest version and rebuild. Takes about 2 minutes. No re-authentication required for this update.

## Steps

- Open a PowerShell or Command Prompt window
- Navigate to the repo folder (wherever you cloned it — likely something like `C:\Users\Mike\mcp-google-workspace`):
  ```
  cd C:\Users\Mike\mcp-google-workspace
  ```
- Pull the latest code from GitHub:
  ```
  git pull
  ```
- Install any new dependencies (fast if nothing changed):
  ```
  npm install
  ```
- Rebuild the TypeScript:
  ```
  npm run build
  ```
- Fully quit Claude Desktop — not just close the window. Right-click the Claude icon in the system tray (bottom-right of the Windows taskbar, next to the clock) and choose **Quit**
- Re-open Claude Desktop. The MCP server will restart automatically with the new code

## Verifying it worked

- In a new Claude conversation, ask "list the google-workspace MCP tools" or similar
- You should see **58 tools** (previously 56). New tools in this update: `gdrive_download_file`, `gslides_delete_slide` (now supports batch), `gsheets_format_dimensions`
- If the count is still 56, the rebuild didn't pick up — repeat the `npm run build` step and restart Claude Desktop again

## If something breaks

- Check for any error messages in Claude's MCP logs (Claude Desktop > Settings > Developer > MCP logs)
- If `npm install` fails, try deleting the `node_modules` folder and running `npm install` again
- If `git pull` complains about local changes, run `git status` first to see what's modified — don't just blow them away without checking
- If all else fails, message Jay
