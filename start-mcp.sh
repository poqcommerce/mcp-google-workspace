#!/bin/bash
# Change to script directory
cd "$(dirname "$0")"

# Ensure proper PATH for finding node
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

# Load environment variables from .env file only if not already set
# This allows mcp.json config to take precedence
if [ -f .env ]; then
  while IFS='=' read -r key value; do
    # Skip comments and empty lines
    [[ $key =~ ^#.*$ ]] && continue
    [[ -z $key ]] && continue
    # Only set if not already in environment
    if [ -z "${!key}" ]; then
      export "$key=$value"
    fi
  done < .env
fi

# Start the MCP server
exec node dist/index.js
