#!/bin/bash
# Change to script directory
cd "$(dirname "$0")"

# Load environment variables
set -a  # automatically export all variables
source .env
set +a  # turn off automatic export

# Start the MCP server
node dist/index.js
