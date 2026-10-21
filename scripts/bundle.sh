#!/bin/sh
# Build the Claude Desktop extension (outlook-mcp.mcpb): the compiled server,
# its manifest, and production dependencies only, packed with the mcpb CLI.
set -e
cd "$(dirname "$0")/.."

npm run build
rm -rf build/mcpb
mkdir -p build/mcpb
cp -R dist manifest.json package.json README.md build/mcpb/
(cd build/mcpb && npm install --omit=dev --no-package-lock --ignore-scripts --silent)
npx -y @anthropic-ai/mcpb validate build/mcpb/manifest.json
npx -y @anthropic-ai/mcpb pack build/mcpb build/outlook-mcp.mcpb
echo "Built build/outlook-mcp.mcpb"
