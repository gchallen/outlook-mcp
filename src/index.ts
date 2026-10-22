#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { resolveConfig } from './config.js';
import { initialize, server } from './server.js';
import * as dotenv from 'dotenv';
import 'isomorphic-fetch';

// quiet: true suppresses dotenv's startup banner. On an MCP stdio server,
// anything written to stdout that isn't JSON-RPC corrupts the protocol stream
// and breaks the connection.
dotenv.config({ quiet: true });

async function main() {
  await initialize(resolveConfig());

  const transport = new StdioServerTransport();
  await server.connect(transport);
  // Exit when the client goes away. Otherwise a pending device-code sign-in
  // keeps polling Microsoft, and the process lingers, for up to 15 minutes.
  process.stdin.on('end', () => process.exit(0));
  console.error('Outlook MCP Server running...');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
