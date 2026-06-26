#!/usr/bin/env node
import { GraphAuthProvider, DEFAULT_SCOPES } from './auth.js';
import * as dotenv from 'dotenv';
import 'isomorphic-fetch';

dotenv.config();

async function main() {
  const clientId = process.env.AZURE_CLIENT_ID;
  if (!clientId) {
    console.error('AZURE_CLIENT_ID is not set. Create a .env file (see .env.example).');
    process.exit(1);
  }

  const auth = new GraphAuthProvider({
    clientId,
    tenantId: process.env.AZURE_TENANT_ID,
    scopes: process.env.GRAPH_SCOPES?.split(',') ?? DEFAULT_SCOPES,
    cachePath: process.env.OUTLOOK_TOKEN_CACHE,
  });

  console.error('Starting interactive sign-in...');
  await auth.login();
  console.error('Done. You can now use the outlook-mcp server.');
  process.exit(0);
}

main().catch((error) => {
  console.error('Login failed:', error instanceof Error ? error.message : error);
  process.exit(1);
});
