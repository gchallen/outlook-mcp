#!/usr/bin/env node
import { GraphAuthProvider, explainAuthError } from './auth.js';
import { PACKAGE_NAME, discoverTenantId, resolveConfig } from './config.js';
import * as dotenv from 'dotenv';
import 'isomorphic-fetch';

dotenv.config({ quiet: true });

// Re-run the device-code sign-in with the existing configuration, e.g. after
// the cached sign-in expires. First-time configuration is `setup`'s job.
async function main() {
  const config = resolveConfig();
  if (!config.clientId) {
    console.error(`No client ID is configured. Run \`npx -y ${PACKAGE_NAME} setup\` first.`);
    process.exit(1);
  }

  const auth = new GraphAuthProvider({
    clientId: config.clientId,
    tenantId: config.tenantId ?? (config.email ? await discoverTenantId(config.email) : undefined),
    scopes: config.scopes,
    loginScopes: config.loginScopes,
    cachePath: config.cachePath,
  });

  console.error('Starting interactive sign-in...');
  await auth.login();
  console.error('Done. You can now use the outlook-mcp server.');
  process.exit(0);
}

main().catch((error) => {
  console.error('Login failed:', explainAuthError(error));
  process.exit(1);
});
