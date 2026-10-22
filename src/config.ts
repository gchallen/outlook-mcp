import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DEFAULT_SCOPES, TEAMS_SCOPES } from './auth.js';

/** The npm package name, as people type it in `npx -y <name>`. */
export const PACKAGE_NAME = '@gchallen/outlook-mcp';

/** The AI guided-setup walkthrough. */
export const GUIDE_URL = 'https://github.com/gchallen/outlook-mcp/blob/main/docs/GUIDED-SETUP.md';

// Per-user settings live in ~/.outlook-mcp/config.json, written by `setup` and
// by the set_signature tool, so every MCP client (Claude, Codex, Cursor, ...)
// can launch the server with no arguments or environment at all. Environment
// variables still override the file for the connection settings, which keeps
// older env-only installs (and the .mcpb bundle, which passes user_config as
// env) working.

export interface FileConfig {
  clientId?: string;
  tenantId?: string;
  email?: string;
  signature?: string;
  teams?: boolean;
  extraScopes?: string[];
}

export interface ResolvedConfig {
  clientId?: string;
  tenantId?: string;
  email?: string;
  signature?: string;
  /** Scopes requested on every silent token refresh. */
  scopes: string[];
  /** Scopes requested at interactive sign-in (adds Teams and extras). */
  loginScopes: string[];
  cachePath: string;
  configPath: string;
}

// The tenant id Microsoft reports for personal (outlook.com/hotmail) accounts.
const CONSUMER_TENANT_ID = '9188040d-6c67-4c5b-b112-36a304b66dad';

export function configHome(env: NodeJS.ProcessEnv = process.env): string {
  return env.OUTLOOK_MCP_HOME || path.join(os.homedir(), '.outlook-mcp');
}

export function configPath(env: NodeJS.ProcessEnv = process.env): string {
  return path.join(configHome(env), 'config.json');
}

export function readConfigFile(file: string): FileConfig {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf-8')) as FileConfig;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return {};
    throw new Error(`Could not read ${file}: ${(error as Error).message}`);
  }
}

/** Merge `updates` into the config file, creating it (and its folder) if needed. */
export function writeConfigFile(file: string, updates: FileConfig): FileConfig {
  const merged = { ...readConfigFile(file), ...updates };
  fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
  fs.writeFileSync(file, JSON.stringify(merged, null, 2) + '\n', { mode: 0o600 });
  return merged;
}

// Clients that template optional settings into env (the .mcpb bundle does)
// pass unset ones as empty strings, or possibly as the unfilled placeholder
// itself ("${user_config.signature}"). Both must read as "not set", or the
// placeholder would be appended to every email as a signature.
function envValue(env: NodeJS.ProcessEnv, name: string): string | undefined {
  const value = env[name]?.trim();
  return value && !/^\$\{[^}]*\}$/.test(value) ? value : undefined;
}

function list(value?: string): string[] | undefined {
  return value
    ?.split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

export function resolveConfig(env: NodeJS.ProcessEnv = process.env): ResolvedConfig {
  const file = configPath(env);
  const fromFile = readConfigFile(file);

  const teamsEnv = envValue(env, 'OUTLOOK_TEAMS');
  const teams = teamsEnv ? !/^(0|false|no|off)$/i.test(teamsEnv) : fromFile.teams !== false;
  const graphScopes = list(envValue(env, 'GRAPH_SCOPES'));
  const extra = list(envValue(env, 'EXTRA_GRAPH_SCOPES')) ?? fromFile.extraScopes ?? [];

  // The signature is the one setting the file wins on: set_signature saves it
  // there, and an env value would otherwise silently undo that on restart.
  // Env values may spell line breaks as a literal "\n".
  const signature =
    fromFile.signature !== undefined
      ? fromFile.signature
      : envValue(env, 'OUTLOOK_SIGNATURE')?.replace(/\\n/g, '\n');

  return {
    clientId: envValue(env, 'AZURE_CLIENT_ID') ?? fromFile.clientId,
    tenantId: envValue(env, 'AZURE_TENANT_ID') ?? fromFile.tenantId,
    email: envValue(env, 'OUTLOOK_EMAIL') ?? fromFile.email,
    signature,
    scopes: graphScopes ?? DEFAULT_SCOPES,
    loginScopes: [
      ...(graphScopes ?? [...DEFAULT_SCOPES, ...(teams ? TEAMS_SCOPES : [])]),
      ...extra,
    ],
    cachePath:
      envValue(env, 'OUTLOOK_TOKEN_CACHE') ?? path.join(configHome(env), 'token-cache.json'),
    configPath: file,
  };
}

/**
 * Find the Microsoft 365 tenant for an email address from its domain's public
 * OpenID metadata, so nobody has to dig the "Directory (tenant) ID" out of the
 * Entra portal.
 */
export async function discoverTenantId(email: string): Promise<string> {
  const domain = email.split('@').pop()?.trim().toLowerCase();
  if (!domain || !domain.includes('.')) {
    throw new Error(`"${email}" does not look like an email address.`);
  }
  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(domain)}/v2.0/.well-known/openid-configuration`,
  );
  if (!response.ok) {
    throw new Error(
      `${domain} does not appear to use Microsoft 365 (Microsoft returned ${response.status}).`,
    );
  }
  const { issuer } = (await response.json()) as { issuer?: string };
  const tenantId = issuer?.match(/[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}/i)?.[0];
  if (!tenantId) throw new Error(`Could not find a tenant id for ${domain}.`);
  if (tenantId === CONSUMER_TENANT_ID) {
    throw new Error(
      `${domain} is a personal Microsoft account domain. This server needs a work or school account.`,
    );
  }
  return tenantId;
}

const GUID = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;

export function isGuid(value: string): boolean {
  return GUID.test(value.trim());
}
