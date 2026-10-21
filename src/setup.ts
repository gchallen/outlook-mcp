#!/usr/bin/env node
import { GraphAuthProvider, explainAuthError } from './auth.js';
import {
  GUIDE_URL,
  PACKAGE_NAME,
  discoverTenantId,
  isGuid,
  readConfigFile,
  resolveConfig,
  writeConfigFile,
} from './config.js';
import { createInterface } from 'readline/promises';
import 'isomorphic-fetch';

// Interactive first-time setup: find the user's Microsoft 365 organization from
// their email address, record the client ID of the app they registered, sign
// in once, and print how to add the server to each AI app. Every prompt has a
// flag, so an AI agent driving a terminal can run it without answering
// prompts: `setup --email a@b.edu --client-id <id> --no-login`.

const HELP = `Usage: npx -y ${PACKAGE_NAME} setup [options]

Options:
  --email <address>     Your work or school email address
  --client-id <id>      Application (client) ID of the app you registered
  --tenant-id <id>      Directory (tenant) ID (normally found from your email)
  --no-teams            Do not request Microsoft Teams permissions
  --no-login            Save settings only; sign in later from your AI app
  --help                Show this help

Step-by-step guide: ${GUIDE_URL}`;

interface Options {
  email?: string;
  clientId?: string;
  tenantId?: string;
  teams: boolean;
  login: boolean;
}

function parseArgs(argv: string[]): Options {
  const options: Options = { teams: true, login: true };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const value = () => {
      const next = argv[++i];
      if (next === undefined) throw new Error(`${arg} needs a value.`);
      return next;
    };
    if (arg === '--email') options.email = value();
    else if (arg === '--client-id') options.clientId = value();
    else if (arg === '--tenant-id') options.tenantId = value();
    else if (arg === '--no-teams') options.teams = false;
    else if (arg === '--no-login') options.login = false;
    else if (arg === '--help' || arg === '-h') {
      console.log(HELP);
      process.exit(0);
    } else throw new Error(`Unknown option ${arg}.\n\n${HELP}`);
  }
  return options;
}

const rl = process.stdin.isTTY
  ? createInterface({ input: process.stdin, output: process.stdout })
  : undefined;

async function ask(question: string, current?: string, flag?: string): Promise<string> {
  if (!rl) {
    if (current) return current;
    throw new Error(`Missing ${flag}. Pass it as an option when not running interactively.`);
  }
  const suffix = current ? ` [${current}]` : '';
  const answer = (await rl.question(`${question}${suffix}: `)).trim();
  return answer || current || '';
}

async function askYesNo(question: string, fallback: boolean): Promise<boolean> {
  if (!rl) return fallback;
  const answer = (await rl.question(`${question} ${fallback ? '[Y/n]' : '[y/N]'}: `)).trim();
  return answer ? /^y/i.test(answer) : fallback;
}

async function askSignature(): Promise<string | undefined> {
  if (!rl) return undefined;
  console.log(
    '\nOptional: paste the signature to add to your emails, then press Enter on an empty line.' +
      '\nPress Enter now to skip (you can also ask your AI to set it later).',
  );
  const lines: string[] = [];
  for (;;) {
    const line = await rl.question(lines.length ? '' : '> ');
    if (!line.trim()) break;
    lines.push(line);
  }
  return lines.length ? lines.join('\n') : undefined;
}

function printClientInstructions() {
  const npx = `npx -y ${PACKAGE_NAME}`;
  const json = JSON.stringify(
    { mcpServers: { outlook: { command: 'npx', args: ['-y', PACKAGE_NAME] } } },
    null,
    2,
  );
  console.log(`
Add the Outlook server to your AI app (then restart the app):

  Claude Code:     claude mcp add outlook --scope user -- ${npx}
  Codex:           codex mcp add outlook -- ${npx}
  Claude Desktop:  install the outlook-mcp .mcpb extension (double-click it), or add
                   the JSON below to claude_desktop_config.json
  Cursor and other apps that use an "mcpServers" JSON file:

${json
  .split('\n')
  .map((line) => '    ' + line)
  .join('\n')}

On Windows, use "command": "cmd" and "args": ["/c", "npx", "-y", "${PACKAGE_NAME}"].
Then ask your AI app to "check my Outlook connection".`);
}

async function main() {
  const options = parseArgs(process.argv.slice(3));
  const { configPath } = resolveConfig();
  const existing = readConfigFile(configPath);

  console.log('Outlook MCP setup. Settings are saved to ' + configPath + '\n');

  const email = await ask(
    'Your work or school email address',
    options.email ?? existing.email,
    '--email',
  );
  let tenantId = options.tenantId ?? existing.tenantId;
  if (!options.tenantId) {
    tenantId = await discoverTenantId(email);
    console.log(`Found your Microsoft 365 organization (tenant ${tenantId}).`);
  }

  let clientId = '';
  while (!isGuid(clientId)) {
    if (clientId)
      console.log('That does not look like a client ID (it should look like 1a2b3c4d-...).');
    clientId = await ask(
      'Application (client) ID of the app you registered',
      options.clientId ?? existing.clientId,
      '--client-id',
    );
    if (!rl && !isGuid(clientId)) throw new Error(`"${clientId}" is not a valid client ID.`);
  }

  writeConfigFile(configPath, { email, tenantId, clientId: clientId.trim(), teams: options.teams });
  console.log('Settings saved.');

  if (options.login) {
    let teams = options.teams;
    for (;;) {
      const config = resolveConfig();
      const auth = new GraphAuthProvider({
        clientId: config.clientId!,
        tenantId,
        scopes: config.scopes,
        loginScopes: config.loginScopes,
        cachePath: config.cachePath,
      });
      try {
        console.log('\nSign in with your work or school account:');
        await auth.login();
        break;
      } catch (error) {
        console.error('\nSign-in failed: ' + explainAuthError(error));
        const consent = /AADSTS(65001|90094|90095|900941)/.test(String(error));
        if (teams && consent && (await askYesNo('Try again without Teams permissions?', false))) {
          teams = false;
          writeConfigFile(configPath, { teams: false });
          continue;
        }
        process.exit(1);
      }
    }
  } else {
    console.log(
      'Skipping sign-in. Your AI app will show a sign-in link the first time you use it.',
    );
  }

  if (existing.signature === undefined) {
    const signature = await askSignature();
    if (signature) {
      writeConfigFile(configPath, { signature });
      console.log('Signature saved.');
    }
  }

  printClientInstructions();
  rl?.close();
  process.exit(0);
}

main().catch((error) => {
  console.error('Setup failed: ' + (error instanceof Error ? error.message : String(error)));
  process.exit(1);
});
