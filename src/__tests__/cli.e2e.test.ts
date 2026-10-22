import { execFileSync, spawn, spawnSync, type ChildProcess } from 'child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { createInterface } from 'readline';

// End-to-end tests of the compiled package, run the way people run it: the
// `outlook-mcp` binary (dist/cli.js) in a fresh home folder and an empty
// working directory, so neither your own ~/.outlook-mcp nor this repo's .env
// can leak in. Tenant discovery is stubbed by fixtures/fake-fetch.cjs; nothing
// here signs in to Microsoft.

const ROOT = join(__dirname, '../..');
const CLI = join(ROOT, 'dist/cli.js');
const FAKE_FETCH = join(__dirname, 'fixtures/fake-fetch.cjs');
const CLIENT_ID = '11111111-2222-4333-8444-555555555555';
const TENANT_ID = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';

let home: string;
let cwd: string;

function env(extra: Record<string, string> = {}): NodeJS.ProcessEnv {
  return {
    PATH: process.env.PATH,
    HOME: home,
    USERPROFILE: home,
    OUTLOOK_MCP_HOME: join(home, '.outlook-mcp'),
    FAKE_TENANT_ID: TENANT_ID,
    FAKE_FETCH_LOG: join(home, 'fetches.log'),
    ...extra,
  };
}

/** Run the CLI to completion without a terminal (stdin is not a TTY). */
function run(args: string[], extra: Record<string, string> = {}) {
  const result = spawnSync('node', ['--require', FAKE_FETCH, CLI, ...args], {
    cwd,
    env: env(extra),
    encoding: 'utf-8',
    input: '',
    timeout: 20_000,
  });
  return { status: result.status, stdout: result.stdout, stderr: result.stderr };
}

const configFile = () => join(home, '.outlook-mcp', 'config.json');
const savedConfig = () => JSON.parse(readFileSync(configFile(), 'utf-8'));
const fetches = () =>
  existsSync(join(home, 'fetches.log'))
    ? readFileSync(join(home, 'fetches.log'), 'utf-8').trim().split('\n')
    : [];

beforeAll(() => {
  // The tests run the compiled output, so compile first.
  execFileSync('npx', ['tsc'], { cwd: ROOT, stdio: 'pipe' });
}, 120_000);

beforeEach(() => {
  home = mkdtempSync(join(tmpdir(), 'outlook-mcp-home-'));
  cwd = mkdtempSync(join(tmpdir(), 'outlook-mcp-cwd-'));
});

afterEach(() => {
  rmSync(home, { recursive: true, force: true });
  rmSync(cwd, { recursive: true, force: true });
});

describe('outlook-mcp command', () => {
  it('starts with a shebang so npx can run it', () => {
    expect(readFileSync(CLI, 'utf-8').split('\n')[0]).toBe('#!/usr/bin/env node');
  });

  it('rejects unknown commands', () => {
    const { status, stderr } = run(['frobnicate']);
    expect(status).toBe(1);
    expect(stderr).toContain('Unknown command "frobnicate"');
  });

  it('login explains that setup comes first', () => {
    const { status, stderr } = run(['login'], { FAKE_FETCH: 'forbid' });
    expect(status).toBe(1);
    expect(stderr).toContain('No client ID is configured');
    expect(stderr).toContain('npx -y @gchallen/outlook-mcp setup');
  });
});

describe('setup', () => {
  it('prints help', () => {
    const { status, stdout } = run(['setup', '--help']);
    expect(status).toBe(0);
    expect(stdout).toContain('Usage: npx -y @gchallen/outlook-mcp setup [options]');
    expect(stdout).toContain('--no-teams');
    expect(stdout).toContain('GUIDED-SETUP.md');
  });

  it('rejects unknown options and options missing a value', () => {
    const unknown = run(['setup', '--frob']);
    expect(unknown.status).toBe(1);
    expect(unknown.stderr).toContain('Unknown option --frob.');

    const missing = run(['setup', '--email']);
    expect(missing.status).toBe(1);
    expect(missing.stderr).toContain('--email needs a value.');
  });

  it('asks for flags when it cannot prompt', () => {
    const noEmail = run(['setup', '--no-login']);
    expect(noEmail.status).toBe(1);
    expect(noEmail.stderr).toContain('Missing --email');

    const noClient = run(['setup', '--no-login', '--email', 'teacher@example.edu']);
    expect(noClient.status).toBe(1);
    expect(noClient.stderr).toContain('Missing --client-id');
    expect(existsSync(configFile())).toBe(false);
  });

  it('saves settings, finding the tenant from the email domain', () => {
    const { status, stdout } = run([
      'setup',
      '--no-login',
      '--email',
      'teacher@example.edu',
      '--client-id',
      CLIENT_ID,
    ]);
    expect(status).toBe(0);
    expect(fetches()).toEqual([
      'https://login.microsoftonline.com/example.edu/v2.0/.well-known/openid-configuration',
    ]);
    expect(stdout).toContain(`Found your Microsoft 365 organization (tenant ${TENANT_ID}).`);
    expect(savedConfig()).toEqual({
      email: 'teacher@example.edu',
      tenantId: TENANT_ID,
      clientId: CLIENT_ID,
      teams: true,
    });
  });

  it('keeps settings private', () => {
    run(['setup', '--no-login', '--email', 'teacher@example.edu', '--client-id', CLIENT_ID]);
    expect(statSync(configFile()).mode & 0o777).toBe(0o600);
    expect(statSync(join(home, '.outlook-mcp')).mode & 0o777).toBe(0o700);
  });

  it('prints how to add the server to each app', () => {
    const { stdout } = run([
      'setup',
      '--no-login',
      '--email',
      'teacher@example.edu',
      '--client-id',
      CLIENT_ID,
    ]);
    expect(stdout).toContain('claude mcp add outlook --scope user -- npx -y @gchallen/outlook-mcp');
    expect(stdout).toContain('codex mcp add outlook -- npx -y @gchallen/outlook-mcp');
    expect(stdout).toContain('"args": [');
    expect(stdout).toContain('Your AI app will show a sign-in link');
  });

  it('skips the lookup when given a tenant ID, and records --no-teams', () => {
    const { status } = run(
      [
        'setup',
        '--no-login',
        '--no-teams',
        '--email',
        'teacher@example.edu',
        '--client-id',
        CLIENT_ID,
        '--tenant-id',
        TENANT_ID,
      ],
      { FAKE_FETCH: 'forbid' },
    );
    expect(status).toBe(0);
    expect(fetches()).toEqual([]);
    expect(savedConfig()).toMatchObject({ tenantId: TENANT_ID, teams: false });
  });

  it('reuses saved settings when run again, keeping the signature', () => {
    run(['setup', '--no-login', '--email', 'teacher@example.edu', '--client-id', CLIENT_ID]);
    const withSignature = { ...savedConfig(), signature: 'Pat Teacher' };
    writeFileSync(configFile(), JSON.stringify(withSignature));

    const again = run(['setup', '--no-login'], { FAKE_FETCH: 'tenant' });
    expect(again.status).toBe(0);
    expect(savedConfig()).toEqual(withSignature);
  });

  it('rejects a client ID that is not a GUID', () => {
    const { status, stderr } = run([
      'setup',
      '--no-login',
      '--email',
      'teacher@example.edu',
      '--client-id',
      'Outlook for AI',
    ]);
    expect(status).toBe(1);
    expect(stderr).toContain('"Outlook for AI" is not a valid client ID.');
    expect(existsSync(configFile())).toBe(false);
  });

  it('explains addresses that cannot work', () => {
    const notEmail = run(['setup', '--no-login', '--email', 'teacher', '--client-id', CLIENT_ID]);
    expect(notEmail.stderr).toContain('"teacher" does not look like an email address.');

    const personal = run(
      ['setup', '--no-login', '--email', 'someone@outlook.com', '--client-id', CLIENT_ID],
      { FAKE_FETCH: 'consumer' },
    );
    expect(personal.status).toBe(1);
    expect(personal.stderr).toContain('needs a work or school account');

    const unknown = run(
      ['setup', '--no-login', '--email', 'someone@example.org', '--client-id', CLIENT_ID],
      { FAKE_FETCH: 'missing' },
    );
    expect(unknown.status).toBe(1);
    expect(unknown.stderr).toContain('example.org does not appear to use Microsoft 365');
  });
});

/** A running server, spoken to over stdio the way an MCP client does. */
class StdioSession {
  private child: ChildProcess;
  private nextId = 1;
  private pending = new Map<number, (message: any) => void>();
  readonly notifications: any[] = [];
  readonly nonJson: string[] = [];

  constructor(extra: Record<string, string> = {}, args: string[] = []) {
    this.child = spawn('node', ['--require', FAKE_FETCH, CLI, ...args], {
      cwd,
      env: env(extra),
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    createInterface({ input: this.child.stdout! }).on('line', (line) => {
      let message: any;
      try {
        message = JSON.parse(line);
      } catch {
        this.nonJson.push(line);
        return;
      }
      if (message.id !== undefined && this.pending.has(message.id)) {
        this.pending.get(message.id)!(message);
        this.pending.delete(message.id);
      } else this.notifications.push(message);
    });
  }

  request(method: string, params: unknown = {}): Promise<any> {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error(`No reply to ${method}`)), 15_000);
      this.pending.set(id, (message) => {
        clearTimeout(timer);
        resolve(message);
      });
      this.child.stdin!.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
    });
  }

  notify(method: string) {
    this.child.stdin!.write(JSON.stringify({ jsonrpc: '2.0', method }) + '\n');
  }

  async start() {
    const init = await this.request('initialize', {
      protocolVersion: '2025-06-18',
      capabilities: {},
      clientInfo: { name: 'e2e', version: '1.0.0' },
    });
    this.notify('notifications/initialized');
    return init.result;
  }

  async callTool(name: string, args: Record<string, unknown> = {}): Promise<string> {
    const reply = await this.request('tools/call', { name, arguments: args });
    return reply.result.content[0].text;
  }

  stop(): Promise<void> {
    return new Promise((resolve) => {
      this.child.once('exit', () => resolve());
      this.child.stdin!.end();
      setTimeout(() => this.child.kill(), 2_000);
    });
  }
}

describe('MCP server over stdio', () => {
  let session: StdioSession | undefined;

  afterEach(async () => {
    await session?.stop();
    session = undefined;
  });

  it('starts with no setup and explains what to do', async () => {
    session = new StdioSession({ FAKE_FETCH: 'forbid' });
    const info = await session.start();
    expect(info.serverInfo).toEqual({
      name: 'outlook-mcp',
      version: JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf-8')).version,
    });

    const { result } = await session.request('tools/list');
    expect(result.tools.length).toBeGreaterThan(50);

    const status = await session.callTool('outlook_status');
    expect(status).toContain('Outlook is not set up yet');
    expect(status).toContain('npx -y @gchallen/outlook-mcp setup');
  });

  it('writes nothing but JSON-RPC to stdout', async () => {
    // A stray console.log or a dotenv banner on stdout corrupts the protocol
    // stream and disconnects the client.
    writeFileSync(join(cwd, '.env'), 'UNRELATED=1\n');
    session = new StdioSession({ FAKE_FETCH: 'forbid' });
    await session.start();
    await session.request('tools/list');
    await session.callTool('get_signature');
    await session.callTool('list_messages');
    expect(session.nonJson).toEqual([]);
  });

  it('sets the signature before setup and notifies the client', async () => {
    session = new StdioSession({ FAKE_FETCH: 'forbid' });
    await session.start();
    const saved = await session.callTool('set_signature', { signature: 'Pat Teacher\nProfessor' });
    expect(saved).toContain('Signature saved');
    expect(savedConfig()).toEqual({ signature: 'Pat Teacher\nProfessor' });
    expect(statSync(configFile()).mode & 0o777).toBe(0o600);
    expect(session.notifications.map((message) => message.method)).toContain(
      'notifications/tools/list_changed',
    );
  });

  it('uses the settings that setup saved', async () => {
    run(['setup', '--no-login', '--email', 'teacher@example.edu', '--client-id', CLIENT_ID]);
    writeFileSync(configFile(), JSON.stringify({ ...savedConfig(), signature: 'Pat Teacher' }));

    session = new StdioSession({ FAKE_FETCH: 'forbid' });
    await session.start();
    expect(await session.callTool('get_signature')).toContain('Pat Teacher');
  });

  it('reads Claude Desktop extension settings from the environment', async () => {
    session = new StdioSession({
      FAKE_FETCH: 'tenant',
      OUTLOOK_EMAIL: 'teacher@example.edu',
      AZURE_CLIENT_ID: CLIENT_ID,
      OUTLOOK_SIGNATURE: 'Pat Teacher\\nProfessor',
      OUTLOOK_TEAMS: 'true',
    });
    await session.start();
    expect(await session.callTool('get_signature')).toContain('Pat Teacher\nProfessor');
    expect(fetches()).toEqual([
      'https://login.microsoftonline.com/example.edu/v2.0/.well-known/openid-configuration',
    ]);
  });

  it('exits when the client closes its input, even with work pending', async () => {
    session = new StdioSession({ FAKE_FETCH: 'forbid', FAKE_KEEPALIVE: '1' });
    await session.start();
    const started = Date.now();
    await session.stop();
    session = undefined;
    // stop() kills the server after 2 seconds if it has not exited by itself.
    expect(Date.now() - started).toBeLessThan(1_500);
  });

  it('also runs as `outlook-mcp serve`', async () => {
    session = new StdioSession({ FAKE_FETCH: 'forbid' }, ['serve']);
    await session.start();
    expect(await session.callTool('get_signature')).toContain('No signature is configured');
  });
});
