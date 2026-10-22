import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { ToolListChangedNotificationSchema, type Tool } from '@modelcontextprotocol/sdk/types.js';
import { mkdtempSync, readFileSync, rmSync, statSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { AuthRequiredError, GraphAuthProvider } from '../auth';
import { resolveConfig, writeConfigFile, type ResolvedConfig } from '../config';
import { SERVER_VERSION, buildTools, initialize, server } from '../server';
import { MessageTools } from '../tools/message-tools';
import { FolderTools } from '../tools/folder-tools';
import { RuleTools } from '../tools/rule-tools';
import { CalendarTools } from '../tools/calendar-tools';
import { ContactsTools } from '../tools/contacts-tools';
import { PeopleTools } from '../tools/people-tools';
import { TeamsTools } from '../tools/teams-tools';

// Every Graph tool class is auto-mocked: each method is a jest.fn, so these
// tests check the MCP layer (listing, dispatch, argument mapping, sign-in and
// setup handling) without touching Microsoft Graph.
jest.mock('../tools/message-tools');
jest.mock('../tools/folder-tools');
jest.mock('../tools/rule-tools');
jest.mock('../tools/calendar-tools');
jest.mock('../tools/contacts-tools');
jest.mock('../tools/people-tools');
jest.mock('../tools/teams-tools');

const TOOL_CLASSES = [
  MessageTools,
  FolderTools,
  RuleTools,
  CalendarTools,
  ContactsTools,
  PeopleTools,
  TeamsTools,
] as unknown as jest.Mock[];

// Tools the server answers itself instead of forwarding to a Graph tool class.
const BUILT_IN = ['outlook_status', 'get_signature', 'set_signature'];

const CLIENT_ID = '11111111-2222-4333-8444-555555555555';
const TENANT_ID = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';

function fakeAuth(overrides: Partial<Record<keyof GraphAuthProvider, jest.Mock>> = {}) {
  return {
    getAccessToken: jest.fn().mockResolvedValue('token'),
    getUsername: jest.fn().mockResolvedValue('teacher@example.edu'),
    beginDeviceLogin: jest
      .fn()
      .mockResolvedValue(
        'To sign in, go to https://microsoft.com/devicelogin and enter ABCD-1234.',
      ),
    takeLoginError: jest.fn().mockReturnValue(undefined),
    getClient: jest.fn().mockReturnValue({}),
    ...overrides,
  } as unknown as GraphAuthProvider & Record<string, jest.Mock>;
}

/** Every call made to any method of any mocked Graph tool instance. */
function graphCalls(): { method: string; args: unknown[] }[] {
  const calls: { method: string; args: unknown[] }[] = [];
  for (const cls of TOOL_CLASSES) {
    for (const [method, fn] of Object.entries(cls.prototype)) {
      if (method === 'constructor') continue;
      for (const args of (fn as jest.Mock).mock?.calls ?? []) calls.push({ method, args });
    }
  }
  return calls;
}

/** A plausible value for each property of a tool's input schema. */
function sampleArgs(tool: Tool): Record<string, unknown> {
  const args: Record<string, unknown> = {};
  for (const [key, schema] of Object.entries(tool.inputSchema.properties ?? {})) {
    const { type, enum: choices } = schema as { type?: string; enum?: unknown[] };
    if (choices?.length) args[key] = choices[0];
    else if (type === 'number' || type === 'integer') args[key] = 3;
    else if (type === 'boolean') args[key] = true;
    else if (type === 'array') args[key] = [`${key}-item`];
    else if (type === 'object') args[key] = { marker: `${key}-object` };
    else args[key] = `${key}-value`;
  }
  return args;
}

const text = (result: unknown) =>
  ((result as { content: { type: string; text: string }[] }).content[0]?.text ?? '') as string;

describe('MCP server', () => {
  let client: Client;
  let home: string;
  let listChanged: number;

  const configFor = (overrides: Partial<ResolvedConfig> = {}): ResolvedConfig => ({
    ...resolveConfig({ OUTLOOK_MCP_HOME: home }),
    ...overrides,
  });

  const call = async (name: string, args: Record<string, unknown> = {}) =>
    text(await client.callTool({ name, arguments: args }));

  beforeAll(async () => {
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    client = new Client({ name: 'test', version: '1.0.0' });
    listChanged = 0;
    client.setNotificationHandler(ToolListChangedNotificationSchema, async () => {
      listChanged++;
    });
    await server.connect(serverTransport);
    await client.connect(clientTransport);
  });

  afterAll(async () => {
    await client.close();
  });

  beforeEach(() => {
    home = mkdtempSync(join(tmpdir(), 'outlook-mcp-server-'));
    jest.clearAllMocks();
  });

  afterEach(() => rmSync(home, { recursive: true, force: true }));

  describe('tool list', () => {
    let tools: Tool[];

    beforeEach(async () => {
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), {
        auth: fakeAuth(),
      });
      tools = (await client.listTools()).tools;
    });

    it('reports the package version', () => {
      const pkg = JSON.parse(readFileSync(join(__dirname, '../../package.json'), 'utf-8'));
      expect(SERVER_VERSION).toBe(pkg.version);
      expect(client.getServerVersion()).toMatchObject({
        name: 'outlook-mcp',
        version: pkg.version,
      });
    });

    it('advertises tool list changes', () => {
      expect(client.getServerCapabilities()?.tools).toEqual({ listChanged: true });
    });

    it('has unique, well-formed tool names', () => {
      const names = tools.map((tool) => tool.name);
      expect(new Set(names).size).toBe(names.length);
      for (const name of names) expect(name).toMatch(/^[a-z][a-z_]*[a-z]$/);
      expect(names).toEqual(expect.arrayContaining(BUILT_IN));
    });

    it('gives every tool a description and a valid object schema', () => {
      for (const tool of tools) {
        expect(tool.description?.trim().length).toBeGreaterThan(10);
        expect(tool.inputSchema.type).toBe('object');
        const properties = Object.keys(tool.inputSchema.properties ?? {});
        for (const required of (tool.inputSchema.required as string[] | undefined) ?? []) {
          expect({ tool: tool.name, required, ok: properties.includes(required) }).toEqual({
            tool: tool.name,
            required,
            ok: true,
          });
        }
        for (const [key, schema] of Object.entries(tool.inputSchema.properties ?? {})) {
          const { type, items } = schema as { type?: string; items?: unknown };
          if (type === 'array')
            expect({ tool: tool.name, key, items: !!items }).toMatchObject({ items: true });
        }
      }
    });

    it('handles every listed tool, and lists every handled tool', () => {
      const source = readFileSync(join(__dirname, '../server.ts'), 'utf-8');
      const handled = [...source.matchAll(/case '([a-z_]+)':/g)].map((match) => match[1]);
      const special = [...source.matchAll(/name === '([a-z_]+)'/g)].map((match) => match[1]);
      expect([...handled, ...special].sort()).toEqual(tools.map((tool) => tool.name).sort());
    });
  });

  describe('Graph tool dispatch', () => {
    const tools = buildTools().filter((tool) => !BUILT_IN.includes(tool.name));

    beforeEach(async () => {
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), {
        auth: fakeAuth(),
      });
    });

    it.each(tools.map((tool) => [tool.name, tool] as const))(
      '%s calls exactly one Graph method with its arguments',
      async (_name, tool) => {
        const args = sampleArgs(tool);
        const result = await call(tool.name, args);

        expect(result).not.toMatch(/^Error:/);
        const calls = graphCalls();
        expect(calls).toHaveLength(1);

        // Every required argument reaches the Graph method. This catches a
        // handler that reads the wrong argument name (args.messageID).
        const forwarded = JSON.stringify(calls[0].args);
        for (const required of (tool.inputSchema.required as string[] | undefined) ?? []) {
          const value = JSON.stringify(args[required]);
          const inner = value.replace(/^\[|\]$/g, '');
          expect({ required, forwarded: forwarded.includes(inner) }).toEqual({
            required,
            forwarded: true,
          });
        }
      },
    );

    it('returns Graph results as JSON text', async () => {
      (MessageTools.prototype.listMessages as jest.Mock).mockResolvedValueOnce({
        value: [{ id: 'm1', subject: 'Hello' }],
      });
      expect(JSON.parse(await call('list_messages'))).toEqual({
        value: [{ id: 'm1', subject: 'Hello' }],
      });
    });

    it('returns text when Graph returns no content', async () => {
      (TeamsTools.prototype.listJoinedTeams as jest.Mock).mockResolvedValueOnce(undefined);
      expect(await call('list_joined_teams')).toBe('Done.');
    });

    it('maps send_mail arguments into a Graph message', async () => {
      await call('send_mail', {
        subject: 'Office Hours',
        body: { contentType: 'Text', content: 'See you there.' },
        toRecipients: [{ emailAddress: { address: 'student@example.edu' } }],
      });
      expect(MessageTools.prototype.sendMail).toHaveBeenCalledWith({
        message: {
          subject: 'Office Hours',
          body: { contentType: 'Text', content: 'See you there.' },
          toRecipients: [{ emailAddress: { address: 'student@example.edu' } }],
        },
        saveToSentItems: true,
      });
    });

    it('summarizes reply drafts instead of returning the whole message', async () => {
      (MessageTools.prototype.replyToMessageDraft as jest.Mock).mockResolvedValueOnce({
        id: 'd1',
        subject: 'RE: Question',
        isDraft: true,
        webLink: 'https://outlook.example/d1',
        body: { content: '<html>very long quoted thread</html>' },
      });
      const draft = JSON.parse(
        await call('reply_to_message_draft', { messageId: 'm1', comment: 'Yes.' }),
      );
      expect(draft).toMatchObject({
        id: 'd1',
        isDraft: true,
        webLink: 'https://outlook.example/d1',
      });
      expect(draft.body).toBeUndefined();
    });

    it('reports Graph errors as tool text instead of failing the request', async () => {
      (MessageTools.prototype.getMessage as jest.Mock).mockRejectedValueOnce(
        new Error('Not found'),
      );
      expect(await call('get_message', { messageId: 'missing' })).toBe('Error: Not found');
    });

    it('explains a 403 as a missing permission', async () => {
      const forbidden = Object.assign(new Error('Forbidden'), { statusCode: 403 });
      (TeamsTools.prototype.listChats as jest.Mock).mockRejectedValueOnce(forbidden);
      const result = await call('list_teams_chats');
      expect(result).toMatch(/^Error: Forbidden/);
      expect(result).toContain('the permission the Teams feature needs was not granted');

      (MessageTools.prototype.listMessages as jest.Mock).mockRejectedValueOnce(forbidden);
      expect(await call('list_messages')).toContain('the permission this feature needs');
    });

    it('does not add the permission hint to other errors', async () => {
      const notFound = Object.assign(new Error('Missing'), { statusCode: 404 });
      (MessageTools.prototype.listMessages as jest.Mock).mockRejectedValueOnce(notFound);
      expect(await call('list_messages')).toBe('Error: Missing');
    });

    it('rejects unknown tools', async () => {
      expect(await call('no_such_tool')).toBe('Error: Unknown tool: no_such_tool');
    });
  });

  describe('outlook_status', () => {
    it('reports the account and the signature', async () => {
      await initialize(
        configFor({
          clientId: CLIENT_ID,
          tenantId: TENANT_ID,
          signature: 'Pat Teacher\nProfessor',
        }),
        {
          auth: fakeAuth(),
        },
      );
      const status = await call('outlook_status');
      expect(status).toContain('Connected to Outlook as teacher@example.edu.');
      expect(status).toContain(join(home, 'config.json'));
      expect(status).toContain('Signature: configured (Pat Teacher / Professor)');
    });

    it('says when no signature is configured', async () => {
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), {
        auth: fakeAuth(),
      });
      expect(await call('outlook_status')).toContain('Signature: none configured');
    });
  });

  describe('when setup is incomplete', () => {
    it('explains how to set up when there is no client ID', async () => {
      await initialize(configFor());
      for (const name of ['outlook_status', 'list_messages', 'send_mail']) {
        const result = await call(name);
        expect(result).toContain('Outlook is not set up yet');
        expect(result).toContain('GUIDED-SETUP.md');
        expect(result).toContain('npx -y @gchallen/outlook-mcp setup');
      }
      expect(graphCalls()).toHaveLength(0);
    });

    it('explains when neither a tenant nor an email is configured', async () => {
      await initialize(configFor({ clientId: CLIENT_ID }));
      expect(await call('list_messages')).toContain('neither a tenant ID nor an email address');
    });

    describe('tenant discovery', () => {
      const realFetch = global.fetch;
      afterEach(() => {
        global.fetch = realFetch;
      });

      it('finds the tenant from the email domain', async () => {
        global.fetch = jest.fn().mockResolvedValue({
          ok: true,
          json: async () => ({ issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0` }),
        }) as unknown as typeof fetch;
        await initialize(configFor({ clientId: CLIENT_ID, email: 'teacher@example.edu' }), {
          auth: fakeAuth(),
        });
        expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining('/example.edu/v2.0/'));
        expect(await call('outlook_status')).toContain('Connected to Outlook');
      });

      it('reports a domain that does not use Microsoft 365', async () => {
        global.fetch = jest
          .fn()
          .mockResolvedValue({ ok: false, status: 400 }) as unknown as typeof fetch;
        await initialize(configFor({ clientId: CLIENT_ID, email: 'teacher@example.edu' }));
        expect(await call('list_messages')).toBe(
          'Could not look up the Microsoft 365 organization for teacher@example.edu: example.edu does not appear to use Microsoft 365 (Microsoft returned 400).',
        );
      });

      it('reports a network failure', async () => {
        global.fetch = jest
          .fn()
          .mockRejectedValue(new Error('getaddrinfo ENOTFOUND')) as unknown as typeof fetch;
        await initialize(configFor({ clientId: CLIENT_ID, email: 'teacher@example.edu' }));
        expect(await call('list_messages')).toContain('getaddrinfo ENOTFOUND');
      });
    });
  });

  describe('sign-in from a tool call', () => {
    it('starts a device-code sign-in when nobody is signed in', async () => {
      const auth = fakeAuth({
        getAccessToken: jest.fn().mockRejectedValue(new AuthRequiredError()),
      });
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), { auth });
      const result = await call('list_messages');
      expect(result).toContain('https://microsoft.com/devicelogin');
      expect(result).toContain('ABCD-1234');
      expect(result).toContain('call this tool again');
      expect(auth.beginDeviceLogin).toHaveBeenCalledTimes(1);
      expect(graphCalls()).toHaveLength(0);
    });

    it('says why sign-in is needed again when the sign-in expired', async () => {
      const auth = fakeAuth({
        getAccessToken: jest
          .fn()
          .mockRejectedValue(new AuthRequiredError('Your Microsoft 365 sign-in has expired.')),
      });
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), { auth });
      expect(await call('outlook_status')).toContain('Your Microsoft 365 sign-in has expired.');
    });

    it('explains why the previous sign-in attempt failed', async () => {
      const auth = fakeAuth({
        getAccessToken: jest.fn().mockRejectedValue(new AuthRequiredError()),
        takeLoginError: jest.fn().mockReturnValue(new Error('AADSTS65001: consent required')),
      });
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), { auth });
      const result = await call('list_messages');
      expect(result).toMatch(/^The previous sign-in attempt failed: .*administrator/);
      expect(result).toContain('https://microsoft.com/devicelogin');
    });

    it('explains when a sign-in cannot be started', async () => {
      const auth = fakeAuth({
        getAccessToken: jest.fn().mockRejectedValue(new AuthRequiredError()),
        beginDeviceLogin: jest.fn().mockRejectedValue(new Error('AADSTS700016: app not found')),
      });
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), { auth });
      const result = await call('list_messages');
      expect(result).toMatch(/^Could not start a Microsoft sign-in: /);
      expect(result).toContain('AADSTS700016');
    });

    it('passes through errors that are not about sign-in', async () => {
      const auth = fakeAuth({
        getAccessToken: jest.fn().mockRejectedValue(new Error('socket hang up')),
      });
      await initialize(configFor({ clientId: CLIENT_ID, tenantId: TENANT_ID }), { auth });
      expect(await call('list_messages')).toBe('Error: socket hang up');
      expect(auth.beginDeviceLogin).not.toHaveBeenCalled();
    });
  });

  describe('signature tools', () => {
    const descriptionOf = async (name: string) =>
      (await client.listTools()).tools.find((tool) => tool.name === name)!.description!;

    it('work before setup and without a sign-in', async () => {
      await initialize(configFor());
      expect(await call('get_signature')).toBe(
        'No signature is configured. Nothing is appended to outgoing mail.',
      );
      expect(await call('set_signature', { signature: 'Pat Teacher' })).toContain(
        'Signature saved',
      );
    });

    it('saves the signature privately and keeps other settings', async () => {
      const configPath = join(home, 'config.json');
      writeConfigFile(configPath, { clientId: CLIENT_ID, email: 'teacher@example.edu' });
      await initialize(configFor());

      await call('set_signature', { signature: 'Pat Teacher\nProfessor of Physics' });

      const saved = JSON.parse(readFileSync(configPath, 'utf-8'));
      expect(saved).toEqual({
        clientId: CLIENT_ID,
        email: 'teacher@example.edu',
        signature: 'Pat Teacher\nProfessor of Physics',
      });
      expect(statSync(configPath).mode & 0o777).toBe(0o600);
      expect(await call('get_signature')).toContain('Pat Teacher\nProfessor of Physics');
    });

    it('updates tool descriptions and notifies the client', async () => {
      await initialize(configFor());
      expect(await descriptionOf('create_message')).toContain('No signature is configured');

      const before = listChanged;
      await call('set_signature', { signature: 'Pat Teacher' });
      await new Promise((resolve) => setImmediate(resolve));

      expect(listChanged).toBe(before + 1);
      expect(await descriptionOf('create_message')).toContain('Do not write your own signature');

      await call('set_signature', { signature: '' });
      expect(await call('get_signature')).toContain('No signature is configured');
      expect(await descriptionOf('create_message')).toContain('No signature is configured');
    });

    it('accepts an HTML signature', async () => {
      await initialize(configFor());
      await call('set_signature', {
        signature: '<b>Pat Teacher</b> // <a href="https://example.edu">example.edu</a>',
      });
      const result = await call('get_signature');
      expect(result).toContain('Pat Teacher');
      expect(result).toContain('<div id="Signature">');
    });

    it('uses the configured signature from startup', async () => {
      await initialize(configFor({ signature: 'From Config' }));
      expect(await call('get_signature')).toContain('From Config');
      expect(await descriptionOf('send_mail')).toContain('styled HTML signature is appended');
    });
  });
});
