import {
  configPath,
  discoverTenantId,
  isGuid,
  readConfigFile,
  resolveConfig,
  writeConfigFile,
} from '../config';
import { DEFAULT_SCOPES, TEAMS_SCOPES } from '../auth';
import { mkdtempSync, rmSync, statSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

describe('config', () => {
  let home: string;
  let env: NodeJS.ProcessEnv;

  beforeEach(() => {
    home = mkdtempSync(join(tmpdir(), 'outlook-mcp-config-'));
    env = { OUTLOOK_MCP_HOME: home };
  });

  afterEach(() => rmSync(home, { recursive: true, force: true }));

  describe('resolveConfig', () => {
    it('defaults to the config folder with Teams scopes at login', () => {
      const config = resolveConfig(env);
      expect(config.clientId).toBeUndefined();
      expect(config.configPath).toBe(join(home, 'config.json'));
      expect(config.cachePath).toBe(join(home, 'token-cache.json'));
      expect(config.scopes).toEqual(DEFAULT_SCOPES);
      expect(config.loginScopes).toEqual([...DEFAULT_SCOPES, ...TEAMS_SCOPES]);
    });

    it('reads the config file', () => {
      writeConfigFile(configPath(env), {
        clientId: 'cid',
        tenantId: 'tid',
        email: 'a@b.edu',
        signature: 'Sig',
        teams: false,
        extraScopes: ['ChannelMessage.Read.All'],
      });
      const config = resolveConfig(env);
      expect(config).toMatchObject({ clientId: 'cid', tenantId: 'tid', email: 'a@b.edu' });
      expect(config.signature).toBe('Sig');
      expect(config.loginScopes).toEqual([...DEFAULT_SCOPES, 'ChannelMessage.Read.All']);
    });

    it('lets env override connection settings but ignores empty values', () => {
      writeConfigFile(configPath(env), { clientId: 'file-cid', tenantId: 'file-tid' });
      const config = resolveConfig({
        ...env,
        AZURE_CLIENT_ID: 'env-cid',
        AZURE_TENANT_ID: '',
        OUTLOOK_TOKEN_CACHE: '/tmp/cache.json',
        OUTLOOK_TEAMS: 'false',
      });
      expect(config.clientId).toBe('env-cid');
      expect(config.tenantId).toBe('file-tid');
      expect(config.cachePath).toBe('/tmp/cache.json');
      expect(config.loginScopes).toEqual(DEFAULT_SCOPES);
    });

    it('ignores unfilled placeholders from the Claude Desktop extension', () => {
      writeConfigFile(configPath(env), { email: 'file@example.edu' });
      const config = resolveConfig({
        ...env,
        OUTLOOK_EMAIL: '${user_config.email}',
        OUTLOOK_SIGNATURE: '${user_config.signature}',
        OUTLOOK_TEAMS: '${user_config.teams}',
      });
      expect(config.email).toBe('file@example.edu');
      expect(config.signature).toBeUndefined();
      expect(config.loginScopes).toEqual([...DEFAULT_SCOPES, ...TEAMS_SCOPES]);
    });

    it('reads the extension settings as Claude Desktop passes them', () => {
      const config = resolveConfig({
        ...env,
        OUTLOOK_EMAIL: 'teacher@example.edu',
        AZURE_CLIENT_ID: '11111111-2222-4333-8444-555555555555',
        OUTLOOK_SIGNATURE: '',
        OUTLOOK_TEAMS: 'false',
      });
      expect(config).toMatchObject({
        email: 'teacher@example.edu',
        clientId: '11111111-2222-4333-8444-555555555555',
        tenantId: undefined,
        signature: undefined,
      });
      expect(config.loginScopes).toEqual(DEFAULT_SCOPES);
      expect(resolveConfig({ ...env, OUTLOOK_TEAMS: 'true' }).loginScopes).toEqual([
        ...DEFAULT_SCOPES,
        ...TEAMS_SCOPES,
      ]);
    });

    it('prefers the file signature over env, which may spell newlines as \\n', () => {
      expect(resolveConfig({ ...env, OUTLOOK_SIGNATURE: 'A\\nB' }).signature).toBe('A\nB');
      writeConfigFile(configPath(env), { signature: '' });
      expect(resolveConfig({ ...env, OUTLOOK_SIGNATURE: 'A' }).signature).toBe('');
    });

    it('uses GRAPH_SCOPES for both refresh and login when set', () => {
      const config = resolveConfig({ ...env, GRAPH_SCOPES: 'Mail.Read, User.Read' });
      expect(config.scopes).toEqual(['Mail.Read', 'User.Read']);
      expect(config.loginScopes).toEqual(['Mail.Read', 'User.Read']);
    });
  });

  describe('config file', () => {
    it('merges updates and keeps the file private', () => {
      const file = configPath(env);
      writeConfigFile(file, { clientId: 'cid' });
      writeConfigFile(file, { signature: 'Sig' });
      expect(readConfigFile(file)).toEqual({ clientId: 'cid', signature: 'Sig' });
      expect(statSync(file).mode & 0o777).toBe(0o600);
    });

    it('reports a malformed file instead of ignoring it', () => {
      const file = configPath(env);
      writeConfigFile(file, {});
      writeFileSync(file, '{ not json');
      expect(() => readConfigFile(file)).toThrow(/Could not read/);
    });
  });

  describe('discoverTenantId', () => {
    const realFetch = global.fetch;
    afterEach(() => {
      global.fetch = realFetch;
    });

    const mockFetch = (ok: boolean, body: unknown, status = 200) => {
      global.fetch = jest.fn().mockResolvedValue({
        ok,
        status,
        json: async () => body,
      }) as unknown as typeof fetch;
    };

    it('reads the tenant id from the domain metadata', async () => {
      mockFetch(true, {
        issuer: 'https://login.microsoftonline.com/11111111-2222-4333-8444-555555555555/v2.0',
      });
      await expect(discoverTenantId('someone@Example.edu')).resolves.toBe(
        '11111111-2222-4333-8444-555555555555',
      );
      expect(global.fetch).toHaveBeenCalledWith(
        'https://login.microsoftonline.com/example.edu/v2.0/.well-known/openid-configuration',
      );
    });

    it('rejects domains that are not on Microsoft 365', async () => {
      mockFetch(false, {}, 400);
      await expect(discoverTenantId('a@gmail.com')).rejects.toThrow(/does not appear to use/);
    });

    it('rejects personal Microsoft accounts', async () => {
      mockFetch(true, {
        issuer: 'https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0',
      });
      await expect(discoverTenantId('a@outlook.com')).rejects.toThrow(/personal/);
    });

    it('rejects things that are not email addresses', async () => {
      await expect(discoverTenantId('nope')).rejects.toThrow(/email address/);
    });
  });

  it('recognizes client ids', () => {
    expect(isGuid('aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee')).toBe(true);
    expect(isGuid('not-a-guid')).toBe(false);
  });
});
