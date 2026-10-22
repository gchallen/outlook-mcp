import { GraphAuthProvider, DEFAULT_SCOPES, AuthRequiredError, explainAuthError } from '../auth';
import { InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

jest.mock('@azure/msal-node');
jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: { initWithMiddleware: jest.fn(() => ({ isGraphClient: true })) },
}));

const MockPCA: any = PublicClientApplication;

describe('GraphAuthProvider', () => {
  let getAllAccounts: jest.Mock;
  let acquireTokenSilent: jest.Mock;
  let acquireTokenByDeviceCode: jest.Mock;

  beforeEach(() => {
    jest.clearAllMocks();
    getAllAccounts = jest.fn();
    acquireTokenSilent = jest.fn();
    acquireTokenByDeviceCode = jest.fn();
    MockPCA.mockImplementation(() => ({
      getTokenCache: () => ({ getAllAccounts }),
      acquireTokenSilent,
      acquireTokenByDeviceCode,
    }));
  });

  describe('token cache file', () => {
    let dir: string;
    let cachePath: string;
    let plugin: any;

    beforeEach(() => {
      dir = mkdtempSync(join(tmpdir(), 'outlook-mcp-cache-'));
      cachePath = join(dir, 'nested', 'token-cache.json');
      new GraphAuthProvider({ clientId: 'cid', tenantId: 'tid', cachePath });
      plugin = MockPCA.mock.calls[0][0].cache.cachePlugin;
    });

    afterEach(() => rmSync(dir, { recursive: true, force: true }));

    const context = (overrides: Record<string, unknown> = {}) => ({
      cacheHasChanged: false,
      tokenCache: {
        deserialize: jest.fn(),
        serialize: jest.fn().mockReturnValue('{"cached":true}'),
      },
      ...overrides,
    });

    it('creates the cache folder privately', () => {
      expect(statSync(join(dir, 'nested')).mode & 0o777).toBe(0o700);
    });

    it('does nothing on read when there is no cache yet', async () => {
      const ctx = context();
      await plugin.beforeCacheAccess(ctx);
      expect(ctx.tokenCache.deserialize).not.toHaveBeenCalled();
    });

    it('writes a changed cache readable only by the user, and reads it back', async () => {
      await plugin.afterCacheAccess(context({ cacheHasChanged: true }));
      expect(readFileSync(cachePath, 'utf-8')).toBe('{"cached":true}');
      expect(statSync(cachePath).mode & 0o777).toBe(0o600);

      const ctx = context();
      await plugin.beforeCacheAccess(ctx);
      expect(ctx.tokenCache.deserialize).toHaveBeenCalledWith('{"cached":true}');
    });

    it('does not rewrite an unchanged cache', async () => {
      const ctx = context();
      await plugin.afterCacheAccess(ctx);
      expect(ctx.tokenCache.serialize).not.toHaveBeenCalled();
      expect(existsSync(cachePath)).toBe(false);
    });

    it('logs instead of crashing when the cache cannot be read or written', async () => {
      const log = jest.spyOn(console, 'error').mockImplementation(() => undefined);
      writeFileSync(cachePath, 'not json');
      const bad = context();
      bad.tokenCache.deserialize.mockImplementation(() => {
        throw new Error('bad cache');
      });
      await expect(plugin.beforeCacheAccess(bad)).resolves.toBeUndefined();

      rmSync(cachePath);
      mkdirSync(cachePath); // a folder where the file should be
      await expect(
        plugin.afterCacheAccess(context({ cacheHasChanged: true })),
      ).resolves.toBeUndefined();
      expect(log).toHaveBeenCalledWith('Failed to read token cache:', expect.any(Error));
      expect(log).toHaveBeenCalledWith(
        'Failed to write token cache:',
        expect.objectContaining({ code: 'EISDIR' }),
      );
      log.mockRestore();
    });
  });

  describe('construction', () => {
    it('builds the authority from the tenant id', () => {
      new GraphAuthProvider({ clientId: 'cid', tenantId: 'tid' });
      expect(MockPCA).toHaveBeenCalledWith(
        expect.objectContaining({
          auth: expect.objectContaining({
            clientId: 'cid',
            authority: 'https://login.microsoftonline.com/tid',
          }),
        }),
      );
    });

    it('falls back to the common authority when no tenant is given', () => {
      new GraphAuthProvider({ clientId: 'cid' });
      expect(MockPCA).toHaveBeenCalledWith(
        expect.objectContaining({
          auth: expect.objectContaining({
            authority: 'https://login.microsoftonline.com/common',
          }),
        }),
      );
    });

    it('wires the provider into a Graph client', () => {
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      expect(Client.initWithMiddleware).toHaveBeenCalledWith(
        expect.objectContaining({ authProvider: auth }),
      );
      expect(auth.getClient()).toEqual({ isGraphClient: true });
    });
  });

  describe('getAccessToken', () => {
    it('throws a helpful error when there are no cached accounts', async () => {
      getAllAccounts.mockResolvedValue([]);
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      await expect(auth.getAccessToken()).rejects.toThrow(AuthRequiredError);
    });

    it('treats an expired refresh token as needing sign-in', async () => {
      getAllAccounts.mockResolvedValue([{ username: 'u' }]);
      acquireTokenSilent.mockRejectedValue(
        new InteractionRequiredAuthError('interaction_required'),
      );
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      await expect(auth.getAccessToken()).rejects.toThrow(/expired/);
      await expect(auth.getAccessToken()).rejects.toBeInstanceOf(AuthRequiredError);
    });

    it('passes other silent-acquisition errors through', async () => {
      getAllAccounts.mockResolvedValue([{ username: 'u' }]);
      acquireTokenSilent.mockRejectedValue(new Error('network down'));
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      await expect(auth.getAccessToken()).rejects.toThrow('network down');
    });

    it('returns the token from a silent acquisition using the configured scopes', async () => {
      getAllAccounts.mockResolvedValue([{ username: 'u' }]);
      acquireTokenSilent.mockResolvedValue({ accessToken: 'tok' });
      const auth = new GraphAuthProvider({ clientId: 'cid', scopes: ['Scope.A'] });

      await expect(auth.getAccessToken()).resolves.toBe('tok');
      expect(acquireTokenSilent).toHaveBeenCalledWith({
        account: { username: 'u' },
        scopes: ['Scope.A'],
      });
    });

    it('throws when silent acquisition yields no token', async () => {
      getAllAccounts.mockResolvedValue([{ username: 'u' }]);
      acquireTokenSilent.mockResolvedValue(null);
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      await expect(auth.getAccessToken()).rejects.toBeInstanceOf(AuthRequiredError);
    });
  });

  describe('getUsername', () => {
    it('returns the cached account, or undefined before sign-in', async () => {
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      getAllAccounts.mockResolvedValueOnce([{ username: 'teacher@example.edu' }]);
      await expect(auth.getUsername()).resolves.toBe('teacher@example.edu');
      getAllAccounts.mockResolvedValueOnce([]);
      await expect(auth.getUsername()).resolves.toBeUndefined();
    });
  });

  describe('login', () => {
    it('runs the device-code flow with the default scopes', async () => {
      acquireTokenByDeviceCode.mockResolvedValue({ account: { username: 'u@illinois.edu' } });
      const errSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
      const auth = new GraphAuthProvider({ clientId: 'cid' });

      await auth.login();

      expect(acquireTokenByDeviceCode).toHaveBeenCalledWith(
        expect.objectContaining({ scopes: DEFAULT_SCOPES }),
      );
      errSpy.mockRestore();
    });

    it('throws when device-code login returns no account', async () => {
      acquireTokenByDeviceCode.mockResolvedValue({});
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      await expect(auth.login()).rejects.toThrow(/did not return an account/);
    });

    it('passes a deviceCodeCallback that writes to stderr', async () => {
      acquireTokenByDeviceCode.mockImplementation(async (opts: any) => {
        opts.deviceCodeCallback({ message: 'go enter code XYZ' });
        return { account: { username: 'u' } };
      });
      const errSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
      const auth = new GraphAuthProvider({ clientId: 'cid' });

      await auth.login();

      expect(errSpy).toHaveBeenCalledWith(expect.stringContaining('go enter code XYZ'));
      errSpy.mockRestore();
    });
  });

  describe('login scopes', () => {
    it('requests loginScopes at sign-in and scopes on refresh', async () => {
      acquireTokenByDeviceCode.mockResolvedValue({ account: { username: 'u' } });
      const errSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
      const auth = new GraphAuthProvider({
        clientId: 'cid',
        scopes: ['Scope.A'],
        loginScopes: ['Scope.A', 'Scope.B'],
      });

      await auth.login();

      expect(acquireTokenByDeviceCode).toHaveBeenCalledWith(
        expect.objectContaining({ scopes: ['Scope.A', 'Scope.B'] }),
      );
      errSpy.mockRestore();
    });
  });

  describe('beginDeviceLogin', () => {
    it('resolves with the device-code instructions while sign-in continues', async () => {
      let finish!: (value: unknown) => void;
      acquireTokenByDeviceCode.mockImplementation((opts: any) => {
        opts.deviceCodeCallback({ message: 'Go to https://microsoft.com/devicelogin, enter ABC' });
        return new Promise((resolve) => (finish = resolve));
      });
      const auth = new GraphAuthProvider({ clientId: 'cid' });

      await expect(auth.beginDeviceLogin()).resolves.toMatch(/enter ABC/);
      // A second call while pending reuses the same sign-in.
      await expect(auth.beginDeviceLogin()).resolves.toMatch(/enter ABC/);
      expect(acquireTokenByDeviceCode).toHaveBeenCalledTimes(1);

      finish({ account: { username: 'u' } });
      await new Promise((resolve) => setImmediate(resolve));
      auth.beginDeviceLogin();
      expect(acquireTokenByDeviceCode).toHaveBeenCalledTimes(2);
    });

    it('rejects when sign-in fails before a code is issued', async () => {
      acquireTokenByDeviceCode.mockRejectedValue(new Error('AADSTS700016: app not found'));
      const auth = new GraphAuthProvider({ clientId: 'cid' });
      await expect(auth.beginDeviceLogin()).rejects.toThrow(/AADSTS700016/);
      expect(auth.takeLoginError()).toBeUndefined();
    });

    it('records a failure after the code was shown for the next call', async () => {
      let fail!: (error: Error) => void;
      acquireTokenByDeviceCode.mockImplementation((opts: any) => {
        opts.deviceCodeCallback({ message: 'enter ABC' });
        return new Promise((_resolve, reject) => (fail = reject));
      });
      const auth = new GraphAuthProvider({ clientId: 'cid' });

      await auth.beginDeviceLogin();
      fail(new Error('AADSTS65001: consent required'));
      await new Promise((resolve) => setImmediate(resolve));

      expect(String(auth.takeLoginError())).toMatch(/AADSTS65001/);
      expect(auth.takeLoginError()).toBeUndefined();
    });
  });

  describe('explainAuthError', () => {
    it('explains admin-consent failures in plain language', () => {
      const text = explainAuthError(
        new Error('AADSTS65001: The user or administrator has not consented'),
      );
      expect(text).toMatch(/administrator to approve/);
      expect(text).toMatch(/--no-teams/);
      expect(text).toContain('AADSTS65001');
    });

    it('explains a disabled public client flow', () => {
      expect(explainAuthError(new Error('AADSTS7000218: client_assertion'))).toMatch(
        /Allow public client flows/,
      );
    });

    it('tells Claude Desktop users where the Teams switch is', () => {
      expect(explainAuthError(new Error('AADSTS90094'))).toContain('Include Microsoft Teams');
    });

    it.each([
      ['AADSTS65001: consent', /administrator to approve/],
      ['AADSTS90094: admin consent', /administrator to approve/],
      ['AADSTS90095: admin consent', /administrator to approve/],
      ['AADSTS900941: admin consent', /administrator to approve/],
      ['AADSTS700016: application not found', /could not find this app/],
      ['AADSTS7000218: client_assertion', /Allow public client flows/],
      ['AADSTS50020: user account from another tenant', /does not belong to the organization/],
      ['AADSTS50194: not multi-tenant', /does not belong to the organization/],
      ['AADSTS70019: verification code expired', /code expired/],
      ['expired_token', /code expired/],
      ['AADSTS70000: declined', /cancelled or the permissions were declined/],
      ['authorization_declined', /cancelled or the permissions were declined/],
    ])('explains %s', (raw, hint) => {
      const text = explainAuthError(new Error(raw));
      expect(text).toMatch(hint);
      expect(text).toContain(`(Details: ${raw})`);
    });

    it('explains non-Error values', () => {
      expect(explainAuthError('AADSTS7000218')).toMatch(/Allow public client flows/);
    });

    it('returns unknown errors unchanged', () => {
      expect(explainAuthError(new Error('something odd'))).toBe('something odd');
    });
  });
});
