import { GraphAuthProvider, DEFAULT_SCOPES } from '../auth';
import { PublicClientApplication } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';

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
      await expect(auth.getAccessToken()).rejects.toThrow(/Not authenticated.*npm run login/s);
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
      await expect(auth.getAccessToken()).rejects.toThrow(/Failed to acquire/);
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
});
