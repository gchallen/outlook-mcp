import { GraphAuthProvider, DEFAULT_SCOPES, AuthRequiredError, explainAuthError } from '../auth';
import { InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-node';
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

    it('returns unknown errors unchanged', () => {
      expect(explainAuthError(new Error('something odd'))).toBe('something odd');
    });
  });
});
