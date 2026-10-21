import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type ICachePlugin,
  type TokenCacheContext,
} from '@azure/msal-node';
import { Client, type AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import * as fs from 'fs';
import * as path from 'path';

// Delegated Graph scopes the server needs. `offline_access` (refresh tokens) and
// `openid`/`profile` are added by MSAL automatically.
export const DEFAULT_SCOPES = [
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.ReadWrite',
  'Contacts.ReadWrite',
  'People.Read',
  'User.Read',
];

// Teams scopes, requested at login alongside DEFAULT_SCOPES but not on silent
// refresh: once consented they ride along in every Graph token, and if the
// tenant refuses them the mail tools keep working. All are user-consentable;
// reading channel posts additionally needs ChannelMessage.Read.All, which
// requires admin consent, so it is opt-in via EXTRA_GRAPH_SCOPES.
export const TEAMS_SCOPES = [
  'Chat.ReadWrite',
  'ChatMessage.Send',
  'Team.ReadBasic.All',
  'Channel.ReadBasic.All',
  'ChannelMessage.Send',
];

export interface AuthConfig {
  clientId: string;
  tenantId?: string;
  scopes?: string[];
  /** Scopes requested at interactive sign-in; defaults to `scopes`. */
  loginScopes?: string[];
  cachePath?: string;
}

/** Thrown when there is no usable cached sign-in and the user must sign in. */
export class AuthRequiredError extends Error {
  constructor(message = 'Not signed in to Microsoft 365.') {
    super(message);
    this.name = 'AuthRequiredError';
  }
}

// Entra error codes a self-service user is likely to hit, in plain language.
// Matching on the AADSTS code rather than the message text keeps this stable
// across Microsoft's wording changes and localizations.
const AUTH_ERROR_HINTS: [RegExp, string][] = [
  [
    /AADSTS(65001|90094|90095|900941)/,
    'Your organization requires an administrator to approve one or more of the requested permissions. The Teams permissions are the most likely cause: try again without them by re-running setup with --no-teams (or adding "teams": false to ~/.outlook-mcp/config.json and restarting the AI app). If that still fails, self-service setup is blocked in this organization and IT would need to approve the app.',
  ],
  [
    /AADSTS700016/,
    'Microsoft could not find this app in your organization. Check that the client ID is the "Application (client) ID" of an app registered with your own work or school account.',
  ],
  [
    /AADSTS7000218/,
    'The app registration does not allow device sign-in. In the Entra portal, open the app, go to Authentication, set "Allow public client flows" to Yes, and save.',
  ],
  [
    /AADSTS50020|AADSTS50194/,
    'The account you signed in with does not belong to the organization this app is registered in. Sign in with the same work or school account you used to register the app.',
  ],
  [
    /AADSTS(70019|70016)|code_expired|expired_token/,
    'The sign-in code expired before sign-in finished. Try again and complete the sign-in within 15 minutes.',
  ],
  [
    /AADSTS(70000|65004)|access_denied|authorization_declined/,
    'Sign-in was cancelled or the permissions were declined.',
  ],
];

/** A plain-language explanation of a sign-in failure, with the raw message kept for reference. */
export function explainAuthError(error: unknown): string {
  const raw = error instanceof Error ? error.message : String(error);
  const hint = AUTH_ERROR_HINTS.find(([pattern]) => pattern.test(raw))?.[1];
  return hint ? `${hint} (Details: ${raw})` : raw;
}

/**
 * Persist the MSAL token cache to a JSON file so the user only has to complete
 * the interactive device-code login once. Subsequent server starts refresh
 * silently from the cached refresh token.
 */
function createCachePlugin(cachePath: string): ICachePlugin {
  return {
    beforeCacheAccess: async (context: TokenCacheContext) => {
      try {
        if (fs.existsSync(cachePath)) {
          context.tokenCache.deserialize(fs.readFileSync(cachePath, 'utf-8'));
        }
      } catch (error) {
        console.error('Failed to read token cache:', error);
      }
    },
    afterCacheAccess: async (context: TokenCacheContext) => {
      if (context.cacheHasChanged) {
        try {
          fs.writeFileSync(cachePath, context.tokenCache.serialize(), { mode: 0o600 });
        } catch (error) {
          console.error('Failed to write token cache:', error);
        }
      }
    },
  };
}

export class GraphAuthProvider implements AuthenticationProvider {
  private pca: PublicClientApplication;
  private scopes: string[];
  private loginScopes: string[];
  private client: Client;
  private pendingLogin?: Promise<string>;
  private lastLoginError?: unknown;

  constructor(config: AuthConfig) {
    this.scopes = config.scopes ?? DEFAULT_SCOPES;
    this.loginScopes = config.loginScopes ?? this.scopes;
    const cachePath = config.cachePath ?? path.join(process.cwd(), '.token-cache.json');
    fs.mkdirSync(path.dirname(cachePath), { recursive: true, mode: 0o700 });

    this.pca = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId ?? 'common'}`,
      },
      cache: { cachePlugin: createCachePlugin(cachePath) },
    });

    this.client = Client.initWithMiddleware({ authProvider: this });
  }

  /** Acquire a token silently from the cache; never prompts. */
  async getAccessToken(): Promise<string> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    if (accounts.length === 0) {
      throw new AuthRequiredError();
    }

    let result;
    try {
      result = await this.pca.acquireTokenSilent({
        account: accounts[0],
        scopes: this.scopes,
      });
    } catch (error) {
      // An expired or revoked refresh token: the cached sign-in is unusable.
      if (error instanceof InteractionRequiredAuthError) {
        throw new AuthRequiredError('Your Microsoft 365 sign-in has expired.');
      }
      throw error;
    }

    if (!result?.accessToken) {
      throw new AuthRequiredError('Your Microsoft 365 sign-in could not be refreshed.');
    }

    return result.accessToken;
  }

  /** The signed-in account's username, if there is a cached sign-in. */
  async getUsername(): Promise<string | undefined> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    return accounts[0]?.username;
  }

  /**
   * Start a device-code sign-in in the background and resolve with the
   * instructions to show the user ("go to https://microsoft.com/devicelogin
   * and enter ABCD-1234"). This lets an MCP client sign in from a tool call,
   * with no terminal: the tool returns the instructions, MSAL keeps polling,
   * and the next call after the user finishes finds the cached token.
   * Repeated calls while a sign-in is pending return the same instructions.
   * A failure before a code is issued (a bad client ID, say) rejects.
   */
  beginDeviceLogin(): Promise<string> {
    if (!this.pendingLogin) {
      let showCode!: (message: string) => void;
      const message = new Promise<string>((resolve) => (showCode = resolve));
      const done = this.pca.acquireTokenByDeviceCode({
        scopes: this.loginScopes,
        deviceCodeCallback: (info) => showCode(info.message),
      });
      const pending = Promise.race([message, done.then(() => 'Signed in.')]);
      this.pendingLogin = pending;
      done.then(
        () => (this.pendingLogin = undefined),
        (error) => {
          this.pendingLogin = undefined;
          this.lastLoginError = error;
        },
      );
      // The race above surfaces an early failure; a later one is reported
      // through takeLoginError instead, so this rejection is handled.
      pending.catch(() => (this.lastLoginError = undefined));
    }
    return this.pendingLogin;
  }

  /** The error from the last background sign-in that failed, cleared once read. */
  takeLoginError(): unknown {
    const error = this.lastLoginError;
    this.lastLoginError = undefined;
    return error;
  }

  /**
   * Interactive device-code login. Prints a URL + code to stderr and waits for
   * the user to complete sign-in in a browser. Run once to seed the token cache.
   */
  async login(): Promise<void> {
    const result = await this.pca.acquireTokenByDeviceCode({
      scopes: this.loginScopes,
      deviceCodeCallback: (info) => {
        // Print to stderr so it never corrupts the stdio MCP channel.
        console.error('\n' + info.message + '\n');
      },
    });

    if (!result?.account) {
      throw new Error('Device-code login did not return an account.');
    }

    console.error(`Signed in as ${result.account.username}. Token cache saved.`);
  }

  getClient(): Client {
    return this.client;
  }
}
