import {
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
  cachePath?: string;
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
  private client: Client;

  constructor(config: AuthConfig) {
    this.scopes = config.scopes ?? DEFAULT_SCOPES;
    const cachePath = config.cachePath ?? path.join(process.cwd(), '.token-cache.json');

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
      throw new Error(
        'Not authenticated. Run `npm run login` in the outlook-mcp directory to sign in once.',
      );
    }

    const result = await this.pca.acquireTokenSilent({
      account: accounts[0],
      scopes: this.scopes,
    });

    if (!result?.accessToken) {
      throw new Error(
        'Failed to acquire access token silently. Run `npm run login` to re-authenticate.',
      );
    }

    return result.accessToken;
  }

  /**
   * Interactive device-code login. Prints a URL + code to stderr and waits for
   * the user to complete sign-in in a browser. Run once to seed the token cache.
   */
  async login(): Promise<void> {
    const result = await this.pca.acquireTokenByDeviceCode({
      scopes: this.scopes,
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
