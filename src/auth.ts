import { ClientSecretCredential, DeviceCodeCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

export interface AuthConfig {
  clientId: string;
  tenantId?: string;
  clientSecret?: string;
  scopes?: string[];
}

export class GraphAuthProvider {
  private client: Client;
  private config: AuthConfig;

  constructor(config: AuthConfig) {
    this.config = {
      scopes: ['Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.ReadWrite'],
      ...config,
    };

    const credential = this.createCredential();
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: this.config.scopes!,
    });

    this.client = Client.initWithMiddleware({
      authProvider,
    });
  }

  private createCredential() {
    if (this.config.clientSecret && this.config.tenantId) {
      return new ClientSecretCredential(
        this.config.tenantId,
        this.config.clientId,
        this.config.clientSecret,
      );
    } else {
      return new DeviceCodeCredential({
        clientId: this.config.clientId,
        tenantId: this.config.tenantId || 'common',
        userPromptCallback: (info) => {
          console.log(info.message);
        },
      });
    }
  }

  getClient(): Client {
    return this.client;
  }
}