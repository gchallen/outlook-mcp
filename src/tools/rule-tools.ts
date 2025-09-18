import { Client } from '@microsoft/microsoft-graph-client';
import { MessageRule } from '@microsoft/microsoft-graph-types';

export class RuleTools {
  constructor(private client: Client) {}

  async listMessageRules(): Promise<MessageRule[]> {
    const response = await this.client.api('/me/mailFolders/inbox/messageRules').get();
    return response.value;
  }

  async getMessageRule(ruleId: string): Promise<MessageRule> {
    return await this.client.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).get();
  }

  async createMessageRule(rule: Partial<MessageRule>): Promise<MessageRule> {
    return await this.client.api('/me/mailFolders/inbox/messageRules').post(rule);
  }

  async updateMessageRule(ruleId: string, updates: Partial<MessageRule>): Promise<MessageRule> {
    return await this.client.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).patch(updates);
  }

  async deleteMessageRule(ruleId: string): Promise<void> {
    await this.client.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).delete();
  }
}