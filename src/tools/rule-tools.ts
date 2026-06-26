import { Client } from '@microsoft/microsoft-graph-client';
import { MessageRule } from '@microsoft/microsoft-graph-types';

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  RULES: { default: 10, max: 25 },
} as const;

export class RuleTools {
  constructor(private client: Client) {}

  async listMessageRules(params?: { top?: number; skip?: number }): Promise<{
    value: MessageRule[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api('/me/mailFolders/inbox/messageRules');

    // Apply default pagination limit to prevent excessive context generation
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.RULES.default,
      PAGINATION_LIMITS.RULES.max,
    );
    request.top(topLimit);
    if (params?.skip) request.skip(params.skip);

    const response = await request.get();
    return {
      value: response.value,
      '@odata.nextLink': response['@odata.nextLink'],
      '@odata.count': response['@odata.count'],
    };
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
