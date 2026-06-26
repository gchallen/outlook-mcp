import { Client } from '@microsoft/microsoft-graph-client';
import { Contact } from '@microsoft/microsoft-graph-types';

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  CONTACTS: { default: 25, max: 100 },
} as const;

export class ContactsTools {
  constructor(private client: Client) {}

  async listContacts(params?: {
    filter?: string;
    select?: string;
    top?: number;
    skip?: number;
    orderBy?: string;
  }): Promise<{
    value: Contact[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api('/me/contacts');

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.CONTACTS.default,
      PAGINATION_LIMITS.CONTACTS.max,
    );
    request.top(topLimit);
    if (params?.skip) request.skip(params.skip);
    if (params?.orderBy) request.orderby(params.orderBy);

    const response = await request.get();
    return {
      value: response.value,
      '@odata.nextLink': response['@odata.nextLink'],
      '@odata.count': response['@odata.count'],
    };
  }

  async searchContacts(
    searchTerm: string,
    params?: {
      select?: string;
      top?: number;
      skip?: number;
    },
  ): Promise<{
    value: Contact[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api('/me/contacts').search(searchTerm);

    if (params?.select) request.select(params.select);
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.CONTACTS.default,
      PAGINATION_LIMITS.CONTACTS.max,
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

  async getContact(contactId: string, params?: { select?: string }): Promise<Contact> {
    const request = this.client.api(`/me/contacts/${contactId}`);
    if (params?.select) request.select(params.select);
    return await request.get();
  }

  async createContact(contact: Partial<Contact>): Promise<Contact> {
    return await this.client.api('/me/contacts').post(contact);
  }

  async updateContact(contactId: string, updates: Partial<Contact>): Promise<Contact> {
    return await this.client.api(`/me/contacts/${contactId}`).patch(updates);
  }

  async deleteContact(contactId: string): Promise<void> {
    await this.client.api(`/me/contacts/${contactId}`).delete();
  }
}
