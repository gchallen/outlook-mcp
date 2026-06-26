import { Client } from '@microsoft/microsoft-graph-client';
import { MailFolder } from '@microsoft/microsoft-graph-types';

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  FOLDERS: { default: 20, max: 50 },
} as const;

export class FolderTools {
  constructor(private client: Client) {}

  async listMailFolders(params?: {
    filter?: string;
    select?: string;
    top?: number;
    skip?: number;
  }): Promise<{
    value: MailFolder[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api('/me/mailFolders');

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    // Apply default pagination limit to prevent excessive context generation
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.FOLDERS.default,
      PAGINATION_LIMITS.FOLDERS.max,
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

  async getMailFolder(
    folderId: string,
    params?: { select?: string; expand?: string },
  ): Promise<MailFolder> {
    const request = this.client.api(`/me/mailFolders/${folderId}`);

    if (params?.select) request.select(params.select);
    if (params?.expand) request.expand(params.expand);

    return await request.get();
  }

  async createMailFolder(folder: {
    displayName: string;
    parentFolderId?: string;
  }): Promise<MailFolder> {
    const endpoint = folder.parentFolderId
      ? `/me/mailFolders/${folder.parentFolderId}/childFolders`
      : '/me/mailFolders';

    return await this.client.api(endpoint).post({
      displayName: folder.displayName,
    });
  }

  async updateMailFolder(folderId: string, updates: Partial<MailFolder>): Promise<MailFolder> {
    return await this.client.api(`/me/mailFolders/${folderId}`).patch(updates);
  }

  async deleteMailFolder(folderId: string): Promise<void> {
    await this.client.api(`/me/mailFolders/${folderId}`).delete();
  }

  async copyMailFolder(folderId: string, destinationId: string): Promise<MailFolder> {
    return await this.client.api(`/me/mailFolders/${folderId}/copy`).post({
      destinationId,
    });
  }

  async moveMailFolder(folderId: string, destinationId: string): Promise<MailFolder> {
    return await this.client.api(`/me/mailFolders/${folderId}/move`).post({
      destinationId,
    });
  }

  async permanentDeleteMailFolder(folderId: string): Promise<void> {
    await this.client.api(`/me/mailFolders/${folderId}/permanentDelete`).post({});
  }

  async listChildFolders(
    folderId: string,
    params?: {
      filter?: string;
      select?: string;
      top?: number;
      skip?: number;
    },
  ): Promise<{
    value: MailFolder[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api(`/me/mailFolders/${folderId}/childFolders`);

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    // Apply default pagination limit to prevent excessive context generation
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.FOLDERS.default,
      PAGINATION_LIMITS.FOLDERS.max,
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

  async getWellKnownFolder(
    folderName: 'inbox' | 'sentitems' | 'deleteditems' | 'drafts' | 'outbox' | 'junkemail',
    params?: { select?: string; expand?: string },
  ): Promise<MailFolder> {
    const request = this.client.api(`/me/mailFolders/${folderName}`);

    if (params?.select) request.select(params.select);
    if (params?.expand) request.expand(params.expand);

    return await request.get();
  }

  async getFolderDelta(params?: { deltaToken?: string; skipToken?: string }): Promise<{
    value: MailFolder[];
    deltaLink?: string;
    nextLink?: string;
  }> {
    let endpoint = '/me/mailFolders/delta';

    if (params?.deltaToken) {
      endpoint += `?$deltatoken=${params.deltaToken}`;
    } else if (params?.skipToken) {
      endpoint += `?$skiptoken=${params.skipToken}`;
    }

    return await this.client.api(endpoint).get();
  }
}
