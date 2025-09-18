import { Client } from '@microsoft/microsoft-graph-client';
import { MailFolder } from '@microsoft/microsoft-graph-types';

export class FolderTools {
  constructor(private client: Client) {}

  async listMailFolders(params?: {
    filter?: string;
    select?: string;
    top?: number;
    skip?: number;
  }): Promise<MailFolder[]> {
    const request = this.client.api('/me/mailFolders');

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    if (params?.top) request.top(params.top);
    if (params?.skip) request.skip(params.skip);

    const response = await request.get();
    return response.value;
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
  ): Promise<MailFolder[]> {
    const request = this.client.api(`/me/mailFolders/${folderId}/childFolders`);

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    if (params?.top) request.top(params.top);
    if (params?.skip) request.skip(params.skip);

    const response = await request.get();
    return response.value;
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

  async getFolderDelta(params?: {
    deltaToken?: string;
    skipToken?: string;
  }): Promise<{
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