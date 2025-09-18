import { Client } from '@microsoft/microsoft-graph-client';
import { Message, Attachment } from '@microsoft/microsoft-graph-types';

export class MessageTools {
  constructor(private client: Client) {}

  async listMessages(params?: {
    filter?: string;
    select?: string;
    top?: number;
    skip?: number;
    orderBy?: string;
    folderId?: string;
  }): Promise<Message[]> {
    let endpoint = params?.folderId
      ? `/me/mailFolders/${params.folderId}/messages`
      : '/me/messages';

    const request = this.client.api(endpoint);

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    if (params?.top) request.top(params.top);
    if (params?.skip) request.skip(params.skip);
    if (params?.orderBy) request.orderby(params.orderBy);

    const response = await request.get();
    return response.value;
  }

  async getMessage(messageId: string, params?: { select?: string; expand?: string }): Promise<Message> {
    const request = this.client.api(`/me/messages/${messageId}`);

    if (params?.select) request.select(params.select);
    if (params?.expand) request.expand(params.expand);

    return await request.get();
  }

  async createMessage(message: Partial<Message>): Promise<Message> {
    return await this.client.api('/me/messages').post(message);
  }

  async updateMessage(messageId: string, updates: Partial<Message>): Promise<Message> {
    return await this.client.api(`/me/messages/${messageId}`).patch(updates);
  }

  async deleteMessage(messageId: string): Promise<void> {
    await this.client.api(`/me/messages/${messageId}`).delete();
  }

  async sendMessage(messageId: string): Promise<void> {
    await this.client.api(`/me/messages/${messageId}/send`).post({});
  }

  async sendMail(message: {
    message: Partial<Message>;
    saveToSentItems?: boolean;
  }): Promise<void> {
    await this.client.api('/me/sendMail').post(message);
  }

  async replyToMessage(
    messageId: string,
    comment?: string,
    message?: Partial<Message>,
  ): Promise<void> {
    await this.client.api(`/me/messages/${messageId}/reply`).post({
      comment,
      message,
    });
  }

  async replyAllToMessage(
    messageId: string,
    comment?: string,
    message?: Partial<Message>,
  ): Promise<void> {
    await this.client.api(`/me/messages/${messageId}/replyAll`).post({
      comment,
      message,
    });
  }

  async forwardMessage(
    messageId: string,
    toRecipients: any[],
    comment?: string,
    message?: Partial<Message>,
  ): Promise<void> {
    await this.client.api(`/me/messages/${messageId}/forward`).post({
      toRecipients,
      comment,
      message,
    });
  }

  async copyMessage(messageId: string, destinationId: string): Promise<Message> {
    return await this.client.api(`/me/messages/${messageId}/copy`).post({
      destinationId,
    });
  }

  async moveMessage(messageId: string, destinationId: string): Promise<Message> {
    return await this.client.api(`/me/messages/${messageId}/move`).post({
      destinationId,
    });
  }

  async permanentDeleteMessage(messageId: string): Promise<void> {
    await this.client.api(`/me/messages/${messageId}/permanentDelete`).post({});
  }

  async searchMessages(searchTerm: string, params?: {
    select?: string;
    top?: number;
    skip?: number;
  }): Promise<Message[]> {
    const request = this.client.api('/me/messages').search(searchTerm);

    if (params?.select) request.select(params.select);
    if (params?.top) request.top(params.top);
    if (params?.skip) request.skip(params.skip);

    const response = await request.get();
    return response.value;
  }

  async getMessageDelta(params?: {
    deltaToken?: string;
    skipToken?: string;
  }): Promise<{
    value: Message[];
    deltaLink?: string;
    nextLink?: string;
  }> {
    let endpoint = '/me/messages/delta';

    if (params?.deltaToken) {
      endpoint += `?$deltatoken=${params.deltaToken}`;
    } else if (params?.skipToken) {
      endpoint += `?$skiptoken=${params.skipToken}`;
    }

    return await this.client.api(endpoint).get();
  }

  async listAttachments(messageId: string): Promise<Attachment[]> {
    const response = await this.client.api(`/me/messages/${messageId}/attachments`).get();
    return response.value;
  }

  async getAttachment(messageId: string, attachmentId: string): Promise<Attachment> {
    return await this.client.api(`/me/messages/${messageId}/attachments/${attachmentId}`).get();
  }

  async addAttachment(messageId: string, attachment: Partial<Attachment>): Promise<Attachment> {
    return await this.client.api(`/me/messages/${messageId}/attachments`).post(attachment);
  }

  async deleteAttachment(messageId: string, attachmentId: string): Promise<void> {
    await this.client.api(`/me/messages/${messageId}/attachments/${attachmentId}`).delete();
  }
}