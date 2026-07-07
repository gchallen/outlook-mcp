import { Client } from '@microsoft/microsoft-graph-client';
import { Message, Attachment } from '@microsoft/microsoft-graph-types';
import { withSignature, injectReplyTop } from '../signature.js';

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  MESSAGES: { default: 25, max: 100 },
  SEARCH: { default: 25, max: 50 },
  ATTACHMENTS: { default: 10, max: 25 },
} as const;

export class MessageTools {
  constructor(private client: Client) {}

  async listMessages(params?: {
    filter?: string;
    select?: string;
    top?: number;
    skip?: number;
    orderBy?: string;
    folderId?: string;
  }): Promise<{
    value: Message[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    let endpoint = params?.folderId
      ? `/me/mailFolders/${params.folderId}/messages`
      : '/me/messages';

    const request = this.client.api(endpoint);

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    // Apply default pagination limit to prevent excessive context generation
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.MESSAGES.default,
      PAGINATION_LIMITS.MESSAGES.max,
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

  async getMessage(
    messageId: string,
    params?: { select?: string; expand?: string },
  ): Promise<Message> {
    const request = this.client.api(`/me/messages/${messageId}`);

    if (params?.select) request.select(params.select);
    if (params?.expand) request.expand(params.expand);

    return await request.get();
  }

  async createMessage(message: Partial<Message>): Promise<Message> {
    return await this.client.api('/me/messages').post({
      ...message,
      body: withSignature(message.body as any) ?? message.body,
    });
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

  async sendMail(message: { message: Partial<Message>; saveToSentItems?: boolean }): Promise<void> {
    await this.client.api('/me/sendMail').post({
      ...message,
      message: {
        ...message.message,
        body: withSignature(message.message.body as any) ?? message.message.body,
      },
    });
  }

  /**
   * Create a draft (reply/replyAll/forward), inject the comment + signature
   * above the quoted thread, and optionally set additional properties. The
   * draft is left in the Drafts folder (not sent) and returned. Using the
   * create* actions (rather than the one-shot reply/forward actions) preserves
   * threading and recipients while letting the styled signature land in the body.
   * `opts.raw` treats `comment` as pre-formatted HTML instead of plain text.
   */
  private async createReplyDraft(
    messageId: string,
    action: 'createReply' | 'createReplyAll' | 'createForward',
    comment?: string,
    updates?: Partial<Message>,
    opts?: { raw?: boolean },
  ): Promise<Message> {
    const draft: Message = await this.client.api(`/me/messages/${messageId}/${action}`).post({});

    const body = injectReplyTop(draft.body as any, comment, opts);
    const patched: Message = await this.client.api(`/me/messages/${draft.id}`).patch({
      ...updates,
      body,
    });

    // Graph returns the updated message on PATCH, but fall back to the created
    // draft's id so downstream send/summary always has a stable identifier.
    return { ...draft, ...patched, id: patched?.id ?? draft.id };
  }

  /**
   * As createReplyDraft, but sends the draft after injecting the body.
   */
  private async createInjectAndSend(
    messageId: string,
    action: 'createReply' | 'createReplyAll' | 'createForward',
    comment?: string,
    updates?: Partial<Message>,
  ): Promise<void> {
    const draft = await this.createReplyDraft(messageId, action, comment, updates);
    await this.client.api(`/me/messages/${draft.id}/send`).post({});
  }

  async replyToMessage(
    messageId: string,
    comment?: string,
    message?: Partial<Message>,
  ): Promise<void> {
    await this.createInjectAndSend(messageId, 'createReply', comment, message);
  }

  async replyAllToMessage(
    messageId: string,
    comment?: string,
    message?: Partial<Message>,
  ): Promise<void> {
    await this.createInjectAndSend(messageId, 'createReplyAll', comment, message);
  }

  /**
   * Create a reply draft (to sender only) without sending it. Returns the draft
   * so it can be reviewed, edited, or sent later. `html` treats `comment` as
   * pre-formatted HTML so rich replies keep their formatting.
   */
  async replyToMessageDraft(
    messageId: string,
    comment?: string,
    html?: boolean,
    message?: Partial<Message>,
  ): Promise<Message> {
    return await this.createReplyDraft(messageId, 'createReply', comment, message, { raw: html });
  }

  /**
   * Create a reply-all draft without sending it. Returns the draft so it can be
   * reviewed, edited, or sent later. `html` treats `comment` as pre-formatted
   * HTML so rich replies keep their formatting.
   */
  async replyAllToMessageDraft(
    messageId: string,
    comment?: string,
    html?: boolean,
    message?: Partial<Message>,
  ): Promise<Message> {
    return await this.createReplyDraft(messageId, 'createReplyAll', comment, message, { raw: html });
  }

  async forwardMessage(
    messageId: string,
    toRecipients: any[],
    comment?: string,
    message?: Partial<Message>,
  ): Promise<void> {
    await this.createInjectAndSend(messageId, 'createForward', comment, {
      toRecipients,
      ...message,
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

  async searchMessages(
    searchTerm: string,
    params?: {
      select?: string;
      top?: number;
      skip?: number;
    },
  ): Promise<{
    value: Message[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api('/me/messages').search(searchTerm);

    if (params?.select) request.select(params.select);
    // Apply default pagination limit to prevent excessive context generation
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.SEARCH.default,
      PAGINATION_LIMITS.SEARCH.max,
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

  async getMessageDelta(params?: { deltaToken?: string; skipToken?: string }): Promise<{
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

  async listAttachments(
    messageId: string,
    params?: {
      top?: number;
      skip?: number;
    },
  ): Promise<{
    value: Attachment[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api(`/me/messages/${messageId}/attachments`);

    // Apply default pagination limit to prevent excessive context generation
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.ATTACHMENTS.default,
      PAGINATION_LIMITS.ATTACHMENTS.max,
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
