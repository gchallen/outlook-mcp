import { Client } from '@microsoft/microsoft-graph-client';
import { Message, Attachment } from '@microsoft/microsoft-graph-types';
import { withSignature, withOutgoingFont, injectReplyTop } from '../signature.js';
import { readFile } from 'fs/promises';
import { basename, extname } from 'path';

// Threshold above which a file must be uploaded via a Graph upload session
// rather than a single POST (3 MiB).
const SIMPLE_ATTACHMENT_MAX_BYTES = 3145728;

// Upload-session chunk size. Must be a multiple of 320 KiB (327680 bytes);
// this is ~4.7 MiB (15 * 320 KiB).
const UPLOAD_CHUNK_SIZE = 15 * 327680;

// Minimal extension -> MIME map for locally-attached files.
const EXTENSION_MIME_TYPES: Record<string, string> = {
  '.zip': 'application/zip',
  '.pdf': 'application/pdf',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.html': 'text/html',
  '.webm': 'video/webm',
};

function contentTypeForFile(filePath: string): string {
  return EXTENSION_MIME_TYPES[extname(filePath).toLowerCase()] ?? 'application/octet-stream';
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

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

  /**
   * PATCH a message. An HTML body gets the outgoing font and paragraph spacing
   * stamped on it, the same as the create path, so revising a draft does not
   * quietly change its typography. The signature is still the caller's to
   * include: a PATCH usually replaces the whole body, and re-appending one
   * would fight with a body that already carries it.
   */
  async updateMessage(messageId: string, updates: Partial<Message>): Promise<Message> {
    const patch = updates.body
      ? { ...updates, body: withOutgoingFont(updates.body as never) as never }
      : updates;
    return await this.client.api(`/me/messages/${messageId}`).patch(patch);
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
    return await this.createReplyDraft(messageId, 'createReplyAll', comment, message, {
      raw: html,
    });
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

  /**
   * Move every message out of `sourceFolderId` into `destinationId` in bulk,
   * using Graph's `/$batch` endpoint (up to 20 moves per request) instead of one
   * round-trip per message. Built for retiring a folder that holds hundreds of
   * messages, which the web UI struggles to move in one action.
   *
   * Each round re-lists from the *top* of the source folder rather than paging
   * with `$skip`: a successful move both removes the message from the folder and
   * rotates its id, so any skip offset would be wrong on the next page. Messages
   * that fail to move stay in the folder and are naturally retried on the next
   * round; `stalls` guards against spinning forever when nothing can progress
   * (e.g. a persistent error or sustained throttling), and `failedById` reports
   * only the ids that were still failing when we stopped.
   */
  async moveFolderMessages(
    sourceFolderId: string,
    destinationId: string,
    opts?: { limit?: number; batchSize?: number },
  ): Promise<{ moved: number; failed: { id: string; status: number }[]; batches: number }> {
    const BATCH_MAX = 20; // Graph $batch hard limit
    const batchSize = Math.min(Math.max(opts?.batchSize ?? BATCH_MAX, 1), BATCH_MAX);
    const limit = opts?.limit;
    const MAX_STALLS = 3; // consecutive zero-progress rounds before giving up

    const failedById = new Map<string, number>();
    let moved = 0;
    let batches = 0;
    let stalls = 0;

    for (;;) {
      if (limit !== undefined && moved >= limit) break;

      const budget = limit !== undefined ? limit - moved : batchSize;
      const pageTop = Math.min(batchSize, budget);

      const page: { value?: { id?: string }[] } = await this.client
        .api(`/me/mailFolders/${sourceFolderId}/messages`)
        .select('id')
        .top(pageTop)
        .get();

      const ids = (page.value ?? [])
        .map((m) => m.id)
        .filter((id): id is string => typeof id === 'string');
      if (ids.length === 0) break; // folder drained

      const requests = ids.map((id, i) => ({
        id: String(i),
        method: 'POST',
        url: `/me/messages/${id}/move`,
        body: { destinationId },
        headers: { 'Content-Type': 'application/json' },
      }));

      const res: { responses?: { id: string; status: number }[] } = await this.client
        .api('/$batch')
        .post({ requests });
      batches++;

      let succeeded = 0;
      for (const r of res.responses ?? []) {
        const id = ids[Number(r.id)];
        if (id === undefined) continue;
        if (r.status >= 200 && r.status < 300) {
          moved++;
          succeeded++;
          failedById.delete(id); // cleared if a prior round had failed it
        } else {
          failedById.set(id, r.status);
        }
      }

      if (succeeded === 0) {
        if (++stalls >= MAX_STALLS) break;
        await sleep(1000 * stalls); // back off before retrying a stalled round
      } else {
        stalls = 0;
      }
    }

    return {
      moved,
      failed: [...failedById].map(([id, status]) => ({ id, status })),
      batches,
    };
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
      folderId?: string;
    },
  ): Promise<{
    value: Message[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const endpoint = params?.folderId
      ? `/me/mailFolders/${params.folderId}/messages`
      : '/me/messages';
    const request = this.client.api(endpoint).search(searchTerm);

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

  /**
   * Attach a local file to a message/draft by absolute path. Small files
   * (<= 3 MiB) are attached with a single POST; larger files are uploaded in
   * sequential chunks via a Graph upload session (the pre-authenticated
   * uploadUrl must NOT carry a Graph Authorization header).
   */
  async addFileAttachment(
    messageId: string,
    filePath: string,
  ): Promise<{ name: string; size: number; method: 'simple' | 'upload-session' }> {
    const data = await readFile(filePath);
    const name = basename(filePath);
    const size = data.byteLength;
    const contentType = contentTypeForFile(filePath);

    if (size <= SIMPLE_ATTACHMENT_MAX_BYTES) {
      await this.client.api(`/me/messages/${messageId}/attachments`).post({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name,
        contentType,
        contentBytes: data.toString('base64'),
      });
      return { name, size, method: 'simple' };
    }

    // Large file: create an upload session, then PUT the bytes in chunks.
    const session = await this.client
      .api(`/me/messages/${messageId}/attachments/createUploadSession`)
      .post({
        AttachmentItem: {
          attachmentType: 'file',
          name,
          size,
          contentType,
        },
      });

    const uploadUrl: string = session.uploadUrl;
    if (!uploadUrl) {
      throw new Error('createUploadSession did not return an uploadUrl.');
    }

    for (let start = 0; start < size; start += UPLOAD_CHUNK_SIZE) {
      const end = Math.min(start + UPLOAD_CHUNK_SIZE, size) - 1;
      const chunk = data.subarray(start, end + 1);

      // The uploadUrl is pre-authenticated; do NOT attach a Graph/Authorization
      // header.
      const response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(chunk.byteLength),
          'Content-Range': `bytes ${start}-${end}/${size}`,
        },
        body: chunk,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => '');
        throw new Error(
          `Upload session chunk failed (bytes ${start}-${end}/${size}): HTTP ${response.status} ${text}`,
        );
      }
    }

    return { name, size, method: 'upload-session' };
  }
}
