import { Client } from '@microsoft/microsoft-graph-client';
import { textToHtml } from '../signature.js';

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  CHATS: { default: 20, max: 50 },
  MESSAGES: { default: 20, max: 50 },
  SEARCH: { default: 25, max: 25 },
} as const;

// Bounds for the local search fallback: chats scanned, messages per chat, and
// how many chats are fetched at once.
const SCAN = { CHATS: 50, MESSAGES: 50, CONCURRENCY: 5 } as const;

function clamp(top: number | undefined, limit: { default: number; max: number }): number {
  return Math.min(top || limit.default, limit.max);
}

const ENTITIES: Record<string, string> = {
  '&nbsp;': ' ',
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&#39;': "'",
  '&apos;': "'",
};

/**
 * Flatten a Teams message body to readable text. Teams stores nearly every
 * message as HTML (mentions as <at>, attachments as <attachment> stubs), and
 * passing that through verbatim wastes context on markup.
 */
export function teamsHtmlToText(html: string): string {
  return html
    .replace(/<at[^>]*>(.*?)<\/at>/gi, '@$1')
    .replace(/<attachment[^>]*>\s*<\/attachment>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|h[1-6]|blockquote)>/gi, '\n')
    .replace(/<li[^>]*>/gi, '- ')
    .replace(/<[^>]+>/g, '')
    .replace(/&(nbsp|amp|lt|gt|quot|#39|apos);/g, (e) => ENTITIES[e])
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function senderName(from: any): string | undefined {
  return from?.user?.displayName ?? from?.application?.displayName ?? undefined;
}

/** Trim a Graph chatMessage to the fields worth reading. */
export function summarizeTeamsMessage(m: any) {
  const body = m?.body?.content ?? '';
  const text = m?.body?.contentType === 'html' ? teamsHtmlToText(body) : body.trim();
  const attachments = (m?.attachments ?? [])
    .map((a: any) => a?.name)
    .filter((n: string | undefined) => n);
  const reactions = (m?.reactions ?? []).map((r: any) => r?.displayName ?? r?.reactionType);
  return {
    id: m?.id,
    createdDateTime: m?.createdDateTime,
    from: senderName(m?.from),
    ...(m?.subject ? { subject: m.subject } : {}),
    text: m?.deletedDateTime ? '(deleted)' : text,
    ...(m?.replyToId ? { replyToId: m.replyToId } : {}),
    ...(attachments.length ? { attachments } : {}),
    ...(reactions.length ? { reactions } : {}),
  };
}

function messageBody(content: string, contentType?: 'text' | 'html') {
  // Sent as HTML either way: Teams collapses newlines in a "text" body and
  // does not hyperlink bare URLs in one, so plain text is escaped and linked.
  return {
    body: {
      contentType: 'html',
      content: contentType === 'html' ? content : textToHtml(content),
    },
  };
}

export class TeamsTools {
  private me?: { id: string; tenantId: string };

  constructor(private client: Client) {}

  private async getMe(): Promise<{ id: string; tenantId: string }> {
    if (!this.me) {
      const user = await this.client.api('/me').select('id').get();
      const org = await this.client.api('/organization').select('id').get();
      this.me = { id: user.id, tenantId: org.value?.[0]?.id };
    }
    return this.me;
  }

  /**
   * List the user's chats, most recently active first, with members and a
   * preview of the last message. A chat is unread when its last message is
   * from someone else and newer than the user's read marker.
   */
  async listChats(params?: { top?: number; unreadOnly?: boolean }) {
    const me = await this.getMe();
    const response = await this.client
      .api('/me/chats')
      .expand('members,lastMessagePreview')
      .orderby('lastMessagePreview/createdDateTime desc')
      .top(clamp(params?.top, PAGINATION_LIMITS.CHATS))
      .get();

    const chats = (response.value ?? []).map((chat: any) => {
      const last = chat.lastMessagePreview;
      const readAt = chat.viewpoint?.lastMessageReadDateTime;
      const unread =
        !!last &&
        last.from?.user?.id !== me.id &&
        (!readAt || new Date(last.createdDateTime) > new Date(readAt));
      return {
        id: chat.id,
        chatType: chat.chatType,
        ...(chat.topic ? { topic: chat.topic } : {}),
        members: (chat.members ?? [])
          .filter((member: any) => member.userId !== me.id)
          .map((member: any) => member.displayName ?? member.email),
        unread,
        lastMessage: last
          ? {
              createdDateTime: last.createdDateTime,
              from: senderName(last.from),
              text:
                last.body?.contentType === 'html'
                  ? teamsHtmlToText(last.body?.content ?? '')
                  : last.body?.content,
            }
          : undefined,
      };
    });

    return { value: params?.unreadOnly ? chats.filter((c: any) => c.unread) : chats };
  }

  /** Recent messages in a chat, returned oldest first so they read in order. */
  async getChatMessages(chatId: string, params?: { top?: number }) {
    const response = await this.client
      .api(`/chats/${chatId}/messages`)
      .orderby('createdDateTime desc')
      .top(clamp(params?.top, PAGINATION_LIMITS.MESSAGES))
      .get();
    return {
      value: (response.value ?? [])
        .filter((m: any) => m.messageType === 'message')
        .map(summarizeTeamsMessage)
        .reverse(),
    };
  }

  async sendChatMessage(chatId: string, content: string, contentType?: 'text' | 'html') {
    const message = await this.client
      .api(`/chats/${chatId}/messages`)
      .post(messageBody(content, contentType));
    return summarizeTeamsMessage(message);
  }

  /**
   * Start a chat with the given people: one-on-one for a single address
   * (Graph returns the existing chat if there already is one), group otherwise.
   */
  async createChat(emails: string[], topic?: string) {
    const me = await this.getMe();
    const member = (user: string) => ({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user}')`,
    });
    const oneOnOne = emails.length === 1;
    const chat = await this.client.api('/chats').post({
      chatType: oneOnOne ? 'oneOnOne' : 'group',
      ...(!oneOnOne && topic ? { topic } : {}),
      members: [member(me.id), ...emails.map(member)],
    });
    return { id: chat.id, chatType: chat.chatType, topic: chat.topic, webUrl: chat.webUrl };
  }

  async markChatRead(chatId: string) {
    const me = await this.getMe();
    await this.client
      .api(`/chats/${chatId}/markChatReadForUser`)
      .post({ user: { id: me.id, tenantId: me.tenantId } });
  }

  async setChatReaction(chatId: string, messageId: string, reactionType: string) {
    await this.client
      .api(`/chats/${chatId}/messages/${messageId}/setReaction`)
      .post({ reactionType });
  }

  async listJoinedTeams() {
    const response = await this.client
      .api('/me/joinedTeams')
      .select('id,displayName,description')
      .get();
    return { value: response.value ?? [] };
  }

  async listChannels(teamId: string) {
    const response = await this.client
      .api(`/teams/${teamId}/channels`)
      .select('id,displayName,description,membershipType')
      .get();
    return { value: response.value ?? [] };
  }

  /** Top-level posts in a channel, newest first; replies come separately. */
  async listChannelMessages(teamId: string, channelId: string, params?: { top?: number }) {
    const response = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/messages`)
      .top(clamp(params?.top, PAGINATION_LIMITS.MESSAGES))
      .get();
    return {
      value: (response.value ?? [])
        .filter((m: any) => m.messageType === 'message')
        .map(summarizeTeamsMessage),
    };
  }

  async getChannelMessageReplies(
    teamId: string,
    channelId: string,
    messageId: string,
    params?: { top?: number },
  ) {
    const response = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
      .top(clamp(params?.top, PAGINATION_LIMITS.MESSAGES))
      .get();
    return {
      value: (response.value ?? [])
        .filter((m: any) => m.messageType === 'message')
        .map(summarizeTeamsMessage)
        .reverse(),
    };
  }

  async sendChannelMessage(
    teamId: string,
    channelId: string,
    content: string,
    opts?: { contentType?: 'text' | 'html'; subject?: string },
  ) {
    const message = await this.client.api(`/teams/${teamId}/channels/${channelId}/messages`).post({
      ...messageBody(content, opts?.contentType),
      ...(opts?.subject ? { subject: opts.subject } : {}),
    });
    return summarizeTeamsMessage(message);
  }

  async replyToChannelMessage(
    teamId: string,
    channelId: string,
    messageId: string,
    content: string,
    contentType?: 'text' | 'html',
  ) {
    const message = await this.client
      .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
      .post(messageBody(content, contentType));
    return summarizeTeamsMessage(message);
  }

  /**
   * Full-text search across the user's Teams messages. Graph's search API
   * demands ChannelMessage.Read.All (admin consent) even for chat-only
   * results, so on a 403 this falls back to scanning recent chats locally.
   */
  async searchMessages(query: string, params?: { top?: number }) {
    const size = clamp(params?.top, PAGINATION_LIMITS.SEARCH);
    let response;
    try {
      response = await this.client.api('/search/query').post({
        requests: [{ entityTypes: ['chatMessage'], query: { queryString: query }, from: 0, size }],
      });
    } catch (error: any) {
      if (error?.statusCode !== 403) throw error;
      return this.scanRecentChats(query, size);
    }
    const hits = response.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    return {
      value: hits.map((hit: any) => {
        const r = hit.resource ?? {};
        return {
          id: r.id,
          createdDateTime: r.createdDateTime,
          from: senderName(r.from),
          ...(r.chatId ? { chatId: r.chatId } : {}),
          ...(r.channelIdentity ? { channelIdentity: r.channelIdentity } : {}),
          ...(r.replyToId ? { replyToId: r.replyToId } : {}),
          // Search highlights matches with <c0> tags; flatten them away.
          summary: teamsHtmlToText(hit.summary ?? ''),
        };
      }),
    };
  }

  /**
   * Case-insensitive match of every query term against the latest messages
   * of the most recent chats, newest hits first. A "from:name" term matches
   * the sender instead of the text.
   */
  private async scanRecentChats(query: string, size: number) {
    const terms = query
      .toLowerCase()
      .split(/\s+/)
      .map((t) => t.replace(/"/g, ''))
      .filter((t) => t && !['and', 'or'].includes(t));
    const matches = (m: ReturnType<typeof summarizeTeamsMessage>) =>
      terms.every((t) =>
        t.startsWith('from:')
          ? (m.from ?? '').toLowerCase().includes(t.slice(5))
          : m.text.toLowerCase().includes(t),
      );

    const { value: chats } = await this.listChats({ top: SCAN.CHATS });
    const hits: any[] = [];
    for (let i = 0; i < chats.length; i += SCAN.CONCURRENCY) {
      const batch = chats.slice(i, i + SCAN.CONCURRENCY);
      const results = await Promise.all(
        batch.map((chat: any) =>
          this.getChatMessages(chat.id, { top: SCAN.MESSAGES }).then(
            ({ value }) => ({ chat, value }),
            () => ({ chat, value: [] }),
          ),
        ),
      );
      for (const { chat, value } of results) {
        for (const m of value.filter(matches)) {
          hits.push({ ...m, chatId: chat.id, chat: chat.topic ?? chat.members.join(', ') });
        }
      }
    }
    hits.sort((a, b) => (b.createdDateTime ?? '').localeCompare(a.createdDateTime ?? ''));
    return {
      note: `Graph search needs ChannelMessage.Read.All (admin consent), so this scanned the latest ${SCAN.MESSAGES} messages of your ${chats.length} most recent chats.`,
      value: hits.slice(0, size),
    };
  }
}
