import { TeamsTools, teamsHtmlToText, summarizeTeamsMessage } from '../tools/teams-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('TeamsTools', () => {
  let teamsTools: TeamsTools;
  let mockClient: Client;
  let mockRequest: any;

  beforeEach(() => {
    mockRequest = {
      get: jest.fn(),
      post: jest.fn().mockResolvedValue({}),
      select: jest.fn().mockReturnThis(),
      expand: jest.fn().mockReturnThis(),
      orderby: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
    };
    mockClient = { api: jest.fn(() => mockRequest) } as any;
    teamsTools = new TeamsTools(mockClient);
  });

  // The first two GETs resolve /me and /organization.
  const withMe = () =>
    mockRequest.get
      .mockResolvedValueOnce({ id: 'me' })
      .mockResolvedValueOnce({ value: [{ id: 'tenant' }] });

  describe('listChats', () => {
    const chats = {
      value: [
        {
          id: 'c1',
          chatType: 'oneOnOne',
          members: [
            { userId: 'me', displayName: 'Me' },
            { userId: 'u1', displayName: 'Alice' },
          ],
          viewpoint: { lastMessageReadDateTime: '2026-10-01T00:00:00Z' },
          lastMessagePreview: {
            createdDateTime: '2026-10-02T00:00:00Z',
            from: { user: { id: 'u1', displayName: 'Alice' } },
            body: { contentType: 'html', content: '<p>Hi &amp; bye</p>' },
          },
        },
        {
          id: 'c2',
          chatType: 'group',
          topic: 'Staff',
          members: [{ userId: 'me', displayName: 'Me' }],
          viewpoint: { lastMessageReadDateTime: '2026-10-01T00:00:00Z' },
          lastMessagePreview: {
            createdDateTime: '2026-10-02T00:00:00Z',
            from: { user: { id: 'me', displayName: 'Me' } },
            body: { contentType: 'text', content: 'mine' },
          },
        },
      ],
    };

    it('summarizes chats, excluding me from members and flagging unread', async () => {
      withMe().mockResolvedValueOnce(chats);
      const result = await teamsTools.listChats();
      expect(mockClient.api).toHaveBeenCalledWith('/me/chats');
      expect(mockRequest.expand).toHaveBeenCalledWith('members,lastMessagePreview');
      expect(mockRequest.top).toHaveBeenCalledWith(20);
      expect(result.value[0]).toEqual({
        id: 'c1',
        chatType: 'oneOnOne',
        members: ['Alice'],
        unread: true,
        lastMessage: { createdDateTime: '2026-10-02T00:00:00Z', from: 'Alice', text: 'Hi & bye' },
      });
      // My own last message never makes a chat unread.
      expect(result.value[1]).toMatchObject({ topic: 'Staff', unread: false });
    });

    it('filters to unread chats and clamps top', async () => {
      withMe().mockResolvedValueOnce(chats);
      const result = await teamsTools.listChats({ unreadOnly: true, top: 500 });
      expect(mockRequest.top).toHaveBeenCalledWith(50);
      expect(result.value.map((c: any) => c.id)).toEqual(['c1']);
    });
  });

  describe('getChatMessages', () => {
    it('drops system events and returns oldest first', async () => {
      mockRequest.get.mockResolvedValue({
        value: [
          { id: '3', messageType: 'message', body: { contentType: 'text', content: 'later' } },
          { id: '2', messageType: 'systemEventMessage', body: { content: '' } },
          { id: '1', messageType: 'message', body: { contentType: 'text', content: 'first' } },
        ],
      });
      const result = await teamsTools.getChatMessages('c1');
      expect(mockClient.api).toHaveBeenCalledWith('/chats/c1/messages');
      expect(mockRequest.orderby).toHaveBeenCalledWith('createdDateTime desc');
      expect(result.value.map((m: any) => m.id)).toEqual(['1', '3']);
    });
  });

  describe('sendChatMessage', () => {
    it('sends text as escaped, linked HTML with line breaks', async () => {
      await teamsTools.sendChatMessage('c1', 'a < b\nsee https://x.com');
      expect(mockClient.api).toHaveBeenCalledWith('/chats/c1/messages');
      expect(mockRequest.post).toHaveBeenCalledWith({
        body: {
          contentType: 'html',
          content:
            'a &lt; b<br>see <a href="https://x.com" style="color:rgb(12,100,192)">https://x.com</a>',
        },
      });
    });

    it('passes HTML through untouched', async () => {
      await teamsTools.sendChatMessage('c1', '<b>hi</b>', 'html');
      expect(mockRequest.post).toHaveBeenCalledWith({
        body: { contentType: 'html', content: '<b>hi</b>' },
      });
    });
  });

  describe('createChat', () => {
    it('creates a one-on-one chat including me', async () => {
      withMe();
      mockRequest.post.mockResolvedValue({ id: 'c9', chatType: 'oneOnOne' });
      const result = await teamsTools.createChat(['alice@x.com'], 'ignored');
      const body = mockRequest.post.mock.calls[0][0];
      expect(body.chatType).toBe('oneOnOne');
      expect(body.topic).toBeUndefined();
      expect(body.members.map((m: any) => m['user@odata.bind'])).toEqual([
        "https://graph.microsoft.com/v1.0/users('me')",
        "https://graph.microsoft.com/v1.0/users('alice@x.com')",
      ]);
      expect(result.id).toBe('c9');
    });

    it('creates a named group chat for several people', async () => {
      withMe();
      mockRequest.post.mockResolvedValue({ id: 'c10' });
      await teamsTools.createChat(['a@x.com', 'b@x.com'], 'Project');
      expect(mockRequest.post.mock.calls[0][0]).toMatchObject({
        chatType: 'group',
        topic: 'Project',
      });
    });
  });

  describe('markChatRead', () => {
    it('posts my user and tenant, resolving them once', async () => {
      withMe();
      await teamsTools.markChatRead('c1');
      await teamsTools.markChatRead('c2');
      expect(mockClient.api).toHaveBeenCalledWith('/chats/c2/markChatReadForUser');
      expect(mockRequest.post).toHaveBeenCalledWith({ user: { id: 'me', tenantId: 'tenant' } });
      expect(mockRequest.get).toHaveBeenCalledTimes(2);
    });
  });

  describe('channels', () => {
    const post = (id: string, extra: Record<string, unknown> = {}) => ({
      id,
      messageType: 'message',
      createdDateTime: `2026-10-0${id.slice(-1)}T00:00:00Z`,
      from: { user: { displayName: 'Pat Teacher' } },
      body: { contentType: 'html', content: `<p>post ${id}</p>` },
      ...extra,
    });

    it('lists joined teams', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: 't1', displayName: 'CS 124' }] });
      expect(await teamsTools.listJoinedTeams()).toEqual({
        value: [{ id: 't1', displayName: 'CS 124' }],
      });
      expect(mockClient.api).toHaveBeenCalledWith('/me/joinedTeams');
      expect(mockRequest.select).toHaveBeenCalledWith('id,displayName,description');
    });

    it('lists the channels in a team, tolerating an empty response', async () => {
      mockRequest.get.mockResolvedValue({});
      expect(await teamsTools.listChannels('t1')).toEqual({ value: [] });
      expect(mockClient.api).toHaveBeenCalledWith('/teams/t1/channels');
    });

    it('lists channel posts, skipping system events, with a clamped page size', async () => {
      mockRequest.get.mockResolvedValue({
        value: [post('m1', { subject: 'Exam' }), { id: 'e1', messageType: 'systemEventMessage' }],
      });
      const result = await teamsTools.listChannelMessages('t1', 'c1', { top: 500 });
      expect(mockClient.api).toHaveBeenCalledWith('/teams/t1/channels/c1/messages');
      expect(mockRequest.top).toHaveBeenCalledWith(50);
      expect(result.value).toEqual([
        {
          id: 'm1',
          createdDateTime: '2026-10-01T00:00:00Z',
          from: 'Pat Teacher',
          subject: 'Exam',
          text: 'post m1',
        },
      ]);
    });

    it('returns thread replies oldest first', async () => {
      mockRequest.get.mockResolvedValue({ value: [post('r2'), post('r1')] });
      const result = await teamsTools.getChannelMessageReplies('t1', 'c1', 'm1');
      expect(mockClient.api).toHaveBeenCalledWith('/teams/t1/channels/c1/messages/m1/replies');
      expect(result.value.map((m: any) => m.id)).toEqual(['r1', 'r2']);
    });

    it('reacts to a chat message', async () => {
      await teamsTools.setChatReaction('c1', 'm1', 'like');
      expect(mockClient.api).toHaveBeenCalledWith('/chats/c1/messages/m1/setReaction');
      expect(mockRequest.post).toHaveBeenCalledWith({ reactionType: 'like' });
    });

    it('replies in a channel thread', async () => {
      await teamsTools.replyToChannelMessage('t', 'ch', 'm', 'ok');
      expect(mockClient.api).toHaveBeenCalledWith('/teams/t/channels/ch/messages/m/replies');
      expect(mockRequest.post).toHaveBeenCalledWith({
        body: { contentType: 'html', content: 'ok' },
      });
    });

    it('posts a channel message with a subject', async () => {
      await teamsTools.sendChannelMessage('t', 'ch', 'hello', { subject: 'News' });
      expect(mockRequest.post).toHaveBeenCalledWith({
        body: { contentType: 'html', content: 'hello' },
        subject: 'News',
      });
    });
  });

  describe('searchMessages', () => {
    it('queries chatMessage entities and flattens highlights', async () => {
      mockRequest.post.mockResolvedValue({
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    summary: 'the <c0>budget</c0> is due',
                    resource: {
                      id: 'm1',
                      chatId: 'c1',
                      from: { user: { displayName: 'Alice' } },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
      const result = await teamsTools.searchMessages('budget', { top: 100 });
      expect(mockClient.api).toHaveBeenCalledWith('/search/query');
      expect(mockRequest.post.mock.calls[0][0].requests[0]).toMatchObject({
        entityTypes: ['chatMessage'],
        query: { queryString: 'budget' },
        size: 25,
      });
      expect(result.value).toEqual([
        { id: 'm1', from: 'Alice', chatId: 'c1', summary: 'the budget is due' },
      ]);
    });
  });
});

describe('TeamsTools search fallback', () => {
  it('scans recent chats when Graph search is forbidden', async () => {
    const responses: Record<string, any> = {
      '/me': { id: 'me' },
      '/organization': { value: [{ id: 'tenant' }] },
      '/me/chats': {
        value: [{ id: 'c1', chatType: 'oneOnOne', members: [{ userId: 'u', displayName: 'Ann' }] }],
      },
      '/chats/c1/messages': {
        value: [
          {
            id: '2',
            messageType: 'message',
            createdDateTime: '2026-10-02',
            from: { user: { displayName: 'Ann' } },
            body: { contentType: 'text', content: 'Budget is late' },
          },
          {
            id: '1',
            messageType: 'message',
            createdDateTime: '2026-10-01',
            from: { user: { displayName: 'Bob' } },
            body: { contentType: 'text', content: 'budget draft' },
          },
        ],
      },
    };
    const request = (path: string): any => ({
      get: jest.fn().mockResolvedValue(responses[path]),
      post: jest.fn().mockRejectedValue(Object.assign(new Error('no'), { statusCode: 403 })),
      select: jest.fn().mockReturnThis(),
      expand: jest.fn().mockReturnThis(),
      orderby: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
    });
    const tools = new TeamsTools({ api: jest.fn(request) } as any);
    const result: any = await tools.searchMessages('BUDGET from:ann');
    expect(result.note).toMatch(/ChannelMessage\.Read\.All/);
    expect(result.value).toEqual([
      expect.objectContaining({ id: '2', chatId: 'c1', chat: 'Ann', text: 'Budget is late' }),
    ]);
  });
});

describe('teamsHtmlToText', () => {
  it('flattens mentions, breaks, lists, attachments and entities', () => {
    expect(
      teamsHtmlToText(
        '<p><at id="0">Alice</at> see:</p><ul><li>one</li><li>two&nbsp;&#8212; x</li></ul>' +
          '<attachment id="1"></attachment>line<br>next',
      ),
    ).toBe('@Alice see:\n- one\n- two — x\nline\nnext');
  });
});

describe('summarizeTeamsMessage', () => {
  it('keeps the useful fields and marks deletions', () => {
    expect(
      summarizeTeamsMessage({
        id: '1',
        createdDateTime: 't',
        from: { user: { displayName: 'Bob' } },
        deletedDateTime: 'd',
        body: { contentType: 'html', content: '<p>x</p>' },
        attachments: [{ name: 'a.pdf' }],
        reactions: [{ reactionType: 'like', displayName: '👍' }],
      }),
    ).toEqual({
      id: '1',
      createdDateTime: 't',
      from: 'Bob',
      text: '(deleted)',
      attachments: ['a.pdf'],
      reactions: ['👍'],
    });
  });
});
