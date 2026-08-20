import { MessageTools } from '../tools/message-tools';
import { SIGNATURE_HTML } from '../signature';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

// Mirror of the font wrapper withSignature() applies to HTML bodies.
const wrappedHtmlBody = (content: string): string =>
  `<div style="font-family:Arial,Helvetica,sans-serif; font-size:11pt">${content}</div>`;

describe('MessageTools', () => {
  let messageTools: MessageTools;
  let mockClient: Client;
  let mockRequest: any;

  beforeEach(() => {
    mockRequest = {
      get: jest.fn(),
      post: jest.fn().mockResolvedValue(undefined),
      patch: jest.fn().mockResolvedValue(undefined),
      delete: jest.fn().mockResolvedValue(undefined),
      filter: jest.fn().mockReturnThis(),
      select: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      skip: jest.fn().mockReturnThis(),
      orderby: jest.fn().mockReturnThis(),
      expand: jest.fn().mockReturnThis(),
      search: jest.fn().mockReturnThis(),
    };

    mockClient = { api: jest.fn(() => mockRequest) } as any;
    messageTools = new MessageTools(mockClient);
  });

  describe('listMessages', () => {
    it('lists from inbox/messages by default and caps top at the default', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: '1' }] });
      const result = await messageTools.listMessages();
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages');
      expect(mockRequest.top).toHaveBeenCalledWith(25);
      expect(result.value).toEqual([{ id: '1' }]);
    });

    it('lists from a specific folder', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.listMessages({ folderId: 'f1' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/f1/messages');
    });

    it('clamps top to the maximum of 100', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.listMessages({ top: 5000 });
      expect(mockRequest.top).toHaveBeenCalledWith(100);
    });

    it('applies all query parameters', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.listMessages({
        filter: 'isRead eq false',
        select: 'id,subject',
        top: 10,
        skip: 5,
        orderBy: 'receivedDateTime desc',
      });
      expect(mockRequest.filter).toHaveBeenCalledWith('isRead eq false');
      expect(mockRequest.select).toHaveBeenCalledWith('id,subject');
      expect(mockRequest.top).toHaveBeenCalledWith(10);
      expect(mockRequest.skip).toHaveBeenCalledWith(5);
      expect(mockRequest.orderby).toHaveBeenCalledWith('receivedDateTime desc');
    });

    it('passes through pagination metadata', async () => {
      mockRequest.get.mockResolvedValue({
        value: [{ id: '1' }],
        '@odata.nextLink': 'next',
        '@odata.count': 42,
      });
      const result = await messageTools.listMessages();
      expect(result['@odata.nextLink']).toBe('next');
      expect(result['@odata.count']).toBe(42);
    });
  });

  describe('getMessage', () => {
    it('gets a message and applies select/expand', async () => {
      mockRequest.get.mockResolvedValue({ id: '1' });
      const result = await messageTools.getMessage('1', {
        select: 'subject',
        expand: 'attachments',
      });
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1');
      expect(mockRequest.select).toHaveBeenCalledWith('subject');
      expect(mockRequest.expand).toHaveBeenCalledWith('attachments');
      expect(result).toEqual({ id: '1' });
    });
  });

  describe('createMessage', () => {
    it('creates a draft and appends the signature to an HTML body', async () => {
      mockRequest.post.mockResolvedValue({ id: '123' });
      await messageTools.createMessage({
        subject: 'Hi',
        body: { contentType: 'html', content: '<p>Body</p>' },
      });
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages');
      const posted = mockRequest.post.mock.calls[0][0];
      expect(posted.subject).toBe('Hi');
      expect(posted.body.content).toBe(wrappedHtmlBody('<p>Body</p>') + SIGNATURE_HTML);
    });

    it('leaves a body undefined when none is provided', async () => {
      mockRequest.post.mockResolvedValue({ id: '123' });
      await messageTools.createMessage({ subject: 'No body' });
      const posted = mockRequest.post.mock.calls[0][0];
      expect(posted.body).toBeUndefined();
    });
  });

  describe('updateMessage', () => {
    it('patches a message', async () => {
      mockRequest.patch.mockResolvedValue({ id: '1', isRead: true });
      const result = await messageTools.updateMessage('1', { isRead: true });
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1');
      expect(mockRequest.patch).toHaveBeenCalledWith({ isRead: true });
      expect(result).toEqual({ id: '1', isRead: true });
    });
  });

  describe('deleteMessage', () => {
    it('deletes a message', async () => {
      await messageTools.deleteMessage('123');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/123');
      expect(mockRequest.delete).toHaveBeenCalled();
    });
  });

  describe('sendMessage', () => {
    it('sends an existing draft', async () => {
      await messageTools.sendMessage('123');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/123/send');
      expect(mockRequest.post).toHaveBeenCalledWith({});
    });
  });

  describe('sendMail', () => {
    it('appends the signature to the outgoing HTML body', async () => {
      await messageTools.sendMail({
        message: {
          subject: 'Test',
          body: { contentType: 'html', content: '<p>Hello</p>' },
          toRecipients: [{ emailAddress: { address: 'a@b.com' } }],
        },
        saveToSentItems: true,
      });
      expect(mockClient.api).toHaveBeenCalledWith('/me/sendMail');
      const posted = mockRequest.post.mock.calls[0][0];
      expect(posted.saveToSentItems).toBe(true);
      expect(posted.message.body.content).toBe(wrappedHtmlBody('<p>Hello</p>') + SIGNATURE_HTML);
      expect(posted.message.subject).toBe('Test');
    });

    it('converts a plain-text body to HTML with the styled signature', async () => {
      await messageTools.sendMail({
        message: { subject: 'Test', body: { contentType: 'text', content: 'Hello' } },
      });
      const posted = mockRequest.post.mock.calls[0][0];
      expect(posted.message.body.contentType).toBe('html');
      expect(posted.message.body.content).toContain('Hello<br>');
      expect(posted.message.body.content).toContain(SIGNATURE_HTML);
    });
  });

  describe('reply / replyAll / forward', () => {
    const draftBody = {
      contentType: 'html',
      content: '<html><body><div id="divRplyFwdMsg">original</div></body></html>',
    };

    beforeEach(() => {
      // First post() returns the created draft; later posts (send) return undefined.
      mockRequest.post.mockReset();
      mockRequest.post
        .mockResolvedValueOnce({ id: 'draft1', body: draftBody })
        .mockResolvedValue(undefined);
      mockRequest.patch.mockResolvedValue(undefined);
    });

    it('replyToMessage creates a draft, injects the signature, and sends', async () => {
      await messageTools.replyToMessage('m1', 'Thanks!');

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/m1/createReply');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/draft1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/draft1/send');

      const patched = mockRequest.patch.mock.calls[0][0];
      expect(patched.body.content).toContain('Thanks!');
      expect(patched.body.content).toContain('geoffreychallen.com');
      // Signature sits above the quoted thread.
      expect(patched.body.content.indexOf('geoffreychallen.com')).toBeLessThan(
        patched.body.content.indexOf('divRplyFwdMsg'),
      );
    });

    it('replyAllToMessage uses the createReplyAll action', async () => {
      await messageTools.replyAllToMessage('m1', 'Got it');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/m1/createReplyAll');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/draft1/send');
    });

    it('forwardMessage sets recipients on the draft and forwards', async () => {
      const recipients = [{ emailAddress: { address: 'fwd@example.com' } }];
      await messageTools.forwardMessage('m1', recipients, 'FYI');

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/m1/createForward');
      const patched = mockRequest.patch.mock.calls[0][0];
      expect(patched.toRecipients).toEqual(recipients);
      expect(patched.body.content).toContain('FYI');
      expect(patched.body.content).toContain('geoffreychallen.com');
    });

    it('replyAllToMessageDraft creates + injects a draft but does NOT send', async () => {
      const draft = await messageTools.replyAllToMessageDraft('m1', 'Got it');

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/m1/createReplyAll');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/draft1');
      // Never sends.
      expect(mockClient.api).not.toHaveBeenCalledWith('/me/messages/draft1/send');

      const patched = mockRequest.patch.mock.calls[0][0];
      expect(patched.body.content).toContain('Got it');
      expect(patched.body.content).toContain('geoffreychallen.com');
      // Returns the draft with a stable id for later review/send.
      expect(draft.id).toBe('draft1');
    });

    it('replyToMessageDraft with html=true preserves raw markup (no escaping)', async () => {
      await messageTools.replyToMessageDraft('m1', '<ul><li>one</li></ul>', true);

      const patched = mockRequest.patch.mock.calls[0][0];
      expect(patched.body.content).toContain('<ul><li>one</li></ul>');
      expect(patched.body.content).not.toContain('&lt;ul&gt;');
      expect(mockClient.api).not.toHaveBeenCalledWith('/me/messages/draft1/send');
    });

    it('replyToMessageDraft with plain text escapes markup', async () => {
      await messageTools.replyToMessageDraft('m1', '<b>hi</b>');

      const patched = mockRequest.patch.mock.calls[0][0];
      expect(patched.body.content).toContain('&lt;b&gt;hi&lt;/b&gt;');
    });
  });

  describe('moveMessage / copyMessage', () => {
    it('moves a message', async () => {
      mockRequest.post.mockResolvedValue({ id: 'new' });
      const result = await messageTools.moveMessage('1', 'dest');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1/move');
      expect(mockRequest.post).toHaveBeenCalledWith({ destinationId: 'dest' });
      expect(result).toEqual({ id: 'new' });
    });

    it('copies a message', async () => {
      mockRequest.post.mockResolvedValue({ id: 'copy' });
      const result = await messageTools.copyMessage('1', 'dest');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1/copy');
      expect(mockRequest.post).toHaveBeenCalledWith({ destinationId: 'dest' });
      expect(result).toEqual({ id: 'copy' });
    });
  });

  describe('moveFolderMessages', () => {
    it('batch-moves a full page, then stops when the folder is drained', async () => {
      mockRequest.get
        .mockResolvedValueOnce({ value: [{ id: 'a' }, { id: 'b' }] })
        .mockResolvedValueOnce({ value: [] });
      mockRequest.post.mockResolvedValueOnce({
        responses: [
          { id: '0', status: 201 },
          { id: '1', status: 201 },
        ],
      });

      const result = await messageTools.moveFolderMessages('src', 'dest');

      expect(result).toEqual({ moved: 2, failed: [], batches: 1 });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/src/messages');
      expect(mockClient.api).toHaveBeenCalledWith('/$batch');
      expect(mockRequest.post).toHaveBeenCalledWith({
        requests: [
          {
            id: '0',
            method: 'POST',
            url: '/me/messages/a/move',
            body: { destinationId: 'dest' },
            headers: { 'Content-Type': 'application/json' },
          },
          {
            id: '1',
            method: 'POST',
            url: '/me/messages/b/move',
            body: { destinationId: 'dest' },
            headers: { 'Content-Type': 'application/json' },
          },
        ],
      });
    });

    it('honors a limit and never lists more than the remaining budget', async () => {
      mockRequest.get.mockResolvedValueOnce({ value: [{ id: 'a' }] });
      mockRequest.post.mockResolvedValueOnce({ responses: [{ id: '0', status: 201 }] });

      const result = await messageTools.moveFolderMessages('src', 'dest', { limit: 1 });

      expect(result).toEqual({ moved: 1, failed: [], batches: 1 });
      expect(mockRequest.top).toHaveBeenCalledWith(1); // budget, not the batch max
      expect(mockRequest.get).toHaveBeenCalledTimes(1); // stopped once the limit was hit
    });

    it('leaves a failed message in the folder and moves it on the next round', async () => {
      mockRequest.get
        .mockResolvedValueOnce({ value: [{ id: 'a' }, { id: 'b' }] })
        .mockResolvedValueOnce({ value: [{ id: 'b' }] })
        .mockResolvedValueOnce({ value: [] });
      mockRequest.post
        .mockResolvedValueOnce({
          responses: [
            { id: '0', status: 201 },
            { id: '1', status: 429 }, // b throttled
          ],
        })
        .mockResolvedValueOnce({ responses: [{ id: '0', status: 201 }] }); // b retried

      const result = await messageTools.moveFolderMessages('src', 'dest');

      expect(result).toEqual({ moved: 2, failed: [], batches: 2 });
    });

    it('gives up after repeated zero-progress rounds and reports the failures', async () => {
      jest.useFakeTimers();
      try {
        mockRequest.get.mockResolvedValue({ value: [{ id: 'a' }] });
        mockRequest.post.mockResolvedValue({ responses: [{ id: '0', status: 403 }] });

        const promise = messageTools.moveFolderMessages('src', 'dest');
        await jest.runAllTimersAsync(); // flush the stall back-off sleeps
        const result = await promise;

        expect(result).toEqual({ moved: 0, failed: [{ id: 'a', status: 403 }], batches: 3 });
      } finally {
        jest.useRealTimers();
      }
    });
  });

  describe('permanentDeleteMessage', () => {
    it('permanently deletes a message', async () => {
      await messageTools.permanentDeleteMessage('1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1/permanentDelete');
      expect(mockRequest.post).toHaveBeenCalledWith({});
    });
  });

  describe('searchMessages', () => {
    it('wraps a bare term in the outer quotes Graph requires and clamps to the search max of 50', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: '1' }] });
      const result = await messageTools.searchMessages('hello', { top: 999 });
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages');
      expect(mockRequest.search).toHaveBeenCalledWith('"hello"');
      expect(mockRequest.top).toHaveBeenCalledWith(50);
      expect(result.value).toEqual([{ id: '1' }]);
    });

    it('wraps a property-scoped query without inner-quoting a simple value', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.searchMessages('from:a@x.com received>=2024-01-01');
      expect(mockRequest.search).toHaveBeenCalledWith('"from:a@x.com received>=2024-01-01"');
    });

    it('escapes inner quotes on phrase values when wrapping', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.searchMessages('subject:"weekly digest"');
      expect(mockRequest.search).toHaveBeenCalledWith('"subject:\\"weekly digest\\""');
    });

    it('passes through a query already wrapped in outer quotes', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.searchMessages('"subject:\\"weekly digest\\""');
      expect(mockRequest.search).toHaveBeenCalledWith('"subject:\\"weekly digest\\""');
    });

    it('scopes the search to a folder when folderId is given', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.searchMessages('from:a@x.com', { folderId: 'inbox' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/inbox/messages');
      expect(mockRequest.search).toHaveBeenCalledWith('"from:a@x.com"');
    });
  });

  describe('getMessageDelta', () => {
    it('uses the plain delta endpoint with no token', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.getMessageDelta();
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/delta');
    });

    it('includes a delta token when provided', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.getMessageDelta({ deltaToken: 'abc' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/delta?$deltatoken=abc');
    });

    it('includes a skip token when provided', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await messageTools.getMessageDelta({ skipToken: 'xyz' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/delta?$skiptoken=xyz');
    });
  });

  describe('attachments', () => {
    it('lists attachments and clamps top to the max of 25', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: 'a' }] });
      await messageTools.listAttachments('1', { top: 100 });
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1/attachments');
      expect(mockRequest.top).toHaveBeenCalledWith(25);
    });

    it('gets a single attachment', async () => {
      mockRequest.get.mockResolvedValue({ id: 'a' });
      const result = await messageTools.getAttachment('1', 'a');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1/attachments/a');
      expect(result).toEqual({ id: 'a' });
    });

    it('adds an attachment', async () => {
      const attachment = { name: 'file.txt' };
      mockRequest.post.mockResolvedValue({ id: 'a' });
      await messageTools.addAttachment('1', attachment);
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1/attachments');
      expect(mockRequest.post).toHaveBeenCalledWith(attachment);
    });

    it('deletes an attachment', async () => {
      await messageTools.deleteAttachment('1', 'a');
      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1/attachments/a');
      expect(mockRequest.delete).toHaveBeenCalled();
    });
  });
});
