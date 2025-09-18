import { MessageTools } from '../tools/message-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('MessageTools', () => {
  let messageTools: MessageTools;
  let mockClient: jest.Mocked<Client>;
  let mockRequest: jest.MockedFunction<any>;

  beforeEach(() => {
    mockRequest = {
      get: jest.fn(),
      post: jest.fn(),
      patch: jest.fn(),
      delete: jest.fn(),
      filter: jest.fn().mockReturnThis(),
      select: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      skip: jest.fn().mockReturnThis(),
      orderby: jest.fn().mockReturnThis(),
      expand: jest.fn().mockReturnThis(),
      search: jest.fn().mockReturnThis(),
    };

    mockClient = {
      api: jest.fn(() => mockRequest),
    } as any;

    messageTools = new MessageTools(mockClient);
  });

  describe('listMessages', () => {
    it('should list messages from inbox by default', async () => {
      const mockMessages = [{ id: '1', subject: 'Test Message' }];
      mockRequest.get.mockResolvedValue({ value: mockMessages });

      const result = await messageTools.listMessages();

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages');
      expect(result).toEqual(mockMessages);
    });

    it('should list messages from specific folder', async () => {
      const mockMessages = [{ id: '1', subject: 'Test Message' }];
      mockRequest.get.mockResolvedValue({ value: mockMessages });

      const result = await messageTools.listMessages({ folderId: 'folder123' });

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/folder123/messages');
      expect(result).toEqual(mockMessages);
    });

    it('should apply query parameters', async () => {
      const mockMessages = [{ id: '1', subject: 'Test Message' }];
      mockRequest.get.mockResolvedValue({ value: mockMessages });

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
  });

  describe('getMessage', () => {
    it('should get a specific message', async () => {
      const mockMessage = { id: '1', subject: 'Test Message' };
      mockRequest.get.mockResolvedValue(mockMessage);

      const result = await messageTools.getMessage('1');

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/1');
      expect(result).toEqual(mockMessage);
    });
  });

  describe('createMessage', () => {
    it('should create a new message', async () => {
      const messageData = { subject: 'New Message', body: { content: 'Test' } };
      const mockCreatedMessage = { id: '123', ...messageData };
      mockRequest.post.mockResolvedValue(mockCreatedMessage);

      const result = await messageTools.createMessage(messageData);

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages');
      expect(mockRequest.post).toHaveBeenCalledWith(messageData);
      expect(result).toEqual(mockCreatedMessage);
    });
  });

  describe('sendMail', () => {
    it('should send mail directly', async () => {
      const mailData = {
        message: {
          subject: 'Test Email',
          body: { content: 'Test content' },
          toRecipients: [{ emailAddress: { address: 'test@example.com' } }],
        },
        saveToSentItems: true,
      };

      await messageTools.sendMail(mailData);

      expect(mockClient.api).toHaveBeenCalledWith('/me/sendMail');
      expect(mockRequest.post).toHaveBeenCalledWith(mailData);
    });
  });

  describe('deleteMessage', () => {
    it('should delete a message', async () => {
      await messageTools.deleteMessage('123');

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages/123');
      expect(mockRequest.delete).toHaveBeenCalled();
    });
  });

  describe('searchMessages', () => {
    it('should search messages', async () => {
      const mockResults = [{ id: '1', subject: 'Found Message' }];
      mockRequest.get.mockResolvedValue({ value: mockResults });

      const result = await messageTools.searchMessages('test query');

      expect(mockClient.api).toHaveBeenCalledWith('/me/messages');
      expect(mockRequest.search).toHaveBeenCalledWith('test query');
      expect(result).toEqual(mockResults);
    });
  });
});