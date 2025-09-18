import { FolderTools } from '../tools/folder-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('FolderTools', () => {
  let folderTools: FolderTools;
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
      expand: jest.fn().mockReturnThis(),
    };

    mockClient = {
      api: jest.fn(() => mockRequest),
    } as any;

    folderTools = new FolderTools(mockClient);
  });

  describe('listMailFolders', () => {
    it('should list mail folders', async () => {
      const mockFolders = [{ id: '1', displayName: 'Inbox' }];
      const mockResponse = { value: mockFolders, '@odata.nextLink': undefined, '@odata.count': undefined };
      mockRequest.get.mockResolvedValue(mockResponse);

      const result = await folderTools.listMailFolders();

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders');
      expect(result).toEqual({
        value: mockFolders,
        '@odata.nextLink': undefined,
        '@odata.count': undefined,
      });
    });

    it('should apply query parameters', async () => {
      const mockFolders = [{ id: '1', displayName: 'Inbox' }];
      mockRequest.get.mockResolvedValue({ value: mockFolders });

      await folderTools.listMailFolders({
        filter: 'displayName eq \'Inbox\'',
        select: 'id,displayName',
        top: 10,
        skip: 5,
      });

      expect(mockRequest.filter).toHaveBeenCalledWith('displayName eq \'Inbox\'');
      expect(mockRequest.select).toHaveBeenCalledWith('id,displayName');
      expect(mockRequest.top).toHaveBeenCalledWith(10);
      expect(mockRequest.skip).toHaveBeenCalledWith(5);
    });
  });

  describe('getMailFolder', () => {
    it('should get a specific folder', async () => {
      const mockFolder = { id: '1', displayName: 'Inbox' };
      mockRequest.get.mockResolvedValue(mockFolder);

      const result = await folderTools.getMailFolder('1');

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/1');
      expect(result).toEqual(mockFolder);
    });
  });

  describe('createMailFolder', () => {
    it('should create a folder in root', async () => {
      const folderData = { displayName: 'New Folder' };
      const mockCreatedFolder = { id: '123', ...folderData };
      mockRequest.post.mockResolvedValue(mockCreatedFolder);

      const result = await folderTools.createMailFolder(folderData);

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders');
      expect(mockRequest.post).toHaveBeenCalledWith({ displayName: 'New Folder' });
      expect(result).toEqual(mockCreatedFolder);
    });

    it('should create a child folder', async () => {
      const folderData = { displayName: 'Child Folder', parentFolderId: 'parent123' };
      const mockCreatedFolder = { id: '123', displayName: 'Child Folder' };
      mockRequest.post.mockResolvedValue(mockCreatedFolder);

      const result = await folderTools.createMailFolder(folderData);

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/parent123/childFolders');
      expect(mockRequest.post).toHaveBeenCalledWith({ displayName: 'Child Folder' });
      expect(result).toEqual(mockCreatedFolder);
    });
  });

  describe('deleteMailFolder', () => {
    it('should delete a folder', async () => {
      await folderTools.deleteMailFolder('123');

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/123');
      expect(mockRequest.delete).toHaveBeenCalled();
    });
  });

  describe('getWellKnownFolder', () => {
    it('should get inbox folder', async () => {
      const mockFolder = { id: 'inbox', displayName: 'Inbox' };
      mockRequest.get.mockResolvedValue(mockFolder);

      const result = await folderTools.getWellKnownFolder('inbox');

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/inbox');
      expect(result).toEqual(mockFolder);
    });
  });

  describe('listChildFolders', () => {
    it('should list child folders', async () => {
      const mockFolders = [{ id: '1', displayName: 'Child 1' }];
      const mockResponse = { value: mockFolders, '@odata.nextLink': undefined, '@odata.count': undefined };
      mockRequest.get.mockResolvedValue(mockResponse);

      const result = await folderTools.listChildFolders('parent123');

      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/parent123/childFolders');
      expect(result).toEqual({
        value: mockFolders,
        '@odata.nextLink': undefined,
        '@odata.count': undefined,
      });
    });
  });
});