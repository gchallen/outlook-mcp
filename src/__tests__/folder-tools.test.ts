import { FolderTools } from '../tools/folder-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('FolderTools', () => {
  let folderTools: FolderTools;
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
      expand: jest.fn().mockReturnThis(),
    };
    mockClient = { api: jest.fn(() => mockRequest) } as any;
    folderTools = new FolderTools(mockClient);
  });

  describe('listMailFolders', () => {
    it('lists folders with default pagination', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: '1', displayName: 'Inbox' }] });
      const result = await folderTools.listMailFolders();
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders');
      expect(mockRequest.top).toHaveBeenCalledWith(20);
      expect(result.value).toEqual([{ id: '1', displayName: 'Inbox' }]);
    });

    it('applies query parameters and clamps top to 50', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await folderTools.listMailFolders({
        filter: "displayName eq 'Inbox'",
        select: 'id,displayName',
        top: 999,
        skip: 5,
      });
      expect(mockRequest.filter).toHaveBeenCalledWith("displayName eq 'Inbox'");
      expect(mockRequest.select).toHaveBeenCalledWith('id,displayName');
      expect(mockRequest.top).toHaveBeenCalledWith(50);
      expect(mockRequest.skip).toHaveBeenCalledWith(5);
    });
  });

  describe('getMailFolder', () => {
    it('gets a folder and applies select/expand', async () => {
      mockRequest.get.mockResolvedValue({ id: '1' });
      const result = await folderTools.getMailFolder('1', {
        select: 'displayName',
        expand: 'childFolders',
      });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/1');
      expect(mockRequest.select).toHaveBeenCalledWith('displayName');
      expect(mockRequest.expand).toHaveBeenCalledWith('childFolders');
      expect(result).toEqual({ id: '1' });
    });
  });

  describe('createMailFolder', () => {
    it('creates a folder at the root', async () => {
      mockRequest.post.mockResolvedValue({ id: '123' });
      await folderTools.createMailFolder({ displayName: 'New' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders');
      expect(mockRequest.post).toHaveBeenCalledWith({ displayName: 'New' });
    });

    it('creates a child folder under a parent', async () => {
      mockRequest.post.mockResolvedValue({ id: '123' });
      await folderTools.createMailFolder({ displayName: 'Child', parentFolderId: 'p1' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/p1/childFolders');
      expect(mockRequest.post).toHaveBeenCalledWith({ displayName: 'Child' });
    });
  });

  describe('updateMailFolder', () => {
    it('patches a folder', async () => {
      mockRequest.patch.mockResolvedValue({ id: '1' });
      await folderTools.updateMailFolder('1', { displayName: 'Renamed' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/1');
      expect(mockRequest.patch).toHaveBeenCalledWith({ displayName: 'Renamed' });
    });
  });

  describe('deleteMailFolder / permanentDeleteMailFolder', () => {
    it('deletes a folder', async () => {
      await folderTools.deleteMailFolder('1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/1');
      expect(mockRequest.delete).toHaveBeenCalled();
    });

    it('permanently deletes a folder', async () => {
      await folderTools.permanentDeleteMailFolder('1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/1/permanentDelete');
      expect(mockRequest.post).toHaveBeenCalledWith({});
    });
  });

  describe('copyMailFolder / moveMailFolder', () => {
    it('copies a folder', async () => {
      mockRequest.post.mockResolvedValue({ id: 'new' });
      await folderTools.copyMailFolder('1', 'dest');
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/1/copy');
      expect(mockRequest.post).toHaveBeenCalledWith({ destinationId: 'dest' });
    });

    it('moves a folder', async () => {
      mockRequest.post.mockResolvedValue({ id: '1' });
      await folderTools.moveMailFolder('1', 'dest');
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/1/move');
      expect(mockRequest.post).toHaveBeenCalledWith({ destinationId: 'dest' });
    });
  });

  describe('listChildFolders', () => {
    it('lists child folders of a parent', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: '1' }] });
      const result = await folderTools.listChildFolders('p1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/p1/childFolders');
      expect(result.value).toEqual([{ id: '1' }]);
    });
  });

  describe('getWellKnownFolder', () => {
    it.each(['inbox', 'sentitems', 'deleteditems', 'drafts', 'outbox', 'junkemail'] as const)(
      'gets the %s well-known folder',
      async (name) => {
        mockRequest.get.mockResolvedValue({ id: name });
        const result = await folderTools.getWellKnownFolder(name);
        expect(mockClient.api).toHaveBeenCalledWith(`/me/mailFolders/${name}`);
        expect(result).toEqual({ id: name });
      },
    );
  });

  describe('getFolderDelta', () => {
    it('uses the plain delta endpoint with no token', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await folderTools.getFolderDelta();
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/delta');
    });

    it('includes a delta token when provided', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await folderTools.getFolderDelta({ deltaToken: 'abc' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/delta?$deltatoken=abc');
    });

    it('includes a skip token when provided', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await folderTools.getFolderDelta({ skipToken: 'xyz' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/mailFolders/delta?$skiptoken=xyz');
    });
  });
});
