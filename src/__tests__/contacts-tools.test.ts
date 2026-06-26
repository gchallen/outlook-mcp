import { ContactsTools } from '../tools/contacts-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('ContactsTools', () => {
  let contactsTools: ContactsTools;
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
      search: jest.fn().mockReturnThis(),
    };
    mockClient = { api: jest.fn(() => mockRequest) } as any;
    contactsTools = new ContactsTools(mockClient);
  });

  describe('listContacts', () => {
    it('lists contacts with default pagination', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: '1' }] });
      const result = await contactsTools.listContacts();
      expect(mockClient.api).toHaveBeenCalledWith('/me/contacts');
      expect(mockRequest.top).toHaveBeenCalledWith(25);
      expect(result.value).toEqual([{ id: '1' }]);
    });

    it('applies filter, select, ordering and clamps top to 100', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await contactsTools.listContacts({
        filter: "givenName eq 'Sam'",
        select: 'displayName',
        orderBy: 'displayName',
        top: 500,
      });
      expect(mockRequest.filter).toHaveBeenCalledWith("givenName eq 'Sam'");
      expect(mockRequest.select).toHaveBeenCalledWith('displayName');
      expect(mockRequest.orderby).toHaveBeenCalledWith('displayName');
      expect(mockRequest.top).toHaveBeenCalledWith(100);
    });
  });

  describe('searchContacts', () => {
    it('searches contacts', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: '1' }] });
      const result = await contactsTools.searchContacts('sam');
      expect(mockClient.api).toHaveBeenCalledWith('/me/contacts');
      expect(mockRequest.search).toHaveBeenCalledWith('sam');
      expect(result.value).toEqual([{ id: '1' }]);
    });
  });

  describe('getContact', () => {
    it('gets a single contact', async () => {
      mockRequest.get.mockResolvedValue({ id: '1' });
      const result = await contactsTools.getContact('1', { select: 'displayName' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/contacts/1');
      expect(mockRequest.select).toHaveBeenCalledWith('displayName');
      expect(result).toEqual({ id: '1' });
    });
  });

  describe('createContact', () => {
    it('creates a contact', async () => {
      const contact = { givenName: 'Sam', surname: 'Jones' };
      mockRequest.post.mockResolvedValue({ id: '1', ...contact });
      const result = await contactsTools.createContact(contact);
      expect(mockClient.api).toHaveBeenCalledWith('/me/contacts');
      expect(mockRequest.post).toHaveBeenCalledWith(contact);
      expect(result).toEqual({ id: '1', ...contact });
    });
  });

  describe('updateContact / deleteContact', () => {
    it('updates a contact', async () => {
      mockRequest.patch.mockResolvedValue({ id: '1' });
      await contactsTools.updateContact('1', { jobTitle: 'Lead' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/contacts/1');
      expect(mockRequest.patch).toHaveBeenCalledWith({ jobTitle: 'Lead' });
    });

    it('deletes a contact', async () => {
      await contactsTools.deleteContact('1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/contacts/1');
      expect(mockRequest.delete).toHaveBeenCalled();
    });
  });
});
