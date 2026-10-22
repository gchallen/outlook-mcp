import { PeopleTools } from '../tools/people-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('PeopleTools', () => {
  let peopleTools: PeopleTools;
  let mockClient: Client;
  let mockRequest: any;

  beforeEach(() => {
    mockRequest = {
      get: jest.fn().mockResolvedValue({ value: [{ displayName: 'Pat Teacher' }] }),
      query: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
    };
    mockClient = { api: jest.fn(() => mockRequest) } as any;
    peopleTools = new PeopleTools(mockClient);
  });

  it('searches /me/people with a quoted search term', async () => {
    const result = await peopleTools.searchPeople('Pat Teacher');
    expect(mockClient.api).toHaveBeenCalledWith('/me/people');
    expect(mockRequest.query).toHaveBeenCalledWith({ $search: '"Pat Teacher"' });
    expect(mockRequest.top).toHaveBeenCalledWith(10);
    expect(result).toEqual({ value: [{ displayName: 'Pat Teacher' }] });
  });

  it('clamps top to 25', async () => {
    await peopleTools.searchPeople('Pat', { top: 100 });
    expect(mockRequest.top).toHaveBeenCalledWith(25);
  });

  it('honors a smaller top', async () => {
    await peopleTools.searchPeople('Pat', { top: 3 });
    expect(mockRequest.top).toHaveBeenCalledWith(3);
  });
});
