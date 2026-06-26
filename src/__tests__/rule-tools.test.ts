import { RuleTools } from '../tools/rule-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('RuleTools', () => {
  let ruleTools: RuleTools;
  let mockClient: Client;
  let mockRequest: any;

  const RULES = '/me/mailFolders/inbox/messageRules';

  beforeEach(() => {
    mockRequest = {
      get: jest.fn(),
      post: jest.fn().mockResolvedValue(undefined),
      patch: jest.fn().mockResolvedValue(undefined),
      delete: jest.fn().mockResolvedValue(undefined),
      top: jest.fn().mockReturnThis(),
      skip: jest.fn().mockReturnThis(),
    };
    mockClient = { api: jest.fn(() => mockRequest) } as any;
    ruleTools = new RuleTools(mockClient);
  });

  describe('listMessageRules', () => {
    it('lists inbox rules with default pagination', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: 'r1' }] });
      const result = await ruleTools.listMessageRules();
      expect(mockClient.api).toHaveBeenCalledWith(RULES);
      expect(mockRequest.top).toHaveBeenCalledWith(10);
      expect(result.value).toEqual([{ id: 'r1' }]);
    });

    it('clamps top to the rules max of 25', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await ruleTools.listMessageRules({ top: 100, skip: 3 });
      expect(mockRequest.top).toHaveBeenCalledWith(25);
      expect(mockRequest.skip).toHaveBeenCalledWith(3);
    });
  });

  describe('getMessageRule', () => {
    it('gets a single rule', async () => {
      mockRequest.get.mockResolvedValue({ id: 'r1' });
      const result = await ruleTools.getMessageRule('r1');
      expect(mockClient.api).toHaveBeenCalledWith(`${RULES}/r1`);
      expect(result).toEqual({ id: 'r1' });
    });
  });

  describe('createMessageRule', () => {
    it('creates a rule', async () => {
      const rule = { displayName: 'Move newsletters', sequence: 1 };
      mockRequest.post.mockResolvedValue({ id: 'r1', ...rule });
      const result = await ruleTools.createMessageRule(rule);
      expect(mockClient.api).toHaveBeenCalledWith(RULES);
      expect(mockRequest.post).toHaveBeenCalledWith(rule);
      expect(result).toEqual({ id: 'r1', ...rule });
    });
  });

  describe('updateMessageRule / deleteMessageRule', () => {
    it('updates a rule', async () => {
      mockRequest.patch.mockResolvedValue({ id: 'r1' });
      await ruleTools.updateMessageRule('r1', { isEnabled: false });
      expect(mockClient.api).toHaveBeenCalledWith(`${RULES}/r1`);
      expect(mockRequest.patch).toHaveBeenCalledWith({ isEnabled: false });
    });

    it('deletes a rule', async () => {
      await ruleTools.deleteMessageRule('r1');
      expect(mockClient.api).toHaveBeenCalledWith(`${RULES}/r1`);
      expect(mockRequest.delete).toHaveBeenCalled();
    });
  });
});
