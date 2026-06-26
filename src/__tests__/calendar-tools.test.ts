import { CalendarTools } from '../tools/calendar-tools';
import { Client } from '@microsoft/microsoft-graph-client';

jest.mock('@microsoft/microsoft-graph-client');

describe('CalendarTools', () => {
  let calendarTools: CalendarTools;
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
      query: jest.fn().mockReturnThis(),
    };
    mockClient = { api: jest.fn(() => mockRequest) } as any;
    calendarTools = new CalendarTools(mockClient);
  });

  describe('listCalendars', () => {
    it('lists calendars and clamps top to the max of 50', async () => {
      mockRequest.get.mockResolvedValue({ value: [{ id: 'c1' }] });
      const result = await calendarTools.listCalendars({ top: 999, select: 'name' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/calendars');
      expect(mockRequest.select).toHaveBeenCalledWith('name');
      expect(mockRequest.top).toHaveBeenCalledWith(50);
      expect(result.value).toEqual([{ id: 'c1' }]);
    });
  });

  describe('listEvents', () => {
    it('lists from /me/events by default and orders by start time', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await calendarTools.listEvents();
      expect(mockClient.api).toHaveBeenCalledWith('/me/events');
      expect(mockRequest.orderby).toHaveBeenCalledWith('start/dateTime');
    });

    it('lists from a specific calendar with a custom order', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await calendarTools.listEvents({ calendarId: 'c1', orderBy: 'subject', filter: 'x' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/calendars/c1/events');
      expect(mockRequest.filter).toHaveBeenCalledWith('x');
      expect(mockRequest.orderby).toHaveBeenCalledWith('subject');
    });
  });

  describe('getCalendarView', () => {
    it('queries the date range on the default calendar view', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await calendarTools.getCalendarView({
        startDateTime: '2026-06-26T00:00:00',
        endDateTime: '2026-07-03T00:00:00',
      });
      expect(mockClient.api).toHaveBeenCalledWith('/me/calendarView');
      expect(mockRequest.query).toHaveBeenCalledWith({
        startDateTime: '2026-06-26T00:00:00',
        endDateTime: '2026-07-03T00:00:00',
      });
      expect(mockRequest.orderby).toHaveBeenCalledWith('start/dateTime');
    });

    it('uses a specific calendar when given an id', async () => {
      mockRequest.get.mockResolvedValue({ value: [] });
      await calendarTools.getCalendarView({
        startDateTime: 'a',
        endDateTime: 'b',
        calendarId: 'c1',
      });
      expect(mockClient.api).toHaveBeenCalledWith('/me/calendars/c1/calendarView');
    });
  });

  describe('getEvent', () => {
    it('gets a single event', async () => {
      mockRequest.get.mockResolvedValue({ id: 'e1' });
      const result = await calendarTools.getEvent('e1', { select: 'subject' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/events/e1');
      expect(mockRequest.select).toHaveBeenCalledWith('subject');
      expect(result).toEqual({ id: 'e1' });
    });
  });

  describe('createEvent', () => {
    it('creates an event on the default calendar', async () => {
      const event = { subject: 'Meeting' };
      mockRequest.post.mockResolvedValue({ id: 'e1' });
      const result = await calendarTools.createEvent(event);
      expect(mockClient.api).toHaveBeenCalledWith('/me/events');
      expect(mockRequest.post).toHaveBeenCalledWith(event);
      expect(result).toEqual({ id: 'e1' });
    });

    it('creates an event on a specific calendar', async () => {
      mockRequest.post.mockResolvedValue({ id: 'e1' });
      await calendarTools.createEvent({ subject: 'X' }, 'c1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/calendars/c1/events');
    });
  });

  describe('updateEvent / deleteEvent', () => {
    it('updates an event', async () => {
      mockRequest.patch.mockResolvedValue({ id: 'e1' });
      await calendarTools.updateEvent('e1', { subject: 'New' });
      expect(mockClient.api).toHaveBeenCalledWith('/me/events/e1');
      expect(mockRequest.patch).toHaveBeenCalledWith({ subject: 'New' });
    });

    it('deletes an event', async () => {
      await calendarTools.deleteEvent('e1');
      expect(mockClient.api).toHaveBeenCalledWith('/me/events/e1');
      expect(mockRequest.delete).toHaveBeenCalled();
    });
  });

  describe('respondToEvent', () => {
    it('accepts with a comment and default sendResponse', async () => {
      await calendarTools.respondToEvent('e1', 'accept', 'See you there');
      expect(mockClient.api).toHaveBeenCalledWith('/me/events/e1/accept');
      expect(mockRequest.post).toHaveBeenCalledWith({
        comment: 'See you there',
        sendResponse: true,
      });
    });

    it('declines without sending a response', async () => {
      await calendarTools.respondToEvent('e1', 'decline', undefined, false);
      expect(mockClient.api).toHaveBeenCalledWith('/me/events/e1/decline');
      expect(mockRequest.post).toHaveBeenCalledWith({
        comment: undefined,
        sendResponse: false,
      });
    });

    it('routes a tentative response to the tentativelyAccept action', async () => {
      await calendarTools.respondToEvent('e1', 'tentativelyAccept');
      expect(mockClient.api).toHaveBeenCalledWith('/me/events/e1/tentativelyAccept');
    });
  });
});
