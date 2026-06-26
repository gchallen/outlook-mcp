import { Client } from '@microsoft/microsoft-graph-client';
import { Event, Calendar } from '@microsoft/microsoft-graph-types';

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  EVENTS: { default: 25, max: 100 },
  CALENDARS: { default: 20, max: 50 },
} as const;

export class CalendarTools {
  constructor(private client: Client) {}

  async listCalendars(params?: {
    select?: string;
    top?: number;
    skip?: number;
  }): Promise<{
    value: Calendar[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const request = this.client.api('/me/calendars');

    if (params?.select) request.select(params.select);
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.CALENDARS.default,
      PAGINATION_LIMITS.CALENDARS.max,
    );
    request.top(topLimit);
    if (params?.skip) request.skip(params.skip);

    const response = await request.get();
    return {
      value: response.value,
      '@odata.nextLink': response['@odata.nextLink'],
      '@odata.count': response['@odata.count'],
    };
  }

  async listEvents(params?: {
    calendarId?: string;
    filter?: string;
    select?: string;
    top?: number;
    skip?: number;
    orderBy?: string;
  }): Promise<{
    value: Event[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const endpoint = params?.calendarId
      ? `/me/calendars/${params.calendarId}/events`
      : '/me/events';

    const request = this.client.api(endpoint);

    if (params?.filter) request.filter(params.filter);
    if (params?.select) request.select(params.select);
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.EVENTS.default,
      PAGINATION_LIMITS.EVENTS.max,
    );
    request.top(topLimit);
    if (params?.skip) request.skip(params.skip);
    request.orderby(params?.orderBy || 'start/dateTime');

    const response = await request.get();
    return {
      value: response.value,
      '@odata.nextLink': response['@odata.nextLink'],
      '@odata.count': response['@odata.count'],
    };
  }

  /**
   * Get events occurring within a date/time range, expanding recurring series
   * into individual occurrences. `startDateTime` and `endDateTime` are ISO 8601.
   */
  async getCalendarView(params: {
    startDateTime: string;
    endDateTime: string;
    calendarId?: string;
    select?: string;
    top?: number;
    skip?: number;
  }): Promise<{
    value: Event[];
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
  }> {
    const endpoint = params.calendarId
      ? `/me/calendars/${params.calendarId}/calendarView`
      : '/me/calendarView';

    const request = this.client
      .api(endpoint)
      .query({
        startDateTime: params.startDateTime,
        endDateTime: params.endDateTime,
      });

    if (params.select) request.select(params.select);
    const topLimit = Math.min(
      params.top || PAGINATION_LIMITS.EVENTS.default,
      PAGINATION_LIMITS.EVENTS.max,
    );
    request.top(topLimit);
    if (params.skip) request.skip(params.skip);
    request.orderby('start/dateTime');

    const response = await request.get();
    return {
      value: response.value,
      '@odata.nextLink': response['@odata.nextLink'],
      '@odata.count': response['@odata.count'],
    };
  }

  async getEvent(eventId: string, params?: { select?: string }): Promise<Event> {
    const request = this.client.api(`/me/events/${eventId}`);
    if (params?.select) request.select(params.select);
    return await request.get();
  }

  async createEvent(event: Partial<Event>, calendarId?: string): Promise<Event> {
    const endpoint = calendarId
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';
    return await this.client.api(endpoint).post(event);
  }

  async updateEvent(eventId: string, updates: Partial<Event>): Promise<Event> {
    return await this.client.api(`/me/events/${eventId}`).patch(updates);
  }

  async deleteEvent(eventId: string): Promise<void> {
    await this.client.api(`/me/events/${eventId}`).delete();
  }

  /** Respond to a meeting invite. `response` is accept | decline | tentativelyAccept. */
  async respondToEvent(
    eventId: string,
    response: 'accept' | 'decline' | 'tentativelyAccept',
    comment?: string,
    sendResponse = true,
  ): Promise<void> {
    await this.client.api(`/me/events/${eventId}/${response}`).post({
      comment,
      sendResponse,
    });
  }
}
