import { Client } from '@microsoft/microsoft-graph-client';
import { Person } from '@microsoft/microsoft-graph-types';

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  PEOPLE: { default: 10, max: 25 },
} as const;

export class PeopleTools {
  constructor(private client: Client) {}

  /**
   * Search the organization directory and relevant people via /me/people.
   * Unlike /me/contacts (the personal Contacts folder), this covers directory
   * colleagues and frequently-contacted people, so it is the right tool when
   * search_contacts returns nothing.
   */
  async searchPeople(
    searchTerm: string,
    params?: {
      top?: number;
    },
  ): Promise<{ value: Person[] }> {
    const topLimit = Math.min(
      params?.top || PAGINATION_LIMITS.PEOPLE.default,
      PAGINATION_LIMITS.PEOPLE.max,
    );

    const response = await this.client
      .api('/me/people')
      .query({ $search: '"' + searchTerm + '"' })
      .top(topLimit)
      .get();

    return { value: response.value };
  }
}
