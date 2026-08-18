import type { GraphClient } from "../graph/client";
import type { GraphCollection } from "../mappers/item";
import {
  mapGraphGroup,
  mapGraphPerson,
  mapGraphUser,
  mapTypedEmail,
  type GraphDirectoryGroup,
  type GraphDirectoryUser,
  type GraphPerson,
} from "../mappers/person";
import type { DirectoryPerson } from "../types/models";

function sanitizeSearch(query: string): string {
  return query.replace(/["\\]/g, " ").trim();
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function escapeODataSearch(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function samePerson(left: DirectoryPerson, right: DirectoryPerson): boolean {
  if (left.objectId && right.objectId && left.objectId.toLowerCase() === right.objectId.toLowerCase()) return true;
  if (left.email && right.email && left.email.toLowerCase() === right.email.toLowerCase()) return true;
  return left.key.toLowerCase() === right.key.toLowerCase();
}

function addUnique(target: DirectoryPerson[], person: DirectoryPerson | undefined): void {
  if (!person) return;
  const existingIndex = target.findIndex((item) => samePerson(item, person));
  if (existingIndex >= 0) {
    const existing = target[existingIndex];
    if (existing && !existing.objectId && person.objectId) {
      target[existingIndex] = person;
    }
    return;
  }
  target.push(person);
}

export class PeopleService {
  constructor(private readonly graph: GraphClient) {}

  async search(query: string, signal?: AbortSignal): Promise<DirectoryPerson[]> {
    const term = sanitizeSearch(query);
    if (term.length < 1) return this.suggest(signal);

    const results: DirectoryPerson[] = [];
    const sources = await Promise.allSettled([
      this.searchRelevantPeople(term, signal),
      this.searchUsersAdvanced(term, signal),
      this.searchUsersFilter(term, signal),
      this.searchGroups(term, signal),
    ]);

    for (const source of sources) {
      if (source.status !== "fulfilled") continue;
      source.value.forEach((person) => addUnique(results, person));
    }

    addUnique(results, mapTypedEmail(term));
    return results.slice(0, 15);
  }

  /** Recent / relevant people shown when the picker is focused with no query (SharePoint default). */
  async suggest(signal?: AbortSignal): Promise<DirectoryPerson[]> {
    try {
      const result = await this.graph.get<GraphCollection<GraphPerson>>("/me/people", {
        query: {
          $select: "id,displayName,scoredEmailAddresses,userPrincipalName,personType",
          $top: 15,
        },
        signal,
      });
      return (result.value ?? [])
        .map(mapGraphPerson)
        .filter((person): person is DirectoryPerson => Boolean(person));
    } catch {
      return [];
    }
  }

  private async searchRelevantPeople(term: string, signal?: AbortSignal): Promise<DirectoryPerson[]> {
    const result = await this.graph.get<GraphCollection<GraphPerson>>("/me/people", {
      query: {
        $search: `"${escapeODataSearch(term)}"`,
        $select: "id,displayName,scoredEmailAddresses,userPrincipalName,personType",
        $top: 15,
      },
      signal,
    });
    return (result.value ?? []).map(mapGraphPerson).filter((person): person is DirectoryPerson => Boolean(person));
  }

  private async searchUsersAdvanced(term: string, signal?: AbortSignal): Promise<DirectoryPerson[]> {
    const escaped = escapeODataSearch(term);
    const result = await this.graph.get<GraphCollection<GraphDirectoryUser>>("/users", {
      query: {
        $count: "true",
        $search: `"displayName:${escaped}" OR "mail:${escaped}" OR "userPrincipalName:${escaped}"`,
        $select: "id,displayName,mail,userPrincipalName",
        $top: 10,
      },
      headers: { ConsistencyLevel: "eventual" },
      signal,
    });
    return (result.value ?? []).map(mapGraphUser).filter((person): person is DirectoryPerson => Boolean(person));
  }

  /** Fallback when $search is blocked or not licensed in the tenant. */
  private async searchUsersFilter(term: string, signal?: AbortSignal): Promise<DirectoryPerson[]> {
    const e = escapeODataString(term);
    const filter = [
      `startswith(displayName,'${e}')`,
      `startswith(mail,'${e}')`,
      `startswith(userPrincipalName,'${e}')`,
      `startswith(givenName,'${e}')`,
      `startswith(surname,'${e}')`,
    ].join(" or ");

    const result = await this.graph.get<GraphCollection<GraphDirectoryUser>>("/users", {
      query: {
        $filter: filter,
        $select: "id,displayName,mail,userPrincipalName",
        $top: 10,
      },
      signal,
    });
    return (result.value ?? []).map(mapGraphUser).filter((person): person is DirectoryPerson => Boolean(person));
  }

  private async searchGroups(term: string, signal?: AbortSignal): Promise<DirectoryPerson[]> {
    const result = await this.graph.get<GraphCollection<GraphDirectoryGroup>>("/groups", {
      query: {
        $filter: `startswith(displayName,'${escapeODataString(term)}')`,
        $select: "id,displayName,mail",
        $top: 8,
      },
      signal,
    });
    return (result.value ?? []).map(mapGraphGroup).filter((person): person is DirectoryPerson => Boolean(person));
  }
}
