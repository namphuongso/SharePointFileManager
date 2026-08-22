import type { RestODataCollection } from "../types/rest";

/** Tách `value` và `@odata.nextLink` từ JSON nometadata. */
export function parseODataCollection<T>(body: RestODataCollection<T> | undefined): {
  value: T[];
  nextLink?: string;
} {
  const nextLink = body?.["@odata.nextLink"]?.trim();
  return {
    value: body?.value ?? [],
    nextLink: nextLink || undefined,
  };
}
