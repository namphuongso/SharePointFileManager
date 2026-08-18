import type { DirectoryPerson, DirectoryPersonKind, InviteRecipient } from "../types/models";

export interface GraphPerson {
  id?: string;
  displayName?: string;
  userPrincipalName?: string;
  scoredEmailAddresses?: Array<{ address?: string }>;
  personType?: { class?: string; subclass?: string };
}

export interface GraphDirectoryUser {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

export interface GraphDirectoryGroup {
  id?: string;
  displayName?: string;
  mail?: string;
}

function text(value?: string): string {
  return typeof value === "string" ? value.trim() : "";
}

export function isLikelyEmail(value: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim());
}

export function mapGraphPerson(raw: GraphPerson): DirectoryPerson | undefined {
  const email = text(raw.scoredEmailAddresses?.[0]?.address) || text(raw.userPrincipalName);
  const displayName = text(raw.displayName) || email;
  if (!displayName) return undefined;

  const className = text(raw.personType?.class).toLowerCase();
  const kind: DirectoryPersonKind = className === "group" ? "group" : "user";

  return {
    key: `person:${text(raw.id || email || displayName).toLowerCase()}`,
    displayName,
    email: email || undefined,
    kind,
  };
}

export function mapGraphUser(raw: GraphDirectoryUser): DirectoryPerson | undefined {
  const email = text(raw.mail) || text(raw.userPrincipalName);
  const displayName = text(raw.displayName) || email;
  const objectId = text(raw.id);
  if (!objectId && !email) return undefined;

  return {
    key: `user:${(objectId || email).toLowerCase()}`,
    displayName,
    email: email || undefined,
    objectId: objectId || undefined,
    kind: "user",
  };
}

export function mapGraphGroup(raw: GraphDirectoryGroup): DirectoryPerson | undefined {
  const email = text(raw.mail);
  const displayName = text(raw.displayName) || email;
  const objectId = text(raw.id);
  if (!objectId && !displayName) return undefined;

  return {
    key: `group:${(objectId || displayName).toLowerCase()}`,
    displayName,
    email: email || undefined,
    objectId: objectId || undefined,
    kind: "group",
  };
}

export function mapTypedEmail(value: string): DirectoryPerson | undefined {
  const email = text(value);
  if (!isLikelyEmail(email)) return undefined;
  return {
    key: `email:${email.toLowerCase()}`,
    displayName: email,
    email,
    kind: "email",
  };
}

export function toInviteRecipient(person: DirectoryPerson): InviteRecipient | undefined {
  if (person.objectId) return { objectId: person.objectId };
  if (person.email) return { email: person.email };
  return undefined;
}
