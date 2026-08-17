import type { PermissionKind, SharePointPermission, UserInfo } from "../types/models";

interface GraphIdentity {
  id?: string;
  displayName?: string;
  email?: string;
}

interface GraphPermission {
  id?: string;
  roles?: string[];
  expirationDateTime?: string;
  inheritedFrom?: { id?: string; path?: string };
  grantedTo?: { user?: GraphIdentity };
  grantedToV2?: {
    user?: GraphIdentity;
    group?: GraphIdentity;
    siteUser?: GraphIdentity;
    siteGroup?: GraphIdentity;
  };
  grantedToIdentitiesV2?: Array<{
    user?: GraphIdentity;
    group?: GraphIdentity;
    siteUser?: GraphIdentity;
    siteGroup?: GraphIdentity;
  }>;
  link?: {
    type?: string;
    scope?: string;
    webUrl?: string;
    preventsDownload?: boolean;
  };
}

function userFrom(identity?: GraphIdentity): UserInfo | undefined {
  if (!identity) return undefined;
  return {
    id: identity.id,
    displayName: identity.displayName,
    email: identity.email,
  };
}

export function mapPermission(permission: GraphPermission): SharePointPermission {
  if (!permission.id) {
    throw new Error("Graph permission is missing id");
  }

  const inherited = Boolean(permission.inheritedFrom);
  const identities = permission.grantedToIdentitiesV2 ?? [];
  const granted =
    permission.grantedToV2?.user ??
    permission.grantedTo?.user ??
    identities[0]?.user ??
    permission.grantedToV2?.siteUser ??
    identities[0]?.siteUser;
  const group =
    permission.grantedToV2?.group ??
    permission.grantedToV2?.siteGroup ??
    identities[0]?.group ??
    identities[0]?.siteGroup;

  let kind: PermissionKind = "unknown";
  if (permission.link) kind = "link";
  else if (group) kind = permission.grantedToV2?.siteGroup || identities[0]?.siteGroup ? "siteGroup" : "group";
  else if (granted) kind = "user";

  return {
    id: permission.id,
    roles: permission.roles ?? [],
    kind,
    inherited,
    inheritedFromId: permission.inheritedFrom?.id,
    grantedTo: userFrom(granted),
    grantedToGroup: group
      ? { id: group.id, displayName: group.displayName }
      : undefined,
    link: permission.link
      ? {
          type: permission.link.type,
          scope: permission.link.scope,
          webUrl: permission.link.webUrl,
          preventsDownload: permission.link.preventsDownload,
        }
      : undefined,
    expirationDateTime: permission.expirationDateTime,
    canRemove: !inherited,
  };
}
