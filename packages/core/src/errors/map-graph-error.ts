import { SharePointError, SharePointErrorCode } from "./sharepoint-error";

interface GraphErrorBody {
  error?: {
    code?: string;
    message?: string;
    innerError?: { code?: string; message?: string };
  };
}

const GRAPH_CODE_MAP: Record<string, SharePointErrorCode> = {
  unauthenticated: SharePointErrorCode.Unauthorized,
  accessDenied: SharePointErrorCode.Forbidden,
  forbidden: SharePointErrorCode.Forbidden,
  itemNotFound: SharePointErrorCode.NotFound,
  notFound: SharePointErrorCode.NotFound,
  nameAlreadyExists: SharePointErrorCode.Conflict,
  resourceModified: SharePointErrorCode.Conflict,
  activityLimitReached: SharePointErrorCode.Throttled,
  tooManyRetries: SharePointErrorCode.Throttled,
  maxFileSizeExceeded: SharePointErrorCode.TooLarge,
  invalidRequest: SharePointErrorCode.Unknown,
  notSupported: SharePointErrorCode.Unsupported,
  notAllowed: SharePointErrorCode.Forbidden,
};

export function mapStatusToCode(status: number): SharePointErrorCode {
  if (status === 401) return SharePointErrorCode.Unauthorized;
  if (status === 403) return SharePointErrorCode.Forbidden;
  if (status === 404) return SharePointErrorCode.NotFound;
  if (status === 409 || status === 412) return SharePointErrorCode.Conflict;
  if (status === 413) return SharePointErrorCode.TooLarge;
  if (status === 429) return SharePointErrorCode.Throttled;
  if (status === 501) return SharePointErrorCode.Unsupported;
  return SharePointErrorCode.Unknown;
}

export function parseRetryAfterMs(header: string | null): number | undefined {
  if (!header) return undefined;
  const seconds = Number(header);
  if (!Number.isNaN(seconds)) return seconds * 1000;
  const date = Date.parse(header);
  if (!Number.isNaN(date)) return Math.max(0, date - Date.now());
  return undefined;
}

export function mapGraphError(input: {
  status: number;
  body?: unknown;
  retryAfter?: string | null;
  fallbackMessage?: string;
}): SharePointError {
  const graph = input.body as GraphErrorBody | undefined;
  const graphCode = graph?.error?.code ?? graph?.error?.innerError?.code;
  const message =
    graph?.error?.message ??
    input.fallbackMessage ??
    `Microsoft Graph request failed with status ${input.status}`;

  const code =
    (graphCode && GRAPH_CODE_MAP[graphCode]) || mapStatusToCode(input.status);

  return new SharePointError({
    code,
    message,
    status: input.status,
    graphCode,
    retryAfterMs: parseRetryAfterMs(input.retryAfter ?? null),
  });
}
