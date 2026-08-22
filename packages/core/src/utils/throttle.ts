import { parseRetryAfterMs } from "./parse-retry-after";

export const sleep = (ms: number) =>
  new Promise((resolve) => setTimeout(resolve, ms));

/**
 * Thời gian chờ sau HTTP 429. Ưu tiên header Retry-After;
 * không parse được thì tăng dần theo số lần thử (1s, 2s, 3s).
 */
export function throttleWaitMs(
  retryAfterHeader: string | null,
  attempt: number,
): number {
  return parseRetryAfterMs(retryAfterHeader) ?? 1000 * attempt;
}
