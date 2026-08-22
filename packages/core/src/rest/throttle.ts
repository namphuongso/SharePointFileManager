export const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * Thời gian chờ sau HTTP 429. Ưu tiên header Retry-After (giây);
 * không có thì tăng dần theo số lần thử (1s, 2s, 3s).
 */
export function throttleWaitMs(retryAfterHeader: string | null, attempt: number): number {
  if (retryAfterHeader) return Number(retryAfterHeader) * 1000 || 2000;
  return 1000 * attempt;
}
