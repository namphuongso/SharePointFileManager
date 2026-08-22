/** Lấy chuỗi message từ Error / SharePointError; không có thì dùng fallback. */
export function getErrorMessage(error: unknown, fallback = "Unknown error"): string {
  if (!error) return fallback;
  if (typeof error === "object" && error !== null && "message" in error) {
    const message = (error as { message?: unknown }).message;
    if (typeof message === "string" && message.trim()) return message;
  }
  return fallback;
}
