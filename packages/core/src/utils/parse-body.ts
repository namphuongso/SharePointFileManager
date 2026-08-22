/** Đọc JSON lỗi khi HTTP không OK; body hỏng thì trả undefined. */
export async function readErrorBody(response: Response): Promise<unknown> {
  try {
    return await response.json();
  } catch {
    return undefined;
  }
}

/**
 * Parse body thành công: 204 / rỗng → undefined; JSON thì object; còn lại giữ text.
 */
export async function parseSuccessBody<T>(response: Response): Promise<T> {
  if (
    response.status === 204 ||
    response.headers.get("Content-Length") === "0"
  ) {
    return undefined as T;
  }
  const text = await response.text();
  if (!text) return undefined as T;
  try {
    return JSON.parse(text) as T;
  } catch {
    return text as unknown as T;
  }
}
