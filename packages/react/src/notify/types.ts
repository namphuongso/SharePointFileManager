/** API thông báo dùng trong toàn package: success / info / progress / error. */
export interface NotifyApi {
  /** Toast thành công — auto-dismiss 3s. */
  success: (title: string, subtitle?: string) => void;
  /** Toast trạng thái ngắn — auto-dismiss 3s. */
  info: (title: string, subtitle?: string) => string;
  /**
   * Toast tiến trình dài (upload / tải folder…) — không tự tắt.
   * Caller `update` success/error hoặc `dismiss` khi xong.
   */
  progress: (title: string, subtitle?: string) => string;
  /** Toast lỗi — auto-dismiss 6s, intent error, đọc assertive. */
  error: (title: string, subtitle?: string) => void;
  /**
   * Cập nhật toast đã dispatch (cùng id).
   * intent info → vẫn sticky; success/error → auto-dismiss.
   */
  update: (
    id: string,
    options: { title: string; subtitle?: string; intent: "info" | "success" | "error" },
  ) => void;
  /** Đóng toast theo id (no-op nếu đã tự đóng). */
  dismiss: (id: string) => void;
}
