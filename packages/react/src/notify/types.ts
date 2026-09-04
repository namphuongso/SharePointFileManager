/** API thông báo dùng trong toàn package: success / info / error. */
export interface NotifyApi {
  /** Toast thành công — auto-dismiss 3s. */
  success: (title: string, subtitle?: string) => void;
  /** Toast trạng thái (đang tải, v.v.) — auto-dismiss 3s. */
  info: (title: string, subtitle?: string) => string;
  /** Toast lỗi — auto-dismiss 6s, intent error, đọc assertive. */
  error: (title: string, subtitle?: string) => void;
  /** Cập nhật nội dung toast đã dispatch (giữ nguyên id để tránh chồng). Trả về true nếu tìm thấy. */
  update: (id: string, options: { title: string; subtitle?: string; intent: "info" | "success" | "error" }) => void;
  /** Đóng toast theo id (no-op nếu đã tự đóng). */
  dismiss: (id: string) => void;
}
