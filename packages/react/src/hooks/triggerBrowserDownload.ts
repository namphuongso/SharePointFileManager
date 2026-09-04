/**
 * Kích hoạt Save As từ Blob (OAuth — không gắn URL SharePoint trực tiếp lên <a href>).
 * Thu hồi object URL sau một nhịp để trình duyệt kịp bắt đầu tải.
 */
export function triggerBrowserDownload(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.rel = "noopener";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
}
