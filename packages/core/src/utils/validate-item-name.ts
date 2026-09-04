import { SharePointError, SharePointErrorCode } from "../errors/sharepoint-error";

/** Ký tự cấm trong tên file/folder OneDrive & SharePoint. */
const INVALID_CHARS = /["*:<>?/\\|]/;

/** Tên thiết bị Windows — SharePoint từ chối. */
const RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\.|$)/i;

/**
 * Kiểm tra tên leaf trước khi POST tạo folder / upload.
 * @see https://support.microsoft.com/office/invalid-file-names-and-file-types-in-onedrive-and-sharepoint
 */
export function assertValidItemName(name: string, kind: "file" | "folder"): void {
  const trimmed = name.trim();
  if (!trimmed) {
    throw new SharePointError({
      code: SharePointErrorCode.Unsupported,
      message: `${kind} name is required`,
    });
  }
  if (trimmed !== name) {
    throw new SharePointError({
      code: SharePointErrorCode.Unsupported,
      message: `${kind} name cannot start or end with whitespace`,
    });
  }
  if (trimmed === "." || trimmed === "..") {
    throw new SharePointError({
      code: SharePointErrorCode.Unsupported,
      message: `Invalid ${kind} name`,
    });
  }
  if (INVALID_CHARS.test(trimmed) || trimmed.includes("\0")) {
    throw new SharePointError({
      code: SharePointErrorCode.Unsupported,
      message: `${kind} name contains invalid characters`,
    });
  }
  if (/[. ]$/.test(trimmed)) {
    throw new SharePointError({
      code: SharePointErrorCode.Unsupported,
      message: `${kind} name cannot end with a period or space`,
    });
  }
  if (RESERVED.test(trimmed)) {
    throw new SharePointError({
      code: SharePointErrorCode.Unsupported,
      message: `${kind} name is reserved`,
    });
  }
  if (trimmed.length > 255) {
    throw new SharePointError({
      code: SharePointErrorCode.Unsupported,
      message: `${kind} name is too long`,
    });
  }
}
