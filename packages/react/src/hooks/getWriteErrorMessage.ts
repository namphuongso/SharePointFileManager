import {
  isSharePointError,
  SharePointErrorCode,
} from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../types/messages";
import { getErrorMessage } from "./getErrorMessage";

/** Map SharePointError → chuỗi UI cho tạo folder / upload. */
export function getWriteErrorMessage(
  error: unknown,
  messages: Messages,
  fallback: string,
): string {
  if (isSharePointError(error)) {
    if (error.code === SharePointErrorCode.Forbidden) return messages.noAddPermission;
    if (error.code === SharePointErrorCode.Conflict) return messages.nameConflict;
    if (error.code === SharePointErrorCode.TooLarge) return messages.fileTooLarge;
    if (error.code === SharePointErrorCode.Unsupported) return messages.invalidName;
  }
  return getErrorMessage(error, fallback);
}
