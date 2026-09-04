/**
 * Path REST tải nội dung file theo UniqueId.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/general-development/working-with-folders-and-files-with-rest
 */
export function fileDownloadPath(uniqueId: string): string {
  return `web/GetFileById('${uniqueId}')/$value`;
}
