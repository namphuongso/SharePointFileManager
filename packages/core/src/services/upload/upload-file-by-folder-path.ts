import { encodeODataQuotedValue, encodeServerRelativePathArg } from "../../utils/odata-path-arg";

/**
 * POST upload vào folder theo ServerRelativeUrl — endpoint Microsoft document.
 * UniqueId → resolve ServerRelativeUrl trước (GetFolderById), rồi Files/add.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
 */
export function uploadFileByFolderPath(
  folderServerRelativeUrl: string,
  fileName: string,
  overwrite: boolean,
): string {
  const folder = encodeServerRelativePathArg(folderServerRelativeUrl);
  const leaf = encodeODataQuotedValue(fileName);
  return `web/GetFolderByServerRelativePath(decodedUrl='${folder}')/Files/add(url='${leaf}',overwrite=${overwrite})`;
}
