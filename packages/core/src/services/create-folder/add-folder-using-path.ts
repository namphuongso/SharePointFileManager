import { encodeServerRelativePathArg } from "../../utils/odata-path-arg";

/**
 * POST tạo folder theo ServerRelativeUrl đầy đủ (ResourcePath — hỗ trợ % / #).
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/supporting-and-in-file-and-folder-with-the-resourcepath-api
 */
export function addFolderUsingPath(serverRelativeUrl: string): string {
  return `web/Folders/AddUsingPath(decodedurl='${encodeServerRelativePathArg(serverRelativeUrl)}')`;
}
