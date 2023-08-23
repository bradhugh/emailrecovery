import { IFolder } from "./IFolder";

/**
 * A service for communicating with Exchange
 */
export interface IExchangeService {
  /**
   * Finds items in a folder using EWS
   * @param distinguishedFolderId the distinguished folder id
   * @param maxEntries the maximum number of messages to return
   * @param offset the start offset (used for paging)
   */
  findItemAsync(
    distinguishedFolderId: string,
    maxEntries: number,
    offset: number
  ): Promise<FindItemResponse>;

  /**
   * Creates a folder using EWS
   * @param distinguishedParentFolderId the parent folder id
   * @param displayName the folder display name
   * @param folderClass the optional folder class (default IPF.Note)
   */
  createFolderAsync(
    distinguishedParentFolderId: string,
    displayName: string,
    folderClass?: string
  ): Promise<CreateFolderResponse>;

  /**
   * Copies source items to a target folder
   * @param sourceItemIds the source item ids
   * @param targetFolderId the target folder id
   */
  copyItemsAsync(
    sourceItemIds: string[],
    targetFolderId: string
  ): Promise<CopyItemResponse>;

  /**
   * Find subfolders of a parent
   * @param rootFolderId the root folder id
   * @param traversal the traversal type: "Deep" or "Shallow"
   * @param maxEntries the maximum number of entries to return
   * @param pagingOffset the paging offset
   */
  findFolderAsync(
    rootFolderId: string,
    traversal: string,
    maxEntries: number,
    pagingOffset: number
  ): Promise<FindFolderResponse>;
}

export class EwsResponse {
  responseClass: string = "";
  responseCode: string = "";
}

export class FindResponse extends EwsResponse {
  indexedPagingOffset: number = 0;
  //totalItemsInView: number = 0;
  includesLastItemInRange: boolean = true;
}

export class EmailMessage {
  messageType: string = "";
  itemId: string = "";
  subject: string = "";
  lastActiveFolderId: string = "";
  itemClass: string = "";
}

export class FindItemResponse extends FindResponse {
  messages: EmailMessage[] = [];
}

export class CreateFolderResponse extends EwsResponse {
  folderId: string = "";
}

export class CopyItemResponse extends EwsResponse {
  newItemIds: string[] = [];
}

export class FindFolderResponse extends FindResponse {
  folders: IFolder[] = [];
}

export class DiscoveryError extends Error {
  constructor(message: string) {
    super(message);

    // When extending the built-in Error type, you have to fix up the prototype chain
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

export class CopyError extends Error {
  constructor(message: string) {
    super(message);

    // When extending the built-in Error type, you have to fix up the prototype chain
    Object.setPrototypeOf(this, new.target.prototype);
  }
}
