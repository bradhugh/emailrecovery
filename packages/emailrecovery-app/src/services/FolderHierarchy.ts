import { IEwsService } from "./EwsService";
import { IFolder } from "./IFolder";

/**
 * Encapsulates the folder hierarchy for a mailbox.
 */
export class FolderHierarchy {
  private static pageSize = 50;
  private pagingOffset = 0;

  private shortFolderIdIndex: { [shortFolderId: string]: IFolder } = {};
  private distinguishedFolderIdIndex: {
    [distinguishedFolderId: string]: IFolder;
  } = {};

  /**
   * Initializes a new instance of the FolderHierarchy class.
   * @param ewsService the EWS service implementation
   */
  constructor(
    private ewsService: IEwsService,
    public folders: IFolder[] = []
  ) {}

  /**
   * Initializes the folder hierarchy
   */
  public async initialize(): Promise<void> {
    const result = await this.ewsService.findFolderAsync(
      "root",
      "Deep",
      FolderHierarchy.pageSize,
      this.pagingOffset
    );

    for (var i = 0; i < result.folders.length; i++) {
      var folder = result.folders[i];

      // Add and index the folder
      this.folders.push(folder);
      this.shortFolderIdIndex[folder.shortFolderId] = folder;

      if (folder.distinguishedFolderId) {
        this.distinguishedFolderIdIndex[folder.distinguishedFolderId] = folder;
      }

      // Capture the new paging offset
      this.pagingOffset = result.indexedPagingOffset;

      if (!result.includesLastItemInRange) {
        return this.initialize();
      }
    }
  }

  /**
   * Checks whether the given folder short id is located under the IPM subtree
   * @param shortFolderId the short folder ID
   */
  isFromIpmSubtree(shortFolderId: string) {
    if (shortFolderId == null || shortFolderId === "") {
      return true;
    }

    var folderInQuestion = this.shortFolderIdIndex[shortFolderId];

    // For messages where we can"t find the original folder, we have to assume true
    if (!folderInQuestion) {
      return true;
    }

    var ipmRoot = this.distinguishedFolderIdIndex["msgfolderroot"];

    // If the folder in question starts with the same path as the IPM root, it"s IPM
    if (folderInQuestion.folderPath.indexOf(ipmRoot.folderPath) === 0) {
      return true;
    }

    return false;
  }

  /**
   * Checks whether the given short folder ID is a subfolder of Contacts
   * @param shortFolderId the short folder ID
   */
  isContactsSubfolder(shortFolderId: string) {
    if (shortFolderId == null || shortFolderId === "") {
      return false;
    }

    var folderInQuestion = this.shortFolderIdIndex[shortFolderId];

    // For messages where we can"t find the original folder, we have to that it"s not a subfolder of contacts
    if (!folderInQuestion) {
      return false;
    }

    var contacts = this.distinguishedFolderIdIndex["contacts"];

    if (folderInQuestion.folderPath.indexOf(contacts.folderPath + "/") === 0) {
      return true;
    }

    return false;
  }
}
