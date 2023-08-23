import { FindFolderResponse, IExchangeService } from "./IExchangeService";
import { IFolder } from "./IFolder";

/**
 * Encapsulates the folder hierarchy for a mailbox.
 */
export class FolderHierarchy {
  private static pageSize = 50;

  private shortFolderIdIndex: { [shortFolderId: string]: IFolder } = {};
  private distinguishedFolderIdIndex: {
    [distinguishedFolderId: string]: IFolder;
  } = {};

  /**
   * Initializes a new instance of the FolderHierarchy class.
   * @param ewsService the EWS service implementation
   */
  constructor(
    private ewsService: IExchangeService,
    public folders: IFolder[] = []
  ) {}

  /**
   * Initializes the folder hierarchy
   */
  public async initialize(): Promise<void> {
    const folders = await this.getAllChildFoldersRecursive("root");
    console.log(folders);

    for (var i = 0; i < folders.length; i++) {
      var folder = folders[i];

      // Add and index the folder
      this.folders.push(folder);
      this.shortFolderIdIndex[folder.shortFolderId] = folder;

      if (folder.distinguishedFolderId) {
        this.distinguishedFolderIdIndex[folder.distinguishedFolderId] = folder;
      }
    }
  }

  private async getAllChildFoldersRecursive(parentId: string): Promise<IFolder[]> {
    const folders: IFolder[] = [];
    let offset = 0;
    let resp: FindFolderResponse;
    do {
      resp = await this.ewsService.findFolderAsync(parentId, "Shallow", FolderHierarchy.pageSize, offset);
      folders.push(...resp.folders);
      offset += resp.folders.length;
    } while (!resp?.includesLastItemInRange)

    // Process child folders recursively
    for (const folder of folders) {
      if (folder.childFolderCount > 0) {
        const childFolders = await this.getAllChildFoldersRecursive(folder.folderId);
        folders.push(...childFolders);
      }
    }

    return folders;
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

    // For messages where we can't find the original folder, we have to that it's not a subfolder of contacts
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
