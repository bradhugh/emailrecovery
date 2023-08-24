import {
  CopyError,
  FindItemResponse,
  IExchangeService,
} from "./IExchangeService";
import { FolderHierarchy } from "./FolderHierarchy";

/**
 * Copies items from one folder to another
 */
export class ItemCopier {
  private itemIdsToCopy: string[] = [];
  private badBatches: string[][] = [];
  private badItems: string[] = [];

  private discoveryComplete: boolean = false;
  private discoveryOffset = 0;

  private discoveryError: string = "";

  /**
   * Initializes a new instance of the ItemCopier class.
   * @param exchangeService the EWS service implementation
   * @param folderHierarchy the folder hierarchy
   * @param reportStatus the progress reporting implementation
   * @param sourceFolderId the source folder ID
   * @param targetFolderId the target folder ID
   * @param batchSize the initial batch size for discovery and copy
   */
  constructor(
    private exchangeService: IExchangeService,
    private folderHierarchy: FolderHierarchy,
    private reportStatus: (status: string) => void,
    private sourceFolderId: string,
    private targetFolderId: string,
    private batchSize: number
  ) {}

  /**
   * Starts the folder copy processing
   * 
   * @returns - True if all items have been copied
   */
  public async process(): Promise<boolean> {
    await this.startDiscoveryPass();
    if (
      (this.discoveryComplete || this.discoveryError != null) &&
      this.itemIdsToCopy.length === 0
    ) {
      return true;
    }

    if (this.itemIdsToCopy.length === 0) {
      return false;
    }

    await this.startCopyPass();

    this.reportStatus("Pass completed");

    return false;
  }

  private processFindItemResponse(response: FindItemResponse): void {

    // Add the discovered items
    for (var i = 0; i < response.messages.length; i++) {
      var message = response.messages[i];

      if (message.itemClass === "IPM.File.Document") {
        continue;
      }

      // Filter out items we don"t want
      if (!this.folderHierarchy.isFromIpmSubtree(message.lastActiveFolderId)) {
        continue;
      }

      if (
        this.folderHierarchy.isContactsSubfolder(message.lastActiveFolderId)
      ) {
        continue;
      }

      // Add the item to the list
      this.itemIdsToCopy.push(message.itemId);
    }

    this.discoveryOffset = response.indexedPagingOffset;

    if (response.includesLastItemInRange) {
      this.discoveryComplete = true;
    }
  }

  private async startDiscoveryPass(): Promise<void> {

    this.reportStatus("Discovering Items");

    try {
      const resp = await this.exchangeService.findItemAsync(this.sourceFolderId, this.batchSize, this.discoveryOffset);
      this.reportStatus("Processing Items");

      await this.processFindItemResponse(resp);
    } catch (error) {
      this.discoveryError = (error as Error)?.message ?? "Unknown error";
    }
  }

  private handleCopyError(attemptedIds: string[], error: CopyError) {
    if (attemptedIds.length > 1) {
      var half = attemptedIds.splice(0, Math.floor(attemptedIds.length / 2));
      this.badBatches.push(half);
      this.badBatches.push(attemptedIds);
    } else {
      this.badItems.push(attemptedIds[0]);
    }
  }

  private async startCopyPass(): Promise<boolean> {

    // Remove the item IDs from the front of the copy list
    var chunkItemIds = this.itemIdsToCopy.splice(0, this.batchSize);
    this.reportStatus(`Copying ${chunkItemIds.length} Items`);

    try {
      await this.exchangeService.copyItemsAsync(chunkItemIds, this.targetFolderId);
      this.reportStatus("Copy pass complete");

      return true;
    }
    catch (error)
    {
      this.handleCopyError(chunkItemIds, error);
      return false;
    }
  }
}
