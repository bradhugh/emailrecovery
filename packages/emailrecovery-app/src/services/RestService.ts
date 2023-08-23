import {
  CopyItemResponse,
  CreateFolderResponse,
  FindFolderResponse,
  FindItemResponse,
  IExchangeService,
} from "./IExchangeService";
import jwt_decode, { JwtPayload } from "jwt-decode";
import { default as $ } from "jquery";
import { Utils } from "../Utils";

export class RestService implements IExchangeService {
  private static PropNames = {
    EntryId: "Binary 0xfff",
    FolderPathFullName: "String 0x6874",
  };

  private static getBaseUrl(url: string): string {
    const parts = url.split("/");

    return parts[0] + "//" + parts[2];
  }

  private static getRestUrl(accessToken: string): string {
    // Shim function to workaround
    // mailbox.restUrl == null case
    if (Office.context.mailbox.restUrl) {
      return RestService.getBaseUrl(Office.context.mailbox.restUrl);
    }

    // parse the token
    const jwt = jwt_decode<JwtPayload>(accessToken);

    // 'aud' parameter from token can be in a couple of
    // different formats.
    const aud = Array.isArray(jwt.aud) ? jwt.aud[0] : jwt.aud;

    if (aud) {
      // Format 1: It's just the URL
      if (aud.match(/https:\/\/([^@]*)/)) {
        return aud;
      }

      // Format 2: GUID/hostname@GUID
      const match = aud.match(/\/([^@]*)@/);
      if (match && match[1]) {
        return "https://" + match[1];
      }
    }

    // Couldn't find what we expected, default to
    // outlook.office.com
    return "https://outlook.office.com";
  }

  public async findItemAsync(
    distinguishedFolderId: string,
    maxEntries: number,
    offset: number
  ): Promise<FindItemResponse> {
    throw new Error("Method not implemented.");
  }

  public async createFolderAsync(
    distinguishedParentFolderId: string,
    displayName: string,
    folderClass?: string | undefined
  ): Promise<CreateFolderResponse> {
    throw new Error("Method not implemented.");
  }

  public async copyItemsAsync(
    sourceItemIds: string[],
    targetFolderId: string
  ): Promise<CopyItemResponse> {
    throw new Error("Method not implemented.");
  }

  public async findFolderAsync(
    rootFolderId: string,
    traversal: string,
    maxEntries: number,
    pagingOffset: number
  ): Promise<FindFolderResponse> {
    if (!rootFolderId) {
      throw new Error("rootFolderId must be specified for REST.");
    }

    if (traversal !== "Shallow") {
      throw new Error("Shallow traversal is the only one supported for REST.");
    }

    const token = await this.getAccessTokenAsync();
    let url = RestService.getRestUrl(token);
    const fields = "Id,ChildFolderCount,WellKnownName";
    const extendedPropsFilter = `$filter=PropertyId eq '${RestService.PropNames.EntryId}' or PropertyId eq '${RestService.PropNames.FolderPathFullName}'`;
    url += `/api/beta/me/mailFolders/${rootFolderId}/childFolders/?$top=${maxEntries}&$skip=${pagingOffset}&$select=${fields}&$expand=SingleValueExtendedProperties(${extendedPropsFilter})`;

    const odata = await this.ajaxAsync<IODataResponse<IRestFolder>>(url, token);
    const response = new FindFolderResponse();

    for (const folder of odata.value) {
      const entryId = folder.SingleValueExtendedProperties.find(
        (prop) => prop.PropertyId === RestService.PropNames.EntryId
      )!.Value;
      const folderPathRaw = folder.SingleValueExtendedProperties.find(
        (prop) => prop.PropertyId === RestService.PropNames.FolderPathFullName
      )!.Value;
      response.folders.push({
        folderId: folder.Id,
        entryId: entryId,
        distinguishedFolderId: folder.WellKnownName,
        folderPath: folderPathRaw.replaceAll("\uFFFE", "/"),
        shortFolderId: Utils.entryIdToShortFolderId(entryId),
        childFolderCount: folder.ChildFolderCount,
      });
    }

    response.includesLastItemInRange = odata.value.length < maxEntries;
    response.responseClass = ""; // TODO
    response.responseCode = ""; // TODO
    response.indexedPagingOffset = pagingOffset + odata.value.length;

    return response;
  }

  private async ajaxAsync<TResponse>(
    url: string,
    accessToken: string,
    method = "GET"
  ) {
    return new Promise<TResponse>((resolve, reject) => {
      $.ajax({
        method: method,
        url: url,
        dataType: "json",
        headers: { Authorization: "Bearer " + accessToken },
      })
        .done((data) => resolve(data))
        .fail((error) => reject(error));
    });
  }

  private async getAccessTokenAsync(): Promise<string> {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.getCallbackTokenAsync(
        { isRest: true },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            return resolve(result.value);
          }

          return reject(
            new Error(
              `Failed to get callback token. Error: ${result.status}, Diag: ${result.diagnostics}`
            )
          );
        }
      );
    });
  }
}

interface IODataResponse<TData> {
  "@odata.context": string;
  value: TData[];
  "@odata.nextLink": string;
}

interface IRestFolder {
  "@odata.type"?: string;
  Id: string;
  WellKnownName: string | null;
  ChildFolderCount: number;
  SingleValueExtendedProperties: [
    {
      PropertyId: string;
      Value: string;
    }
  ];
}
