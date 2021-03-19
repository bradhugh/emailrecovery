import { format } from "@fluentui/utilities";
import { default as $ } from "jquery";
import { IFolder } from "./IFolder";

/**
 * Various constants related to EWS and XML parsing
 */
export class Constants {
  static messages =
    "http://schemas.microsoft.com/exchange/services/2006/messages";
  static types = "http://schemas.microsoft.com/exchange/services/2006/types";
  static soap = "http://schemas.xmlsoap.org/soap/envelope/";
  static exchangeVersion = "Exchange2013_SP1";
  static elementNodeType = 1;
}

/**
 * EWS Request templates
 */
class EwsRequestTemplates {
  static findItemRequest =
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<soap:Envelope xmlns:m="${Constants.messages}" xmlns:t="${Constants.types}" xmlns:soap="${Constants.soap}">` +
    '<soap:Header>' +
    `<t:RequestServerVersion Version="${Constants.exchangeVersion}" />` +
    '</soap:Header>' +
    '<soap:Body>' +
    '<m:FindItem Traversal="Shallow">' +
    '<m:ItemShape>' +
    '<t:BaseShape>IdOnly</t:BaseShape>' +
    '<t:AdditionalProperties>' +
    '<t:ExtendedFieldURI PropertyTag="0x348a" PropertyType="Binary" />' +
    '<t:FieldURI FieldURI="item:ItemClass" />' +
    //"<t:FieldURI FieldURI="item:DateTimeReceived" />" +
    "</t:AdditionalProperties>" +
    "</m:ItemShape>" +
    '<m:IndexedPageItemView MaxEntriesReturned="{0}" Offset="{1}" BasePoint="Beginning" />' +
    "<m:ParentFolderIds>" +
    '<t:DistinguishedFolderId Id="{2}"/>' +
    "</m:ParentFolderIds>" +
    "</m:FindItem>" +
    "</soap:Body>" +
    "</soap:Envelope>";

  static createFolderRequest =
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<soap:Envelope xmlns:m="${Constants.messages}" xmlns:t="${Constants.types}" xmlns:soap="${Constants.soap}">` +
    "<soap:Header>" +
    `<t:RequestServerVersion Version="${Constants.exchangeVersion}" />` +
    "</soap:Header>" +
    "<soap:Body>" +
    "<m:CreateFolder>" +
    "<m:ParentFolderId>" +
    '<t:DistinguishedFolderId Id="{0}"/>' +
    "</m:ParentFolderId>" +
    "<m:Folders>" +
    "<t:Folder>" +
    "<t:FolderClass>{1}</t:FolderClass>" +
    "<t:DisplayName>{2}</t:DisplayName>" +
    "</t:Folder>" +
    "</m:Folders>" +
    "</m:CreateFolder>" +
    "</soap:Body>" +
    "</soap:Envelope>";

  static copyItemRequest =
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<soap:Envelope xmlns:m="${Constants.messages}" xmlns:t="${Constants.types}" xmlns:soap="${Constants.soap}">` +
    "<soap:Header>" +
    `<t:RequestServerVersion Version="${Constants.exchangeVersion}" />` +
    "</soap:Header>" +
    "<soap:Body>" +
    "<m:CopyItem>" +
    "<m:ToFolderId>" +
    '<t:FolderId Id="{0}" />' +
    "</m:ToFolderId>" +
    "<m:ItemIds>" +
    "{1}" +
    //"<t:ItemId Id="AAMkAGM2OTc0MWE1LWU3ZmMtNGU3ZC1hNTUxLTU5ZDgyMTE0N2RmMQBGAAAAAABU/uPTQj8USZOiHoxSMgyHBwAw7pbBrDnFRZGSmD5m8LGJAAAJcFqPAABx2V1KEKToSItD85dL1W0eAAKdPsPKAAA=" />" +
    "</m:ItemIds>" +
    "<m:ReturnNewItemIds>true</m:ReturnNewItemIds>" +
    "</m:CopyItem>" +
    "</soap:Body>" +
    "</soap:Envelope>";

  static findFolderRequest =
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<soap:Envelope xmlns:m="${Constants.messages}" xmlns:t="${Constants.types}" xmlns:soap="${Constants.soap}">` +
    "<soap:Header>" +
    `<t:RequestServerVersion Version="${Constants.exchangeVersion}" />` +
    "</soap:Header>" +
    "<soap:Body>" +
    '<m:FindFolder Traversal="{0}">' +
    "<m:FolderShape>" +
    "<t:BaseShape>IdOnly</t:BaseShape>" +
    "<t:AdditionalProperties>" +
    '<t:FieldURI FieldURI="folder:DistinguishedFolderId" />' +
    '<t:ExtendedFieldURI PropertyTag="0x6874" PropertyType="String" />' + // FolderPathFullName
    '<t:ExtendedFieldURI PropertyTag="0x0FFF" PropertyType="Binary" />' + // PR_ENTRYID
    "</t:AdditionalProperties>" +
    "</m:FolderShape>" +
    '<m:IndexedPageFolderView MaxEntriesReturned="{1}" Offset="{2}" BasePoint="Beginning" />' +
    "<m:ParentFolderIds>" +
    '<t:DistinguishedFolderId Id="{3}"/>' +
    "</m:ParentFolderIds>" +
    "</m:FindFolder>" +
    "</soap:Body>" +
    "</soap:Envelope>";
}

export class EmailMessage {
  messageType: string = "";
  itemId: string = "";
  subject: string = "";
  lastActiveFolderId: string = "";
  itemClass: string = "";
}

export class EwsResponse {
  responseClass: string = "";
  responseCode: string = "";
}

export class FindResponse extends EwsResponse {
  indexedPagingOffset: number = 0;
  totalItemsInView: number = 0;
  includesLastItemInRange: boolean = true;
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

export class XmlParseException extends Error {
  constructor(message: string) {
    super(message);

    // When extending the built-in Error type, you have to fix up the prototype chain
    Object.setPrototypeOf(this, new.target.prototype);
  }
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

/**
 * A service for communicating with Exchange Web Services
 */
export interface IEwsService {
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

/**
 * A service for communicating with Exchange Web Services
 */
export class EwsService implements IEwsService {

  /**
   * A default instance of the EWS service
   */
  public static Default: IEwsService = new EwsService();

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
  ): Promise<FindItemResponse> {
    return new Promise((resolve, reject) => {
      var mailbox = Office.context.mailbox;
      mailbox.makeEwsRequestAsync(
        format(
          EwsRequestTemplates.findItemRequest,
          maxEntries,
          offset,
          distinguishedFolderId
        ),
        (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            var response = Parser.parseFindItemResponse(res.value);
            resolve(response);
          } else {
            reject(res.error.message);
          }
        },
        null
      );
    });
  }

  /**
   * Creates a folder using EWS
   * @param distinguishedParentFolderId the parent folder id
   * @param displayName the folder display name
   * @param folderClass the optional folder class (default IPF.Note)
   */
  createFolderAsync(
    distinguishedParentFolderId: string,
    displayName: string,
    folderClass: string = "IPF.Note"
  ): Promise<CreateFolderResponse> {
    return new Promise((resolve, reject) => {
      var mailbox = Office.context.mailbox;
      mailbox.makeEwsRequestAsync(
        format(
          EwsRequestTemplates.createFolderRequest,
          distinguishedParentFolderId,
          folderClass,
          displayName
        ),
        (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            var response = Parser.parseCreateFolderResponse(res.value);
            resolve(response);
          } else {
            reject(res.error);
          }
        },
        null
      );
    });
  }

  /**
   * Copies source items to a target folder
   * @param sourceItemIds the source item ids
   * @param targetFolderId the target folder id
   */
  copyItemsAsync(
    sourceItemIds: string[],
    targetFolderId: string
  ): Promise<CopyItemResponse> {
    return new Promise((resolve, reject) => {
      var mailbox = Office.context.mailbox;

      mailbox.makeEwsRequestAsync(
        format(
          EwsRequestTemplates.copyItemRequest,
          targetFolderId,
          Parser.createItemIdsString(sourceItemIds)
        ),
        (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            var response = Parser.parseCopyItemResponse(res.value);
            resolve(response);
          } else {
            reject(res.error);
          }
        },
        null
      );
    });
  }

  /**
   * Find subfolders of a parent
   * @param rootFolderId the root folder id
   * @param traversal the traversal type: "Deep" or "Shallow"
   * @param maxEntries the maximum number of entries to return
   * @param pagingOffset the paging offset
   */
  findFolderAsync(
    rootFolderId: string,
    traversal: string, // "Deep" or "Shallow"
    maxEntries: number,
    pagingOffset: number
  ): Promise<FindFolderResponse> {
    return new Promise<FindFolderResponse>((resolve, reject) => {
      var mailbox = Office.context.mailbox;
      mailbox.makeEwsRequestAsync(
        format(
          EwsRequestTemplates.findFolderRequest,
          traversal,
          maxEntries,
          pagingOffset,
          rootFolderId
        ),
        (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            var response = Parser.parseFindFolderResponse(res.value);
            resolve(response);
          } else {
            reject(res.error);
          }
        },
        null
      );
    });
  }
}

/**
 * Parser various items including the XML response messages
 */
class Parser {
  /**
   * Creates a string of itemID elements given a set of item IDs
   * @param ids the IDs for building the string
   */
  static createItemIdsString(ids: string[]): string {
    var idElementStrings: string[] = [];
    for (var i = 0; i < ids.length; i++) {
      idElementStrings.push(`<t:ItemId Id="${ids[i]}" />`);
    }

    return idElementStrings.join("");
  }

  /**
   * Parses the FindItem response message
   * @param xml the message XML
   */
  static parseFindItemResponse(xml: string): FindItemResponse {
    var doc = $.parseXML(xml);

    var respElem = Parser.findChildElementSingle(
      doc,
      Constants.messages,
      "FindItemResponseMessage"
    );

    var result = new FindItemResponse();
    result.responseClass = $(respElem).attr("ResponseClass") ?? "";
    result.responseCode = $(
      Parser.findChildElementSingle(
        respElem,
        Constants.messages,
        "ResponseCode"
      )
    ).text();

    var rootFolder = $(
      Parser.findChildElementSingle(respElem, Constants.messages, "RootFolder")
    );
    result.indexedPagingOffset = parseInt(
      rootFolder.attr("IndexedPagingOffset") ?? "0"
    );
    result.totalItemsInView = parseInt(
      rootFolder.attr("TotalItemsInView") ?? "0"
    );
    result.includesLastItemInRange =
      rootFolder.attr("IncludesLastItemInRange") === "true";

    // parse messages
    result.messages = [];
    var itemsElem = Parser.findChildElementSingle(
      rootFolder[0],
      Constants.types,
      "Items"
    );
    for (var i = 0; i < itemsElem.childNodes.length; i++) {
      // Check to make sure our node type is element
      if (itemsElem.childNodes[i].nodeType === Constants.elementNodeType) {
        var message = new EmailMessage();
        var itemElem = itemsElem.childNodes[i] as Element;

        message.messageType = itemElem.localName; // Contact, etc.
        message.itemId =
          $(
            Parser.findChildElementSingle(itemElem, Constants.types, "ItemId")
          ).attr("Id") ?? "";

        // Get itemClass
        message.itemClass = $(
          itemElem.getElementsByTagNameNS(Constants.types, "ItemClass")
        ).text();

        // Parse Extended properties
        var extendedProps = itemElem.getElementsByTagNameNS(
          Constants.types,
          "ExtendedProperty"
        );
        for (var iProp = 0; iProp < extendedProps.length; iProp++) {
          var extendedUri = $(
            Parser.findChildElementSingle(
              extendedProps[iProp],
              Constants.types,
              "ExtendedFieldURI"
            )
          );
          var valueElem = $(
            Parser.findChildElementSingle(
              extendedProps[iProp],
              Constants.types,
              "Value"
            )
          );

          // This is the last active folder for deleted items
          if (extendedUri.attr("PropertyTag") === "0x348a") {
            message.lastActiveFolderId = valueElem.text();
          }
        }

        result.messages.push(message);
      }
    }

    return result;
  }

  /**
   * Parses the CreateFolder response message
   * @param xml the response XML
   */
  static parseCreateFolderResponse(xml: string): CreateFolderResponse {
    var doc = $.parseXML(xml);

    var respElem = Parser.findChildElementSingle(
      doc,
      Constants.messages,
      "CreateFolderResponseMessage"
    );

    var result = new CreateFolderResponse();
    result.responseClass = $(respElem).attr("ResponseClass") ?? "";
    result.responseCode = $(
      Parser.findChildElementSingle(
        respElem,
        Constants.messages,
        "ResponseCode"
      )
    ).text();

    // Since we only currently support creating a single folder, only expect a single folder id
    var folderIdElem = Parser.findChildElementSingle(
      respElem,
      Constants.types,
      "FolderId"
    );
    result.folderId = $(folderIdElem).attr("Id") ?? "";

    return result;
  }

  /**
   * Parses the CopyItem response message
   * @param xml the response XML
   */
  static parseCopyItemResponse(xml: string): CopyItemResponse {
    var doc = $.parseXML(xml);

    var respElems = doc.getElementsByTagNameNS(
      Constants.messages,
      "CopyItemResponseMessage"
    );

    var overallResponseClass = "Success";
    var overallResponseCode = "NoError";

    var itemIds: string[] = [];

    for (var i = 0; i < respElems.length; i++) {
      var respClass = $(respElems[i]).attr("ResponseClass");
      if (respClass !== "Success") {
        overallResponseClass = respClass ?? "";
        overallResponseCode = $(
          Parser.findChildElementSingle(
            respElems[i],
            Constants.messages,
            "ResponseCode"
          )
        ).text();
      } else {
        // item-level success response
        var itemIdElem = Parser.findChildElementSingle(
          respElems[i],
          Constants.types,
          "ItemId"
        );
        const id = $(itemIdElem).attr("Id");
        if (id) {
          itemIds.push(id);
        }
      }
    }

    var result = new CopyItemResponse();
    result.responseClass = overallResponseClass;
    result.responseCode = overallResponseCode;

    result.newItemIds = itemIds;
    return result;
  }

  /**
   * Parses the FindFolder response message
   * @param xml the response XML
   */
  static parseFindFolderResponse(xml: string): FindFolderResponse {
    // Find/replace all the &#xFFFE; with / as not to break XML parsing
    xml = xml.replace(new RegExp("&#xFFFE;", "g"), "/");

    var doc = $.parseXML(xml);

    var respElem = Parser.findChildElementSingle(
      doc,
      Constants.messages,
      "FindFolderResponseMessage"
    );

    var result = new FindFolderResponse();
    result.responseClass = $(respElem).attr("ResponseClass") ?? "";
    result.responseCode = $(
      Parser.findChildElementSingle(
        respElem,
        Constants.messages,
        "ResponseCode"
      )
    ).text();

    var rootFolder = $(
      Parser.findChildElementSingle(respElem, Constants.messages, "RootFolder")
    );
    result.indexedPagingOffset = parseInt(
      rootFolder.attr("IndexedPagingOffset") ?? "0"
    );
    result.totalItemsInView = parseInt(
      rootFolder.attr("TotalItemsInView") ?? "0"
    );
    result.includesLastItemInRange =
      rootFolder.attr("IncludesLastItemInRange") === "true";

    // parse messages
    result.folders = [];
    var foldersElem = Parser.findChildElementSingle(
      respElem,
      Constants.types,
      "Folders"
    );

    for (var i = 0; i < foldersElem.childNodes.length; i++) {
      // Check to make sure our node type is element
      if (foldersElem.childNodes[i].nodeType === Constants.elementNodeType) {
        /// TODO: Type-guard
        var folderElem = foldersElem.childNodes[i] as Element;

        var folder: IFolder = {
          folderId: "",
          distinguishedFolderId: "",
          entryId: "",
          folderPath: "",
          shortFolderId: ""
        };

        var folderIdElem = Parser.findChildElementSingle(
          folderElem,
          Constants.types,
          "FolderId"
        );
        folder.folderId = $(folderIdElem).attr("Id") ?? "";

        // Get distinguished folder id
        folder.distinguishedFolderId = $(
          folderElem.getElementsByTagNameNS(
            Constants.types,
            "DistinguishedFolderId"
          )
        ).text();

        // Parse Extended properties
        var extendedProps = folderElem.getElementsByTagNameNS(
          Constants.types,
          "ExtendedProperty"
        );
        for (var iProp = 0; iProp < extendedProps.length; iProp++) {
          var extendedUri = $(
            Parser.findChildElementSingle(
              extendedProps[iProp],
              Constants.types,
              "ExtendedFieldURI"
            )
          );
          var valueElem = $(
            Parser.findChildElementSingle(
              extendedProps[iProp],
              Constants.types,
              "Value"
            )
          );

          // This is FolderPathFullName
          if (extendedUri.attr("PropertyTag") === "0x6874") {
            folder.folderPath = valueElem.text();
          } else if (extendedUri.attr("PropertyTag") === "0xfff") {
            folder.entryId = valueElem.text();
          }
        }

        // Get our short folder id from the EntryId
        folder.shortFolderId = Parser.entryIdToShortFolderId(folder.entryId);

        // add this to the folder list
        result.folders.push(folder);
      }
    }

    return result;
  }

  private static findChildElementSingle(
    parent: Element | XMLDocument,
    namespace: string,
    localName: string
  ): Element {
    var messages = parent.getElementsByTagNameNS(namespace, localName);
    if (messages.length === 0) {
      throw new XmlParseException(
        `Failed to find ${namespace}/${localName} in XML response`
      );
    }

    if (messages.length > 1) {
      throw new XmlParseException(
        `Found multiple items matching ${namespace}/${localName} in XML response`
      );
    }

    return messages[0];
  }

  private static entryIdToShortFolderId(ewsId: string): string {
    var bin = atob(ewsId);

    // We only need bytes 22 to 44 which is DBGuid + GlobalCounter
    var shortIdBin = bin.substr(22, 22);
    return btoa(shortIdBin);
  }
}
