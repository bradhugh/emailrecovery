import { default as sanitizeHtml } from "sanitize-html";

export class Utils {
  public static setTimeoutAsync(milliseconds: number) {
    return new Promise((resolve, _) => setTimeout(resolve, milliseconds));
  }

  public static sanitizeHtmlDefault(html: string): string {
    return sanitizeHtml(html, {
      allowedTags: [ "strong", "br" ],
    });
  }

  public static entryIdToShortFolderId(entryId: string): string {
    const bin = atob(entryId);
    if (bin.length !== 46) {
      throw new Error(`Invalid EntryId Length ${entryId}. Expected 46 bytes. Actual: ${bin.length}`);
    }

    // We only need bytes 22 to 44 which is DBGuid + GlobalCounter
    const shortIdBin = bin.substring(22, 44);
    return btoa(shortIdBin);
  }
}