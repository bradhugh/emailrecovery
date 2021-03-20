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
}