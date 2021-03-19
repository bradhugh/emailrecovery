export class Utils {
  public static setTimeoutAsync(milliseconds: number) {
    return new Promise((resolve, _) => setTimeout(resolve, milliseconds));
  }
}