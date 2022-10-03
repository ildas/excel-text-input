import RequestContext = Excel.RequestContext;


export class TaskpaneElementsModifier {
    public static changeValue(cssSelector: string, value: string): void {
        document.querySelector(cssSelector)["value"] = value;
      }

      public static getElement(context: RequestContext, cssSelector: string): Element {
        return document.querySelector(cssSelector);
      }

      public static getValue(selector): string {
        return document.querySelector(selector).value;
      }
}
