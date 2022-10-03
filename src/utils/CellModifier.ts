import RequestContext = Excel.RequestContext;

export class CellModifier {

  public static getActiveCell(context: RequestContext): Excel.Range {
    return context.workbook.getActiveCell();
  }

  public static getElement(context: RequestContext, cssSelector: string): Element {
    return document.querySelector(cssSelector);
  }

  public static getValue(selector): string {
    return document.querySelector(selector).value;
  }

  public static changeValue(selectedCell: Excel.Range, value: string): void {
    selectedCell.values = [[value]];
  }
}