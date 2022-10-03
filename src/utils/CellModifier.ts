import RequestContext = Excel.RequestContext;

/* global console, document, Excel */
export class CellModifier {
  /**
   * 1. checks if the info could be changed
   * 2. change info
   */
  public static getActiveCell(context: RequestContext): Excel.Range {
    return context.workbook.getActiveCell();
  }

  public static getCell(context: RequestContext, cssSelector: string): Element {
    return document.querySelector(cssSelector);
  }

  public static getCellValue(selector): string {
    return document.querySelector(selector).value;
  }

  public static changeCellValue(selectedCell: Excel.Range, textValue: string): void {
    selectedCell.values = [[textValue]];
  }

  public static changeElementValue(cssSelector: string, value: string): void {
    document.querySelector(cssSelector)["value"] = value;
  }
}