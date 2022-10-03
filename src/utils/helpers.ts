import { CellModifyingError, ClearNonCellElementValueError } from "../exceptions";
import RequestContext = Excel.RequestContext;

/* global console, document, Excel */
export class WorkSheetModifier {
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

/**
 * base class and extend from it --
 * novite imat
 */

export class ExcelHelper {

  // create a different error handler class with different exceptions

  async changeCellValue(cellCssSelector: string): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const selectedCell = await WorkSheetModifier.getActiveCell(context);
        const textBoxContent: string = WorkSheetModifier.getCellValue(cellCssSelector);
        WorkSheetModifier.changeCellValue(selectedCell, textBoxContent);
      });
    } catch (error) {
      throw new CellModifyingError('failed to change cell value')
    }
  }

  async clearNonCellElementValue(): Promise<void> {
    const cssSelector = "textarea";
    try {
      await Excel.run(async () => {
        WorkSheetModifier.changeElementValue(cssSelector, "")
      });
    } catch (error) {
      throw new ClearNonCellElementValueError(`failed to clear value of selector: ${cssSelector}`)
    }
  }
}
