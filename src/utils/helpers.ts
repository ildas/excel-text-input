import { CellModifyingError, ClearNonCellElementValueError } from "../exceptions";
import { CellModifier } from "../utils/CellModifier";
import { TaskpaneElementsModifier } from "../utils/TaskpaneElementsModifier";


export class ExcelHelper {
  async matchTaskpaneElementToActiveCell(cellCssSelector: string): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const selectedCell = await CellModifier.getActiveCell(context);
        const textBoxValue: string = TaskpaneElementsModifier.getValue(cellCssSelector);
        CellModifier.changeValue(selectedCell, textBoxValue);
      });
    } catch (error) {
      throw new CellModifyingError(JSON.stringify(error))
    }
  }

  async clearNonCellElementValue(): Promise<void> {
    const cssSelector = "textarea";
    try {
      await Excel.run(async function() {       
        TaskpaneElementsModifier.changeValue(cssSelector, "")
      });
    } catch (error) {
      throw new ClearNonCellElementValueError(JSON.stringify(error))
    }
  }
}
