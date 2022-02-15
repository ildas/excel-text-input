import RequestContext = Excel.RequestContext;

/* global console, document, Excel */

export class ExcelHelper {
  log() {
    console.log("logging");
  }

  /**
   * Returns the currently active cell
   * @param context
   */
  _getActiveCell(context: RequestContext) {
    return context.workbook.getActiveCell();
  }

  /**
   * Deletes the text box content if it is different from the starting value of ""
   * and the active cell changes
   */
  async handleActiveCellChange(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        console.log("detected text box content change");
        context.workbook.onSelectionChanged.add(function () {
          document.querySelector("textarea").value = "";
          // onSelectionChanged doesn't allow handlers that return void, so
          // return null is needed
          return null;
        });
        console.log("after detected text box content change");
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }

  /**
   * Sets the value of the active cell to the value of the task pane text box
   * @param context
   */
  async syncTextboxAndActiveCell(context: RequestContext): Promise<void> {
    const activeCell = await this._getActiveCell(context);
    const textBoxContent = document.querySelector("textarea").value;
    //changing value of active cell
    activeCell.values = [[textBoxContent]];
    await context.sync();
  }
}
