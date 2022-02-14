import RequestContext = Excel.RequestContext;

/* global console, document, Excel, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
    document.querySelector("textarea").onchange = handleActiveCellChange;
  }
});

/**
 * Deletes the text box content if it is different than the starting value of ""
 * and the active cell changes
 */
async function handleActiveCellChange(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(deleteTextBoxContent);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 * Returns the currently active cell
 * @param context
 */
async function getActiveCell(context: RequestContext) {
  return context.workbook.getActiveCell();
}

/**
 * Sets the value of the active cell to the value of the task pane text box
 * @param context
 */
async function syncTextboxAndActiveCell(context: RequestContext): Promise<void> {
  const activeCell = await getActiveCell(context);
  const textBoxContent = document.querySelector("textarea").value;
  //changing value of active cell
  activeCell.values = [[textBoxContent]];
  await context.sync();
}
/**
 * Clears the text box when the active cell changes
 */
function deleteTextBoxContent(): Promise<void> {
  document.querySelector("textarea").value = "";
  return null;
}

export async function run(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      await syncTextboxAndActiveCell(context);
    });
  } catch (error) {
    console.error(error);
  }
}
