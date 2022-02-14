/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import RequestContext = Excel.RequestContext;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
  }
});

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
function handleSelectionChange(): Promise<void> {
  document.querySelector("textarea").value = "";
  return null;
}

export async function run(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      await syncTextboxAndActiveCell(context);
      context.workbook.onSelectionChanged.add(handleSelectionChange);
    });
  } catch (error) {
    console.error(error);
  }
}
