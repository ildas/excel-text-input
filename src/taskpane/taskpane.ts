import { ExcelHelper } from "../utils/helpers";

const excelHelper = new ExcelHelper();
const taskpaneElementCssSelector = "textarea"

async function implementEventsForElement(elementCssSelector): Promise<void> {
  document.querySelector(elementCssSelector).oninput = excelHelper.matchTaskpaneElementToActiveCell.bind(excelHelper, elementCssSelector);
  document.querySelector(elementCssSelector).onchange = excelHelper.clearNonCellElementValue;
}

implementEventsForElement(taskpaneElementCssSelector);
