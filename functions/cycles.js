import { settings } from "../settingsClass.js";
import { casePicker, listLogicDirector, underlineStyle } from "./format.js";


export function getBorderTypeBox(index) {
  switch (index) {
    case "0":
      return { style: Excel.BorderLineStyle.dot, weight: Excel.BorderWeight.thin }; //dot thin
    case "1":
      return { style: Excel.BorderLineStyle.dash, weight: Excel.BorderWeight.thin }; //dash
    case "2":
      return { style: Excel.BorderLineStyle.dashDotDot, weight: Excel.BorderWeight.thin }; // dash dot dot thin
    case "3":
      return { style: Excel.BorderLineStyle.dashDot, weight: Excel.BorderWeight.thin }; //dash dot thin
    case "4":
      return { style: Excel.BorderLineStyle.dash, weight: Excel.BorderWeight.thin }; //long dash
    case "5":
      return { style: Excel.BorderLineStyle.continuous, weight: Excel.BorderWeight.thin }; //solid thin
    case "6":
      return { style: Excel.BorderLineStyle.dashDotDot, weight: Excel.BorderWeight.medium }; //dash dot dot medium
    case "7":
      return { style: Excel.BorderLineStyle.slantDashDot, weight: Excel.BorderWeight.medium }; //slant dash dot medium
    case "8":
      return { style: Excel.BorderLineStyle.dashDot, weight: Excel.BorderWeight.medium }; //dash dot medium
    case "9":
      return { style: Excel.BorderLineStyle.dash, weight: Excel.BorderWeight.medium }; //dash medium
    case "10":
      return { style: Excel.BorderLineStyle.continuous, weight: Excel.BorderWeight.medium }; //solid medium
    case "11":
      return { style: Excel.BorderLineStyle.continuous, weight: Excel.BorderWeight.thick }; //solid thick
    case "12":
      return { style: Excel.BorderLineStyle.double, weight: Excel.BorderWeight.thick }; //double
    default:
      return { style: Excel.BorderLineStyle.none }; //None
  }
}

export async function borderStyleCycles(name, borderSelectValues) {
  let selected = settings.get().format.other.borderStyleCycle.selected;
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      const index = Office.context.document.settings.get(name);
      const borderType = getBorderTypeBox(selected[index % selected.length]);
      Office.context.document.settings.set(name, index + 1);
      Office.context.document.settings.saveAsync();
      console.log(borderType);
      borderSelectValues.forEach(val => {
        range.format.borders.getItem(val).style = borderType.style;
        if (borderType.weight) {
          range.format.borders.getItem(val).weight = borderType.weight;
        }
      });
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
export const cycleNumberFormat = async type => {
  let numbers = settings.get().format.numbers;
  const limit = numbers[type].length;
  const index = Office.context.document.settings.get(type);
  numberFormat({ extra: numbers[type][index % limit].data });
  Office.context.document.settings.set(type, index + 1);
  await Office.context.document.settings.saveAsync();
};
//numbers
export const numberFormat = async ({ extra: numberFormatProperties }) => {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.numberFormat = numberFormatProperties[0];
      range.load();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
};

export async function fontColorCycle() {
  let colors = settings.get().format.colors.fontColor;
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      await context.sync();
      const index = Office.context.document.settings.get("fontColorCycle");
      Office.context.document.settings.set("fontColorCycle", index + 1);
      Office.context.document.settings.saveAsync();
      range.format.font.color = colors[index % colors.length].data;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function fillColorCycle() {
  let colors = settings.get().format.colors.fillColor;
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      await context.sync();
      const index = Office.context.document.settings.get("fillColorCycle");
      Office.context.document.settings.set("fillColorCycle", index + 1);
      Office.context.document.settings.saveAsync();
      range.format.fill.color = colors[index % colors.length].data;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function borderColorCycle() {
  let colors = settings.get().format.colors.borderColor;
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      const borders = range.format.borders.load("items");
      await context.sync();
      borders.items.forEach(border => {
        if (border.style != "None") {
          const index = Office.context.document.settings.get("borderColorCycle");
          Office.context.document.settings.set("borderColorCycle", index + 1);
          Office.context.document.settings.saveAsync();
          border.color = colors[index % colors.length].data;
        }
      });
    });
  } catch (error) {
    console.error(error);
  }
}

export async function chartColorCycle() {
  let colors = settings.get().format.colors.chartColor;
  try {
    await Excel.run(async context => {
      const charts = context.workbook.getActiveChartOrNullObject();
      await context.sync();
      const id = charts.load("id").id;
      console.log(id);
      const index = Office.context.document.settings.get("chartColorCycle");
      Office.context.document.settings.set("chartColorCycle", index + 1);
      Office.context.document.settings.saveAsync();
      charts.format.fill.setSolidColor(colors[index % colors.length].data);
    });
  } catch (error) {
    console.error(error);
  }
}


export async function alignmentFormat(params) {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.format[params[0]] = params[1];
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function centerAlignmentCycle(params) {
  try {
    const index = Office.context.document.settings.get("centerAlignment");
    Office.context.document.settings.set("centerAlignment", index + 1);
    Office.context.document.settings.saveAsync();
    alignmentFormat(params[index % 3]);
  } catch (error) {
    console.log(error);
  }
}
export async function horizontalAlignmentCycle(params) {
  try {
    const index = Office.context.document.settings.get("horizontalAlignment");
    Office.context.document.settings.set("horizontalAlignment", index + 1);
    Office.context.document.settings.saveAsync();
    alignmentFormat(params[index % 3]);
  } catch (error) {
    console.log(error);
  }
}
export async function verticalAlignmentCycle(params) {
  try {
    const index = Office.context.document.settings.get("verticalAlignment");
    Office.context.document.settings.set("verticalAlignment", index + 1);
    Office.context.document.settings.saveAsync();
    alignmentFormat(params[index % 3]);
  } catch (error) {
    console.log(error);
  }
}

export async function leftIndentCycle() {
  const max = settings.get().format.other.leftIndent;
  const index = Office.context.document.settings.get("leftIndentCycle");
  console.log(max);

  Office.context.document.settings.set("leftIndentCycle", index + 1);
  Office.context.document.settings.saveAsync();
  await Excel.run(async context => {
    const range = context.workbook.getSelectedRange();
    range.format.horizontalAlignment = "Left";
    // range.format.indentLevel = index;
    (index + 1) % max === 0 ? (range.format.indentLevel = 0) : range.format.adjustIndent(1);

    await context.sync();
  });
}
export async function rightIndentCycle() {
  const max = settings.get().format.other.rightIndent;
  const index = Office.context.document.settings.get("rightIndentCycle");

  Office.context.document.settings.set("rightIndentCycle", index + 1);
  Office.context.document.settings.saveAsync();
  await Excel.run(async context => {
    const range = context.workbook.getSelectedRange();
    range.format.horizontalAlignment = "Right";
    // range.format.indentLevel = index;
    (index + 1) % max === 0 ? (range.format.indentLevel = 0) : range.format.adjustIndent(1);

    await context.sync();
  });
}

export async function changeFontSizeCycle() {
  const fontSizes = settings.get().format.font.size;

  const index = Office.context.document.settings.get("fontSizeCycle");
  Office.context.document.settings.set("fontSizeCycle", index + 1);
  Office.context.document.settings.saveAsync();
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.format.font.size = fontSizes[index % fontSizes.length];
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function footnotesCycle() {
  let num = Office.context.document.settings.get("footnotesCycle") % 9;
  Office.context.document.settings.set("footnotesCycle", (num + 1) % 9);
  Office.context.document.settings.saveAsync();
  let strFootnote;
  console.log(num);
  switch (num) {
    case 1:
      strFootnote = "¹";
      break;
    case 2:
      strFootnote = "²";
      break;
    case 3:
      strFootnote = "³";
      break;
    case 4:
      strFootnote = "⁴";
      break;
    case 5:
      strFootnote = "⁵";
      break;
    case 6:
      strFootnote = "⁶";
      break;
    case 7:
      strFootnote = "⁷";
      break;
    case 8:
      strFootnote = "⁸";
      break;
    default:
      strFootnote = "⁹";
      break;
  }
  console.log(strFootnote);
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      await context.sync();

      let numAddress = range.getSpecialCellsOrNullObject("Constants", "Numbers").load("address");

      let everyAddress = range.getSpecialCellsOrNullObject("Constants", "All").load("address");

      await context.sync();
      // console.log(numAddy, everyAddy)
      // let specialRangeNumbers = context.workbook.worksheets.getActiveWorksheet().getRange(numAddy.address);
      `@⁽${strFootnote}⁾`;
      everyAddress.address.split(",").forEach(address => {
        worksheet.getRange(address).numberFormat = [[`@⁽${strFootnote}⁾`]];
      });

      numAddress.address.split(",").forEach(address => {
        worksheet.getRange(address).numberFormat = [[`0⁽${strFootnote}⁾`]];
      });
      // let loadedArr = [],
      //   currentCell;
      // range.load(["rowCount", "columnCount"]);
      // await context.sync();
      // for (var r = 0; r < range.rowCount; r++) {
      //   for (var c = 0; c < range.columnCount; c++) {
      //     loadedArr.push(range.getCell(r, c).load("valueTypes"));
      //   }
      // }
      // await context.sync();
      // for (var r = 0; r < range.rowCount; r++) {
      //   for (var c = 0; c < range.columnCount; c++) {
      //     currentCell = loadedArr[r + c];
      //     if (currentCell.valueTypes[0][0] === "Double") {
      //       console.log(currentCell);
      //       currentCell.numberFormat = [["0⁽¹⁾"]];
      //     } else {
      //       currentCell.numberFormat = [["@⁽¹⁾"]];
      //     }
      //     console.log(currentCell.valueTypes[0]);
      //   }
      // }
      // for (let i = 0; i < h; i++){
      //   newRange.areas.getItemAt(i).numberFormat = [["@⁽¹⁾"]];
      // }

      // range.numberFormat = [["@⁽¹⁾"]];
      // range.valueTypes
      // range.setCellProperties([[{format: {
      //   font: {
      //     superscript: true
      //   }
      // }}]]
      // )

      // await context.sync();

      // console.log(range);
    });
  } catch (error) {
    Office.context.document.settings.set("footnotesCycle", num % 9);
    Office.context.document.settings.saveAsync();
  }
}


export async function underlineCycle(underlineCycles) {
  const index = Office.context.document.settings.get("underlineCycle");
  console.log(underlineCycles[index % underlineCycles.length]);

  Office.context.document.settings.set("underlineCycle", index + 1);
  Office.context.document.settings.saveAsync();
  underlineStyle(underlineCycles[index % underlineCycles.length]);
}

export async function caseCycle(caseCycles) {
  const index = Office.context.document.settings.get("caseCycle");

  Office.context.document.settings.set("caseCycle", index + 1);
  Office.context.document.settings.saveAsync();
  casePicker(caseCycles[index % caseCycles.length]);
}

export async function listCycle(listCycles) {
  const index = Office.context.document.settings.get("listCycle");
  Office.context.document.settings.set("listCycle", index + 1);
  Office.context.document.settings.saveAsync();
  listLogicDirector(listCycles[index % listCycles.length]);
}

