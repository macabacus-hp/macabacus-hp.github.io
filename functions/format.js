// import { handleChange } from "../eventhandlers";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

export async function blueBlackToggle() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      const propertiesToGet = range.getCellProperties({
        format: {
          font: {
            color: true
          }
        }
      });
      await context.sync();
      const fontColor = propertiesToGet.value[0][0].format.font.color;
      if (fontColor === "#0000FF") {
        range.format.font.color = "black";
      } else {
        range.format.font.color = "blue";
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function fontColorCycle(colors) {
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

export async function fillColorCycle(colors) {
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

export async function borderColorCycle(colors) {
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
    // store.dispatch(addBorderCycleIndex());
  } catch (error) {
    console.error(error);
  }
}

export async function chartColorCycle(colors) {
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

export const changeColor = async ({ extra: fontOrFill, name: color }) => {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      fontOrFill === "fill" ? (range.format.fill.color = color) : (range.format.font.color = color);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
};

export const changeBorderColor = async ({ name: color }) => {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      const borders = range.format.borders.load("items");
      await context.sync();
      borders.items.forEach(border => {
        if (border.style != "None") {
          border.color = color;
        }
      });
    });
  } catch (error) {
    console.error(error);
  }
};

export const changeChartColor = async ({ name: color }) => {
  try {
    await Excel.run(async context => {
      const charts = context.workbook.getActiveChartOrNullObject();
      charts.format.fill.setSolidColor(color);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
};

// export const autocolorSelection = async () => {
//   await Excel.run(async context => {
//     const usedRange = context.workbook.getSelectedRange();

//     //getting all of the Input Ranges
//     let inputRanges = usedRange.getSpecialCellsOrNullObject(
//       Excel.SpecialCellType.constants,
//       Excel.SpecialCellValueType.numbers
//     );

//     //getting all of the Formula Ranges
//     let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);

//     //getting all of the Hyperlink Ranges
//     let hyperLinkRanges = usedRange;
//     let hyper = [];
//     hyperLinkRanges.load(["rowCount", "columnCount"]);
//     await context.sync();
//     for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
//       for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
//         hyper.push(
//           hyperLinkRanges
//             .getCell(r, c)
//             .untrack()
//             .load(["hyperlink", "values"])
//         );
//       }
//     }
//     await context.sync();
//     context.application.suspendScreenUpdatingUntilNextSync();
//     let counter = 0;
//     for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
//       for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
//         if (hyper[counter].hyperlink) {
//           hyperLinkRanges.getCell(
//             r,
//             c
//           ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[4];
//         }
//         if (hyper[counter].values[0][0].length) {
//           const i = hyper[counter].values[0][0].indexOf("[");
//           if (i > 1) {
//             const j = hyper[counter].values[0][0].indexOf("]");
//             if (j > i) {
//               const k = hyper[counter].values[0][0].indexOf("!");
//               if (k > j) {
//                 hyperLinkRanges.getCell(
//                   r,
//                   c
//                 ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[3];
//               }
//             }
//           }
//         }
//         if (hyper[counter].values[0][0].length && hyper[counter].values[0][0].indexOf("!") !== -1) {
//           hyperLinkRanges.getCell(
//             r,
//             c
//           ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[2];
//         }
//         counter++;
//       }
//     }
//     if (inputRanges !== null) {
//       var inputColor = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[0];
//       inputRanges.format.font.color = inputColor;
//     }
//     if (formulaRanges !== null) {
//       var formulaColor = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[1];
//       formulaRanges.format.font.color = formulaColor;
//     }

//     await context.sync();
//   });
// };

// export const autocolorSheet = async () => {
//   try {
//     await Excel.run(async context => {
//       let usedRange = context.workbook.worksheets.getActiveWorksheet().getUsedRange();

//       //getting all of the Input Ranges
//       let inputRanges = usedRange.getSpecialCellsOrNullObject(
//         Excel.SpecialCellType.constants,
//         Excel.SpecialCellValueType.numbers
//       );

//       //getting all of the Formula Ranges
//       let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);

//       //getting all of the Hyperlink Ranges
//       let hyperLinkRanges = usedRange;

//       let hyper = [];
//       hyperLinkRanges.load(["rowCount", "columnCount"]);
//       await context.sync();
//       for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
//         for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
//           hyper.push(
//             hyperLinkRanges
//               .getCell(r, c)
//               .untrack()
//               .load(["hyperlink", "values", "formulas"])
//           );
//         }
//       }
//       await context.sync();
//       context.application.suspendScreenUpdatingUntilNextSync();
//       let counter = 0;
//       for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
//         for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
//           //hyperlink
//           if (hyper[counter].hyperlink) {
//             hyperLinkRanges.getCell(
//               r,
//               c
//             ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[4];
//           }
//           //workbook links
//           if (hyper[counter].values[0][0].length) {
//             const i = hyper[counter].values[0][0].indexOf("[");
//             console.log(i);
//             if (i > 1) {
//               const j = hyper[counter].values[0][0].indexOf("]");
//               console.log(j);
//               if (j > i) {
//                 const k = hyper[counter].values[0][0].indexOf("!");
//                 console.log(k);
//                 if (k > j) {
//                   hyperLinkRanges.getCell(
//                     r,
//                     c
//                   ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[3];
//                 }
//               }
//             }
//             //worksheet links
//             else if (hyper[counter].values[0][0].indexOf("!") !== -1) {
//               hyperLinkRanges.getCell(
//                 r,
//                 c
//               ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[2];
//             }
//           }

//           counter++;
//         }
//       }
//       if (inputRanges !== null) {
//         var inputColor = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[0];
//         inputRanges.format.font.color = inputColor;
//       }
//       if (formulaRanges !== null) {
//         var formulaColor = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[1];
//         formulaRanges.format.font.color = formulaColor;
//       }
//       await context.sync();
//     });
//   } catch (error) {
//     console.log(error);
//   }
// };

// export const autocolorWorkbook = async () => {
//   try {
//     await Excel.run(async context => {
//       let workbook = context.workbook.worksheets.load();

//       await context.sync();
//       for (var i = 0; i < workbook.items.length; i++) {
//         let usedRange = workbook.items[i].getUsedRange();

//         //getting all of the InputRanges
//         let inputRanges = usedRange.getSpecialCellsOrNullObject(
//           Excel.SpecialCellType.constants,
//           Excel.SpecialCellValueType.numbers
//         );

//         //getting all of the Formula Ranges
//         let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);

//         //getting all of the Hyperlink Ranges
//         let hyperLinkRanges = usedRange;
//         let hyper = [];
//         hyperLinkRanges.load(["rowCount", "columnCount"]);
//         await context.sync();
//         for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
//           for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
//             hyper.push(
//               hyperLinkRanges
//                 .getCell(r, c)
//                 .untrack()
//                 .load(["hyperlink", "values"])
//             );
//           }
//         }

//         await context.sync();
//         context.application.suspendScreenUpdatingUntilNextSync();
//         let counter = 0;
//         for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
//           for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
//             if (hyper[counter].hyperlink) {
//               hyperLinkRanges.getCell(
//                 r,
//                 c
//               ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[4];
//             }
//             if (hyper[counter].values[0][0].length && hyper[counter].values[0][0].indexOf("!") !== -1) {
//               hyperLinkRanges.getCell(
//                 r,
//                 c
//               ).format.font.color = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[2];
//             }
//             counter++;
//           }
//         }
//         if (inputRanges !== null) {
//           var inputColor = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[0];
//           inputRanges.format.font.color = inputColor;
//         }
//         if (formulaRanges !== null) {
//           var formulaColor = store.getState().colorsReducer.menuItems[0].items[11].items[0].extra[1];
//           formulaRanges.format.font.color = formulaColor;
//         }

//         await context.sync();
//       }
//     });
//   } catch (error) {
//     console.log(error);
//   }
// };

// export const autocolorOnEntry = async () => {
//   if (!store.getState().colorsReducer.autocolorToggle) {
//     try {
//       await Excel.run(async context => {
//         let workbook = context.workbook.worksheets;
//         workbook.load();
//         // let eventResult = workbook.onChanged.add(handleChange);
//         // Office.context.document.settings.set("autoColorOnEntry", eventResult);
//         await context.sync();
//         console.log("Event handler successfully registered for onChanged event in the worksheet: ");
//       });
//     } catch (error) {
//       console.error(error);
//     }
//   } else {
//     try {
//       let eventResult = Office.context.document.settings.get("autoColorOnEntry");
//       await Excel.run(eventResult.context, async context => {
//         eventResult.remove();
//         await context.sync();
//         eventResult = null;
//         console.log("Event handler successfully removed.");
//       });
//     } catch (error) {
//       console.error(error);
//     }
//   }
//   store.dispatch(autocolorToggle());
// };

export const autocolorSelection = async autoColor => {
  await Excel.run(async context => {
    const usedRange = context.workbook.getSelectedRange();

    //getting all of the Input Ranges
    let inputRanges = usedRange.getSpecialCellsOrNullObject(
      Excel.SpecialCellType.constants,
      Excel.SpecialCellValueType.numbers
    );

    //getting all of the Formula Ranges
    let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);

    //getting all of the Hyperlink Ranges
    let hyperLinkRanges = usedRange;
    let hyper = [];
    hyperLinkRanges.load(["rowCount", "columnCount"]);
    await context.sync();
    for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
      for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
        hyper.push(
          hyperLinkRanges
            .getCell(r, c)
            .untrack()
            .load(["hyperlink", "values"])
        );
      }
    }
    await context.sync();
    // context.application.suspendScreenUpdatingUntilNextSync();
    let counter = 0;
    for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
      for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
        if (hyper[counter].hyperlink) {
          hyperLinkRanges.getCell(r, c).format.font.color = autoColor[4].data;
        }
        if (hyper[counter].values[0][0].length && hyper[counter].hyperlink) {
          const i = hyper[counter].values[0][0].indexOf("[");
          if (i > 1) {
            const j = hyper[counter].values[0][0].indexOf("]");
            if (j > i) {
              const k = hyper[counter].values[0][0].indexOf("!");
              if (k > j) {
                hyperLinkRanges.getCell(r, c).format.font.color = autoColor[3].data;
              }
            }
          }
        }
        if (
          hyper[counter].values[0][0].length &&
          hyper[counter].values[0][0].indexOf("!") !== -1 &&
          hyper[counter].hyperlink
        ) {
          hyperLinkRanges.getCell(r, c).format.font.color = autoColor[2].data;
        }
        counter++;
      }
    }
    if (inputRanges !== null) {
      inputRanges.format.font.color = autoColor[0].data;
    }
    if (formulaRanges !== null) {
      formulaRanges.format.font.color = autoColor[1].data;
    }

    await context.sync();
  });
};

export const autocolorSheet = async autoColor => {
  try {
    await Excel.run(async context => {
      let usedRange = context.workbook.worksheets.getActiveWorksheet().getUsedRange();

      //getting all of the Input Ranges
      let inputRanges = usedRange.getSpecialCellsOrNullObject(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers
      );

      //getting all of the Formula Ranges
      let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);

      //getting all of the Hyperlink Ranges
      let hyperLinkRanges = usedRange;

      let hyper = [];
      hyperLinkRanges.load(["rowCount", "columnCount"]);
      await context.sync();
      for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
        for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
          hyper.push(
            hyperLinkRanges
              .getCell(r, c)
              .untrack()
              .load(["hyperlink", "values", "formulas"])
          );
        }
      }
      await context.sync();
      context.application.suspendScreenUpdatingUntilNextSync();
      let counter = 0;
      for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
        for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
          //hyperlink
          if (hyper[counter].hyperlink) {
            hyperLinkRanges.getCell(r, c).format.font.color = autoColor[4].data;
          }
          //workbook links
          if (hyper[counter].values[0][0].length && hyper[counter].hyperlink) {
            const i = hyper[counter].values[0][0].indexOf("[");
            console.log(i);
            if (i > 1) {
              const j = hyper[counter].values[0][0].indexOf("]");
              console.log(j);
              if (j > i) {
                const k = hyper[counter].values[0][0].indexOf("!");
                console.log(k);
                if (k > j) {
                  hyperLinkRanges.getCell(r, c).format.font.color = autoColor[3].data;
                }
              }
            }
            //worksheet links
            else if (hyper[counter].values[0][0].indexOf("!") !== -1) {
              hyperLinkRanges.getCell(r, c).format.font.color = autoColor[2].data;
            }
          }

          counter++;
        }
      }
      if (inputRanges !== null) {
        inputRanges.format.font.color = autoColor[0].data;
      }
      if (formulaRanges !== null) {
        formulaRanges.format.font.color = autoColor[1].data;
      }
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
};

export const autocolorWorkbook = async autoColor => {
  try {
    await Excel.run(async context => {
      let workbook = context.workbook.worksheets.load();

      await context.sync();
      for (var i = 0; i < workbook.items.length; i++) {
        let usedRange = workbook.items[i].getUsedRange();

        //getting all of the InputRanges
        let inputRanges = usedRange.getSpecialCellsOrNullObject(
          Excel.SpecialCellType.constants,
          Excel.SpecialCellValueType.numbers
        );

        //getting all of the Formula Ranges
        let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);

        //getting all of the Hyperlink Ranges
        let hyperLinkRanges = usedRange;
        let hyper = [];
        hyperLinkRanges.load(["rowCount", "columnCount"]);
        await context.sync();
        for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
          for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
            hyper.push(
              hyperLinkRanges
                .getCell(r, c)
                .untrack()
                .load(["hyperlink", "values"])
            );
          }
        }

        await context.sync();
        context.application.suspendScreenUpdatingUntilNextSync();
        let counter = 0;
        for (var r = 0; r < hyperLinkRanges.rowCount; r++) {
          for (var c = 0; c < hyperLinkRanges.columnCount; c++) {
            if (hyper[counter].hyperlink) {
              hyperLinkRanges.getCell(r, c).format.font.color = autoColor[4].data;
            }
            if (hyper[counter].values[0][0].length && hyper[counter].hyperlink) {
              const i = hyper[counter].values[0][0].indexOf("[");
              console.log(i);
              if (i > 1) {
                const j = hyper[counter].values[0][0].indexOf("]");
                console.log(j);
                if (j > i) {
                  const k = hyper[counter].values[0][0].indexOf("!");
                  console.log(k);
                  if (k > j) {
                    hyperLinkRanges.getCell(r, c).format.font.color = autoColor[3].data;
                  }
                }
              }
            }
            if (
              hyper[counter].values[0][0].length &&
              hyper[counter].values[0][0].indexOf("!") !== -1 &&
              hyper[counter].hyperlink
            ) {
              hyperLinkRanges.getCell(r, c).format.font.color = autoColor[2].data;
            }
            counter++;
          }
        }
        if (inputRanges !== null) {
          inputRanges.format.font.color = autoColor[0].data;
        }
        if (formulaRanges !== null) {
          formulaRanges.format.font.color = autoColor[1].data;
        }

        await context.sync();
      }
    });
  } catch (error) {
    console.log(error);
  }
};

export const cycleNumberFormat = async (type, numbers) => {
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
const numberFormatValuesMap = name => {
  console.log("N: ", name);
  switch (name) {
    case "percent":
      return [[0.12345], [-0.12345], ["Text"]];
    case "date":
      return [[""], [new Date().toLocaleDateString], [""]];
    case "binary":
      return [[1], [0], ["Text"]];
    case "ratio":
      return [[0.25], [-0.25], ["Text"]];
    case "multiple":
      return [[12.345], [-12.345], ["Text"]];
    default:
      return [[1234.56789], [-1234.56789], ["Text"]];
  }
};
export const numberFormatTest = async (numberFormatString, dialog) => {
  let err = "";
  console.log(numberFormatString);
  // await startDialogEvents();
  try {
    await Excel.run(async context => {
      context.application.suspendScreenUpdatingUntilNextSync();
      let sheets = context.workbook.worksheets;
      let previewSheet = sheets.getItemOrNullObject("Preview");
      await context.sync();
      if (previewSheet.isNullObject) {
        previewSheet = sheets.add("Preview");
      }
      previewSheet.visibility = Excel.SheetVisibility.hidden;

      let range = previewSheet.getRange("A1:A3");

      await context.sync();

      range.numberFormat = [[numberFormatString[0]]];
      range.load();
      await context.sync();
      range.format.horizontalAlignment = numberFormatString[2];
      // context.application.suspendScreenUpdatingUntilNextSync();
      range.values = numberFormatValuesMap(numberFormatString[1]);
      let img = await range.getImage();
      await previewSheet.delete();

      await context.sync();
      err = img.value;
    });
  } catch (error) {
    err = "error";
    console.log(error.message);
  }
  dialog.messageChild(err);
  // await saveCancelEvents();
};

export async function startDialogEvents() {
  try {
    await Excel.run(async context => {
      let sheets = context.workbook.worksheets;
      const sheet = sheets.add("Preview");
      sheet;
      // sheet.visibility = Excel.SheetVisibility.hidden;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
export async function saveCancelEvents() {
  try {
    await Excel.run(async context => {
      let sheets = context.workbook.worksheets;
      let sheet = sheets.getItem("Preview");
      sheet.delete();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
export async function borderStyle(type) {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.format.borders.getItem(type).style = "Continuous";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
export async function noBorder() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.format.borders.getItem("InsideHorizontal").style = "Continuous";
      range.format.borders.getItem("InsideVertical").style = "Continuous";
      range.format.borders.getItem("EdgeBottom").style = "Continuous";
      range.format.borders.getItem("EdgeLeft").style = "Continuous";
      range.format.borders.getItem("EdgeRight").style = "Continuous";
      range.format.borders.getItem("EdgeTop").style = "Continuous";
      range.format.borders.getItem("InsideHorizontal").style = "None";
      range.format.borders.getItem("InsideVertical").style = "None";
      range.format.borders.getItem("EdgeBottom").style = "None";
      range.format.borders.getItem("EdgeLeft").style = "None";
      range.format.borders.getItem("EdgeRight").style = "None";
      range.format.borders.getItem("EdgeTop").style = "None";
      range.format.borders.getItem("DiagonalDown").style = "None";
      range.format.borders.getItem("DiagonalUp").style = "None";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function doubleBorderStyle({ extra: type }) {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      type.forEach(element => {
        range.format.borders.getItem(element).style = "Continuous";
      });
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function noBorderStyle() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.format.borders.getItem("InsideHorizontal").style = "Continuous";
      range.format.borders.getItem("InsideVertical").style = "Continuous";
      range.format.borders.getItem("EdgeBottom").style = "Continuous";
      range.format.borders.getItem("EdgeLeft").style = "Continuous";
      range.format.borders.getItem("EdgeRight").style = "Continuous";
      range.format.borders.getItem("EdgeTop").style = "Continuous";
      range.format.borders.getItem("InsideHorizontal").style = "None";
      range.format.borders.getItem("InsideVertical").style = "None";
      range.format.borders.getItem("EdgeBottom").style = "None";
      range.format.borders.getItem("EdgeLeft").style = "None";
      range.format.borders.getItem("EdgeRight").style = "None";
      range.format.borders.getItem("EdgeTop").style = "None";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function alignmentFormat({ extra: [alignmentOrientation, formatParam] }) {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.format[alignmentOrientation] = formatParam;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function changeFontStyle(fontStyleName) {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.format.font.name = fontStyleName;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function changeFontSize(fontSize) {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.format.font.size = fontSize;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

const fontSizeArr = [8, 9, 10, 11, 12, 14, 16, 20, 22, 24, 26, 28, 36, 48, 72];
const getClosestNumber = num => {
  // return fontSizeArr.reduce(function(prev, curr) {
  //   return Math.abs(curr - goal) < Math.abs(prev - goal) ? curr : prev;
  // });
  let curr = fontSizeArr[0];
  let currIndex = 0;
  for (let i = 0; i <= fontSizeArr.length - 1; i++) {
    let val = fontSizeArr[i];
    if (Math.abs(num - val) < Math.abs(num - curr)) {
      curr = val;
      currIndex = i;
    }
  }
  return currIndex;
};

export async function increaseFont() {
  try {
    await Excel.run(async context => {
      context.application.suspendApiCalculationUntilNextSync();
      let range = context.workbook.getSelectedRange();
      const props = range.format.font.load("size");
      await context.sync();
      let size = props.size > 72 ? props.size : fontSizeArr[getClosestNumber(props.size) + 1];
      range.format.font.size = size;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function decreaseFont() {
  try {
    await Excel.run(async context => {
      context.application.suspendApiCalculationUntilNextSync();
      let range = context.workbook.getSelectedRange();
      const props = range.format.font.load("size");
      await context.sync();
      let size = 0;
      if (props.size < 8) {
        size = props.size;
      } else if (props.size > 72) {
        size = 72;
      } else {
        size = fontSizeArr[getClosestNumber(props.size) - 1];
      }
      range.format.font.size = size;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function underlineStyle(underlineType) {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.format.font.underline = underlineType;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function leaderDots() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.numberFormat = [["@*."]];
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}

export async function casePicker(type) {
  switch (type) {
    case 1:
      sentenceCase();
      break;
    case 2:
      titleCase();
      break;
    case 3:
      upperCase();
      break;
    case 4:
      lowerCase();
      break;
    default:
      break;
  }
}
export async function sentenceCase() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange().load();
      await context.sync();

      let arr = [];
      for (let i = 0; i < range.values.length; i++) {
        arr.push([]);
        for (let j = 0; j < range.values[i].length; j++) {
          let value = range.values[i][j];
          if (value !== "") {
            value = value.toLowerCase();
            let rg = /(^\w{1}|\.\s*\w{1})/gi;
            value = value.replace(rg, function(toReplace) {
              return toReplace.toUpperCase();
            });
          }
          arr[i].push(value);
        }
      }
      range.values = arr;
    });
  } catch (error) {
    console.log(error);
  }
}

export async function titleCase() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange().load();
      await context.sync();

      let arr = [];
      for (let i = 0; i < range.values.length; i++) {
        arr.push([]);
        for (let j = 0; j < range.values[i].length; j++) {
          let value = range.values[i][j];
          if (value !== "") {
            value = value.replace(/\w\S*/g, function(txt) {
              return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
            });
          }
          arr[i].push(value);
        }
      }
      range.values = arr;
    });
  } catch (error) {
    console.log(error);
  }
}

export async function upperCase() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange().load();
      await context.sync();

      let arr = [];
      for (let i = 0; i < range.values.length; i++) {
        arr.push([]);
        for (let j = 0; j < range.values[i].length; j++) {
          let value = range.values[i][j];
          if (value !== "") {
            value = value.toUpperCase();
          }
          arr[i].push(value);
        }
      }
      range.values = arr;
    });
  } catch (error) {
    console.log(error);
  }
}
export async function lowerCase() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange().load();
      await context.sync();

      let arr = [];
      for (let i = 0; i < range.values.length; i++) {
        arr.push([]);
        for (let j = 0; j < range.values[i].length; j++) {
          let value = range.values[i][j];
          if (value !== "") {
            value = value.toLowerCase();
          }
          arr[i].push(value);
        }
      }
      range.values = arr;
    });
  } catch (error) {
    console.log(error);
  }
}
export async function sumBar() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange().load();
      await context.sync();
      range.format.rowHeight = 5;
      let arr = [];
      for (let i = 0; i < range.values.length; i++) {
        arr.push([]);
        for (let j = 0; j < range.values[i].length; j++) {
          let value = " ";
          arr[i].push(value);
        }
      }
      range.values = arr;
      range.format.font.underline = "SingleAccountant";
    });
  } catch (error) {
    console.log(error);
  }
}

export async function wrapText() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange().load();
      await context.sync();
      range.format.wrapText = true;
    });
  } catch (error) {
    console.log(error);
  }
}
export async function capturePaintbrush() {
  try {
    await Excel.run(async context => {
      let paintBrushArray = Office.context.document.settings.get("paintBrushArray");

      let range = context.workbook.getSelectedRange();

      range.load("address");
      await context.sync();

      paintBrushArray.push(range.address);
      Office.context.document.settings.set("paintBrushArray", paintBrushArray);
      Office.context.document.settings.saveAsync();
      console.log(range.address);
    });
  } catch (error) {
    console.log(error);
  }
}

export async function applyPaintbrush() {
  try {
    await Excel.run(async context => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let currentRange = context.workbook.getSelectedRange().load("address");
      let paintBrushArray = Office.context.document.settings.get("paintBrushArray");

      let range = paintBrushArray[0];
      await context.sync();

      sheet.getRange(currentRange.address).copyFrom(
        range,
        Excel.RangeCopyType.formats,
        false, // skipBlanks
        false
      );
      await context.sync();

      // paintBrushArray.push(range.address);
      // Office.context.document.settings.set("paintBrushArray", paintBrushArray);
      // Office.context.document.settings.saveAsync();
    });
  } catch (error) {
    console.log(error);
  }
}

export function clearPaintbrush() {
  try {
    Office.context.document.settings.set("paintBrushArray", []);
    Office.context.document.settings.saveAsync();
  } catch (error) {
    console.log(error);
  }
}

export async function listLogicDirector(type) {
  switch (type) {
    case "bullet":
      listFormat("• @");
      break;
    case "dash":
      listFormat("- @");
      break;
    case "numbers":
      numbersListFormat();
      break;
    case "upperCaseLetters":
      upperCaseListFormat();
      break;
    case "lowerCaseLetters":
      lowerCaseListFormat();
      break;
    case "upperCaseRoman":
      upperCaseRomanListFormat();
      break;
    case "lowerCaseRoman":
      lowerCaseRomanListFormat();
      break;
    default:
      listFormat("");
      break;
  }
}
export async function listFormat(formatParam) {
  console.log(formatParam);
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.numberFormat = [[formatParam]];
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}

export async function numbersListFormat() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.load("rowCount, columnCount");
      await context.sync();
      let i = 0;
      for (var r = 0; r < range.rowCount; r++) {
        for (var c = 0; c < range.columnCount; c++) {
          i++;
          if (i <= 10) range.getCell(r, c).numberFormat = [[`"${i}." @`]];
        }
      }
      console.log(i);
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}

export async function lowerCaseListFormat() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.load("rowCount, columnCount");
      await context.sync();
      let i = 0;
      for (var r = 0; r < range.rowCount; r++) {
        for (var c = 0; c < range.columnCount; c++) {
          i++;
          if (i <= 26) {
            let alphabet = String.fromCharCode(96 + i);
            let format = `"${alphabet}." @`;
            range.getCell(r, c).numberFormat = [[format]];
          }
        }
      }
      console.log(i);
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}

export async function upperCaseListFormat() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.load("rowCount, columnCount");
      await context.sync();
      let i = 0;
      for (var r = 0; r < range.rowCount; r++) {
        for (var c = 0; c < range.columnCount; c++) {
          i++;
          if (i <= 26) {
            let alphabet = String.fromCharCode(64 + i);
            let format = `"${alphabet}." @`;
            range.getCell(r, c).numberFormat = [[format]];
          }
        }
      }
      console.log(i);
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}

function romanize(num) {
  if (isNaN(num)) return NaN;
  var digits = String(+num).split(""),
    key = [
      "",
      "C",
      "CC",
      "CCC",
      "CD",
      "D",
      "DC",
      "DCC",
      "DCCC",
      "CM",
      "",
      "X",
      "XX",
      "XXX",
      "XL",
      "L",
      "LX",
      "LXX",
      "LXXX",
      "XC",
      "",
      "I",
      "II",
      "III",
      "IV",
      "V",
      "VI",
      "VII",
      "VIII",
      "IX"
    ],
    roman = "",
    i = 3;
  while (i--) roman = (key[+digits.pop() + i * 10] || "") + roman;
  return Array(+digits.join("") + 1).join("M") + roman;
}

export async function upperCaseRomanListFormat() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.load("rowCount, columnCount");
      await context.sync();
      let i = 0;
      for (var r = 0; r < range.rowCount; r++) {
        for (var c = 0; c < range.columnCount; c++) {
          i++;
          if (i <= 26) {
            let roman = romanize(i);
            let format = `"${roman}." @`;
            range.getCell(r, c).numberFormat = [[format]];
          }
        }
      }
      console.log(i);
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}
export async function lowerCaseRomanListFormat() {
  try {
    await Excel.run(async context => {
      let range = context.workbook.getSelectedRange();
      range.load("rowCount, columnCount");
      await context.sync();
      let i = 0;
      for (var r = 0; r < range.rowCount; r++) {
        for (var c = 0; c < range.columnCount; c++) {
          i++;
          if (i <= 26) {
            let roman = romanize(i)
              .toString()
              .toLowerCase();
            let format = `"${roman}." @`;
            range.getCell(r, c).numberFormat = [[format]];
          }
        }
      }
      console.log(i);
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}
