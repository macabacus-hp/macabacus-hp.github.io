/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import * as dataValidation from "./dataValidation";
import * as comments from "./functions/comments";
import * as cycles from "./functions/cycles.js";
// const eventhandlers = require("./eventhandlers.ts");
import { handleSelectionChange, handleSettingsChange } from "./functions/eventhandlers.js";
import * as format from "./functions/format.js";
import * as formulas from "./functions/formulas.js";
import * as paste from "./functions/paste.js";
import { symbols } from "./functions/singleUseFunctions";


/* global global, Office, self, window */
function initiateSettings() {
  
  if (Office.context.document.settings.get("settings") === null) {
    let settings = {
      format: {
        colors: {
          fontColor: [
            { data: "#0000FF", label: "RGB (0, 0, 255)", tooltip: "RGB(0, 0, 255)" },
            { data: "#008000", label: "RGB(0, 128, 0)", tooltip: "RGB(0, 128, 0)" },
            { data: "#993399", label: "RGB(128, 0, 128)", tooltip: "RGB(128, 0, 128)" },
            { data: "#FF0000", label: "RGB(255, 0, 0)", tooltip: "RGB(255, 0, 0)" },
            { data: "#FFFFFF", label: "RGB(255, 255, 255)", tooltip: "RGB(255, 255, 255)" },
            { data: "#000000", label: "RGB(0, 0, 0)", tooltip: "RGB(0, 0, 0)" }
          ],
          fillColor: [
            { data: "#D2F2FF", label: "RGB(201, 218, 248)", tooltip: "RGB(201, 218, 248)" },
            { data: "#F4CCCC", label: "RGB(210, 242, 255)", tooltip: "RGB(210, 242, 255)" },
            { data: "#FCE5CD", label: "RGB(244, 204, 204)", tooltip: "RGB(244, 204, 204)" },
            { data: "#1C4587", label: "RGB(252, 229, 205)", tooltip: "RGB(252, 229, 205)" },
            { data: "#C9DAF8", label: "RGB(28, 69, 135)", tooltip: "RGB(28, 69, 135)" }
          ],
          borderColor: [
            { data: "#000000", label: "RGB(0, 0, 0)", tooltip: "RGB(0, 0, 0)" },
            { data: "#FFFFFF", label: "RGB(255, 255, 255)", tooltip: "RGB(255, 255, 255)" },
            { data: "#808080", label: "RGB(128, 128, 128)", tooltip: "RGB(128, 128, 128)" },
            { data: "#CC0000", label: "RGB(204, 0, 0)", tooltip: "RGB(204, 0, 0)" }
          ],
          chartColor: [
            { data: "#0000FF", label: "RGB(0, 0, 255)", tooltip: "RGB(0, 0, 255)" },
            { data: "#008000", label: "RGB(0, 128, 0)", tooltip: "RGB(0, 128, 0)" },
            { data: "#993399", label: "RGB(128, 0, 128)", tooltip: "RGB(128, 0, 128)" },
            { data: "#FF0000", label: "RGB(255, 0, 0)", tooltip: "RGB(255, 0, 0)" }
          ],
          noAutoColor: [
            { data: "#FF0000", label: "RGB(255, 0, 0)", tooltip: "RGB(255, 0, 0)" },
            { data: "#FFFFFF", label: "RGB(255, 255, 255)", tooltip: "RGB(255, 255, 255)" }
          ],
          rowColumnShading: [
            {
              data: "shadeOddRows",
              label: "Shade Odd Rows",
              tooltip: "Apply the default shading color to the odd rows in the selection using conditional formatting"
            },
            {
              data: "shadeEvenRows",
              label: "Shade Even Rows",
              tooltip: "Apply the default shading color to the even rows in the selection using conditional formatting"
            },
            {
              data: "shadeOddColumns",
              label: "Shade Odd Columns",
              tooltip:
                "Apply the default shading color to the odd columns in the selection using conditional formatting"
            },
            {
              data: "shadeEvenColumns",
              label: "Shade Even Columns",
              tooltip:
                "Apply the default shading color to the even columns in the selection using conditional formatting"
            },
            {
              data: "removeAlternateShading",
              label: "Remove Alternate Shading",
              tooltip: "Remove conditional formatting used to shade alternate rows/columns in the selection"
            }
          ],
          defaultColors: {
            fonts: "#000000",
            borders: "#000000",
            shading: "#D3D3D3"
          },
          autoColorCycle: [
            {
              data: "#0000FF",
              label: "Inputs",
              tooltip:
                "Apply the default shading color to the odd columns in the selection using conditional formatting"
            },
            {
              data: "#000000",
              label: "Formulas",
              tooltip:
                "Apply the default shading color to the even columns in the selection using conditional formatting"
            },
            {
              data: "#008000",
              label: "Worksheet Links",
              tooltip: "Remove conditional formatting used to shade alternate rows/columns in the selection"
            },
            {
              data: "#800080",
              label: "Workbook Links",
              tooltip: "Remove conditional formatting used to shade alternate rows/columns in the selection"
            },
            {
              data: "#FF6600",
              label: "Hyperlinks",
              tooltip: "Remove conditional formatting used to shade alternate rows/columns in the selection"
            }
          ],
          autoColorOnEntry: false,
          autoColorDates: false,
          autoColorText: false
        },
        numbers: {
          general: [
            {
              name: "comma0DecLgAlign",
              label: "Comma 0 Dec Lg Align",
              data: ["_(#,##0_)_%;(#,##0)_%;_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "comma1DecLgAlign",
              label: "Comma 1 Dec Lg Align",
              data: ["_(#,##0.0_)_%;(#,##0.0)_%;_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "comma2DecLgAlign",
              label: "Comma 2 Dec Lg Align",
              data: ["_(#,##0.00_)_%;(#,##0.00)_%;_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "comma0DecNoAlign",
              label: "Comma 0 Dec No Align",
              data: ["#,##0;(#,##0);'–';@", "No"]
            }
          ],
          localCurrency: [
            {
              name: "usd0DecLgAlign",
              label: "USD 0 Dec Lg Align",
              data: ["_([$$]#,##0_)_%;([$$]#,##0)_%;_'–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "usd1DecLgAlign",
              label: "USD 1 Dec Lg Align",
              data: ["_([$$]#,##0.0_)_%;([$$]#,##0.0)_%;_'–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "usd2DecLgAlign",
              label: "USD 2 Dec Lg Align",
              data: ["_([$$]#,##0.00_)_%;([$$]#,##0.00)_%;_'–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "comma0DecNoAlign",
              label: "USD 0 Dec No Align",
              data: ["[$$]#,##0;([$$]#,##0);'–';@", "No"]
            }
          ],
          foreignCurrency: [
            {
              name: "eur0DecLgAlign",
              label: "EUR 0 Dec Lg Align",
              data: ["_([$€-2]#,##0_)_%;([$€-2]#,##0)_%;_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "gbp0DecLgAlign",
              label: "GBP Dec Lg Align",
              data: ["_([$£-809]#,##0_)_%;([$£-809]#,##0)_%;_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "yen0DecLgAlign",
              label: "YEN 0 Dec Lg Align",
              data: ["_([$¥-2]#,##0_)_%;([$¥-2]#,##0)_%;_('–'_)_%;_(@_)_%", "Right"]
            }
          ],
          percent: [
            {
              name: "percentAlignedNegPct",
              label: "Percent Aligned Neg Pct",
              data: ["_(#,##0.0%_);(#,##0.0%);_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "percentUnaligned",
              label: "Percent Unaligned",
              data: ["#,##0.0%;(#,##0.0%);'–';@", "Right"]
            },
            {
              name: "hardPercentAlignedNegPct",
              label: "Hard Percent Aligned Neg Pct",
              data: ["_(#,##0.0'%'_);(#,##0.0'%');_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "hardPercentAligned",
              label: "Hard Percent Aligned",
              data: ["#,##0.0'%';(#,##0.0'%');'–';@", "Right"]
            },
            {
              name: "libor+",
              label: "LIBOR +",
              data: ["L+0_)_%;L-0_)_%;L+0_)_%", "Right"]
            }
          ],
          multiple: [
            {
              name: "mult1DecimalAlignedNegPct",
              label: "Mult 1 Decimal Aligned Neg Pct",
              data: ["_(0.0x_)_)_';_((0.0x)_'_';_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "mult2DecimalAlignedNegPct",
              label: "Mult 2 Decimal Aligned Neg Pct",
              data: ["_(0.00x_)_)_';_((0.00x)_'_';_('–'_)_%;_(@_)_%", "Right"]
            },
            {
              name: "mult1DecimalUnaligned",
              label: "Mult 1 Decimal Unaligned",
              data: ["0.0x;(0.0x);'–'", "Right"]
            },
            {
              name: "mult2DecimalUnaligned",
              label: "Mult 2 Decimal Unaligned",
              data: ["0.00x;(0.00x);'–'", "Right"]
            }
          ],
          date: [
            {
              name: "m/d/yyyy",
              label: "m/d/yyyy",
              data: ["m/d/yyyy;@", "Right"]
            },
            {
              name: "dateTextLong",
              label: "Date Text Long",
              data: ["mmmm d, yyyy;@", "Right"]
            },
            {
              name: "dateActualYear",
              label: "Date Actual Year",
              // prettier-ignore
              data: ["0000\A", "Right"]
            },
            {
              name: "dateEstimatedYear",
              label: "Date Estimated Year",
              // prettier-ignore
              data: ["0000\E", "Right"]
            }
          ],
          binary: [
            {
              name: "yes/no",
              label: "Yes/No",
              data: ['"Yes";"ERROR";"No";"ERROR", "Right"']
            },
            {
              name: "y/n",
              label: "Y/N",
              data: ['"Y";"ERROR";"N";"ERROR", "Right"']
            },
            {
              name: "on/off",
              label: "On/Off",
              data: ['"On";"ERROR";"Off";"ERROR", "Right"']
            },
            {
              name: "true/false",
              label: "True/False",
              data: ['"True";"ERROR";"False";"ERROR", "Right"']
            }
          ],
          ratio: [
            {
              name: "exchangeRatio",
              label: "Exchange Ratio",
              data: ["0.0:1_);(0.0):1_);0.0:1_);@_)", "Right"]
            },
            {
              name: "fraction1",
              label: "Fraction 1",
              data: ["# ?/?", "Right"]
            },
            {
              name: "fraction2",
              label: "Fraction 2",
              data: ["# ??/??", "Right"]
            },
            {
              name: "fraction3",
              label: "Fraction 3",
              data: ["# ???/???", "Right"]
            },
            {
              name: "halves",
              label: "Halves",
              data: ["# ?/2", "Right"]
            },
            {
              name: "thirds",
              label: "Thirds",
              data: ["# ?/3", "Right"]
            }
          ]
        },
        font: {
          size: [7, 11, 15],
          style: ["Arial", "Times New Roman", "Courier New"]
        },
        cells: {
          rowHeight: [5, 20, 13.5],
          colHeight: [1, 20, 8.43]
        },
        other: {
          borderStyleCycle: {
            available: ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
            selected: ["11", "12", "13"]
          },
          numOfPaintbrushes: 1,
          leftIndent: 1,
          rightIndent: 1
        }
      }
    };
    Office.context.document.settings.set("settings", settings);
  }
  Office.context.document.settings.set("paintBrushArray", []);

  Office.context.document.settings.set("blueBlackToggle", 0);
  Office.context.document.settings.set("fontColorCycle", 0);
  Office.context.document.settings.set("fillColorCycle", 0);
  Office.context.document.settings.set("borderColorCycle", 0);
  Office.context.document.settings.set("chartColorCycle", 0);
  Office.context.document.settings.set("autoColorCycle", 0);

  Office.context.document.settings.set("general", 0);
  Office.context.document.settings.set("localCurrency", 0);
  Office.context.document.settings.set("foreignCurrency", 0);
  Office.context.document.settings.set("percent", 0);
  Office.context.document.settings.set("multiple", 0);
  Office.context.document.settings.set("date", 0);
  Office.context.document.settings.set("binary", 0);
  Office.context.document.settings.set("ratio", 0);

  Office.context.document.settings.set("topBorder", 0);
  Office.context.document.settings.set("bottomBorder", 0);
  Office.context.document.settings.set("leftBorder", 0);
  Office.context.document.settings.set("rightBorder", 0);
  Office.context.document.settings.set("outsideBorder", 0);
  Office.context.document.settings.set("insideBorder", 0);

  Office.context.document.settings.set("centerAlignment", 0);
  Office.context.document.settings.set("horizontalAlignment", 0);
  Office.context.document.settings.set("verticalAlignment", 0);

  Office.context.document.settings.set("leftIndentCycle", 0);
  Office.context.document.settings.set("rightIndentCycle", 0);

  Office.context.document.settings.set("fontSizeCycle", 0);

  Office.context.document.settings.set("underlineCycle", 0);

  Office.context.document.settings.set("caseCycle", 0);

  Office.context.document.settings.set("listCycle", 0);

  Office.context.document.settings.set("footnotesCycle", 0);
  
  Office.context.document.settings.saveAsync();
}
Office.initialize(async () => {
  // If needed, Office.js is ready to be called
  console.log("Start");
  initiateSettings();
  try {
    await Excel.run(async context => {
      const settings = context.workbook.settings;
      settings.onSettingsChanged.add(handleSettingsChange);
      let workbook = context.workbook.worksheets;
      workbook.load();
      await context.sync();
      for (var i = 0; i < workbook.items.length; i++) {
        workbook.items[i].onSelectionChanged.add(handleSelectionChange);
      }
      await context.sync();
      console.log("Event handler successfully registered for onSelectionChanged event in the worksheet: ");
    });
  } catch (error) {
    console.error(error);
  }
  let h = Office.context.document.settings.get("settings");
  console.log(h);
});

//COLOR
Office.actions.associate("BLUEBLACKTOGGLE", function(event) {
  format.blueBlackToggle();
  event.completed();
});
Office.actions.associate("FONTCOLORCYCLE", function() {
  cycles.fontColorCycle();
});
Office.actions.associate("FILLCOLORCYCLE", function() {
  cycles.fillColorCycle();
});
Office.actions.associate("BORDERCOLORCYCLE", function() {
  cycles.borderColorCycle();
});
Office.actions.associate("CHARTCOLORCYCLE", function() {
  cycles.chartColorCycle();
});
Office.actions.associate("AUTOCOLORCYCLE", function() {
  cycles.cycleNumberFormat("ratio");
});
Office.actions.associate("AUTOCOLORSELECTION", function() {
  format.autocolorSelection();
});
Office.actions.associate("AUTOCOLORSHEET", function() {
  cycles.cycleNumberFormat("ratio");
});
Office.actions.associate("AUTOCOLORWORKBOOK", function() {
  cycles.cycleNumberFormat("ratio");
});

Office.actions.associate("SETCOLOR", function() {
  var context = new Excel.RequestContext();
  var range = context.workbook.getSelectedRange();
  var rangeFormat = range.format;
  rangeFormat.fill.load();
  var colors = ["#FFFFFF", "#C7CC7A", "#7560BA", "#9DD9D2", "#FFE1A8", "#E26D5C"];
  return context.sync().then(function() {
    var rangeTarget = context.workbook.getSelectedRange();
    var currentColor = -1;
    for (var i = 0; i < colors.length; i++) {
      if (colors[i] == rangeFormat.fill.color) {
        currentColor = i;
        break;
      }
    }
    if (currentColor == -1) {
      currentColor = 0;
    } else if (currentColor == colors.length - 1) {
      currentColor = 0;
    } else {
      currentColor++;
    }
    rangeTarget.format.fill.color = colors[currentColor];
    return context.sync();
  });
});

Office.actions.associate("TESTCONFLICT", function() {
  // format.blueBlackToggle();
  console.log("test");
});

//NUMBERS
Office.actions.associate("GENERALNUMBERCYCLE", function() {
  cycles.cycleNumberFormat("general");
});
Office.actions.associate("LOCALNUMBERCYCLE", function() {
  cycles.cycleNumberFormat("localCurrency");
});
Office.actions.associate("FOREIGNCURRENCYCYCLE", function() {
  cycles.cycleNumberFormat("foreignCurrency");
});
Office.actions.associate("PERCENTCYCLE", function() {
  cycles.cycleNumberFormat("percent");
});
Office.actions.associate("MULTIPLECYCLE", function() {
  cycles.cycleNumberFormat("multiple");
});
Office.actions.associate("DATECYCLE", function() {
  cycles.cycleNumberFormat("date");
});
Office.actions.associate("BINARYCYCLE", function() {
  cycles.cycleNumberFormat("binary");
});
Office.actions.associate("RATIOCYCLE", function() {
  cycles.cycleNumberFormat("ratio");
});

//BORDER
Office.actions.associate("TOPBORDERCYCLE", function() {
  cycles.borderStyleCycles("topBorder", ["EdgeTop"]);
});
Office.actions.associate("BOTTOMBORDERCYCLE", function() {
  cycles.borderStyleCycles("bottomBorder", ["EdgeBottom"]);
});
Office.actions.associate("LEFTBORDERCYCLE", function() {
  cycles.borderStyleCycles("leftBorder", ["EdgeLeft"]);
});
Office.actions.associate("RIGHTBORDERCYCLE", function() {
  cycles.borderStyleCycles("rightBorder", ["EdgeRight"]);
});
Office.actions.associate("OUTSIDEBORDERCYCLE", function() {
  cycles.borderStyleCycles("outsideBorder", ["EdgeRight", "EdgeLeft", "EdgeTop", "EdgeBottom"]);
});
Office.actions.associate("INSIDEBORDERCYCLE", function() {
  cycles.borderStyleCycles("insideBorder", ["InsideHorizontal", "InsideVertical"]);
});
Office.actions.associate("NOBORDER", function() {
  format.noBorder();
});

//Alignment
Office.actions.associate("CENTERALIGNMENT", function() {
  cycles.centerAlignmentCycle([
    ["horizontalAlignment", "Center"],
    ["horizontalAlignment", "CenterAcrossSelection"],
    ["horizontalAlignment", "General"]
  ])
});
Office.actions.associate("HORIZONTALALIGNMENT", function() {
  cycles.horizontalAlignmentCycle([
    ["horizontalAlignment", "Right"],
    ["horizontalAlignment", "Left"],
    ["horizontalAlignment", "Center"],
    ["horizontalAlignment", "CenterAcrossSelection"],
    ["horizontalAlignment", "General"]
  ])
});
Office.actions.associate("VERTICALALIGNMENT", function() {
  cycles.verticalAlignmentCycle([
    ["verticalAlignment", "Center"],
    ["verticalAlignment", "Top"],
    ["verticalAlignment", "Bottom"]
  ])
});
Office.actions.associate("LEFTINDENTCYCLE", function() {
  cycles.leftIndentCycle();
});
Office.actions.associate("RIGHTINDENTCYCLE", function() {
  cycles.rightIndentCycle();
});

//Font Size
Office.actions.associate("FONTSIZECYCLE", function() {
  cycles.changeFontSizeCycle();
});
Office.actions.associate("INCREASEFONTSIZE", function() {
  format.increaseFont();
});
Office.actions.associate("DECREASEFONTSIZE", function() {
  format.decreaseFont();
});

//Underline
Office.actions.associate("UNDERLINECYCLE", function() {
  cycles.underlineCycle(["Single", "SingleAccountant", "Double", "DoubleAccountant", "None"]);
});

//Case
Office.actions.associate("CASECYCLE", function() {
  cycles.caseCycle([1, 2, 3, 4]);
});

//List
Office.actions.associate("LISTCYCLE", function() {
  cycles.listCycle([
    "bullet",
    "dash",
    "numbers",
    "upperCaseLetters",
    "lowerCaseLetters",
    "upperCaseRoman",
    "lowerCaseRoman",
    "none"
  ]);
});

//Paintbrush
Office.actions.associate("CAPTUREPAINTBRUSH", function() {
  format.capturePaintbrush();
});
Office.actions.associate("APPLYPAINTBRUSH", function() {
  format.applyPaintbrush();
});
Office.actions.associate("CLEARPAINTBRUSH", function() {
  format.clearPaintbrush();
});

//More
Office.actions.associate("LEADERDOTS", function() {
  format.leaderDots();
});
Office.actions.associate("SUMBAR", function() {
  format.sumBar();
});
Office.actions.associate("WRAPTEXT", function() {
  format.wrapText();
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function action(event) {
  try {
   dataValidation.number()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function date(event) {
  try {
    dataValidation.date()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function text(event) {
  try {
    dataValidation.text()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function greaterThan(event) {
  try {
    dataValidation.greaterThan()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function lessThan(event) {
  try {
    dataValidation.lessThan()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function positivePercent(event) {
  try {
    dataValidation.positivePercent()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function anyPercent(event) {
  try {
    dataValidation.anyPercent()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function clear(event) {
  try {
    dataValidation.clear()
  } catch (error) {
    console.error(error)
  }
  
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function deleteCommentsAndNotes(event) {
  comments.deleteCommentsAndNotes();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function deleteResolvedComments(event) {
  comments.deleteResolvedComments();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function resolveComments(event) {
  comments.resolveComments();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function reopenComments(event) {
  comments.reopenComments();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function removeAuthor(event) {
  comments.removeAuthor();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function symbolsDollar(event) {
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function copyAddress(event) {
   paste.copyAddress();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function pasteDuplicate(event) {
   paste.pasteDuplicate();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function pasteExact(event) {
  paste.pasteExact();
 // Be sure to indicate when the add-in command function is complete
 event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function pasteNumberFormats(event) {
   paste.pasteNumberFormats();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function pasteLinks(event) {
   paste.pasteLinks();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function pasteTranspose(event) {
   paste.pasteTranspose();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function pasteInsert(event) {
   paste.pasteInsert();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolDollar(event) {
   symbols("$");
 // Be sure to indicate when the add-in command function is complete
 event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolPound(event) {
  symbols("£");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolEuro(event) {
  symbols("€");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolWon(event) {
  symbols("₩");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolYenYuan(event) {
  symbols("¥");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolBitcoin(event) {
  symbols("₿");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolCent(event) {
  symbols("¢");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolSubtract(event) {
  symbols("–");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolMultiply(event) {
  symbols("×");
// Be sure to indicate when the add-in command function is complete
event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolDivide(event) {
  symbols("÷");
// Be sure to indicate when the add-in command function is complete
event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolEllipsis(event) {
  symbols("…");
// Be sure to indicate when the add-in command function is complete
event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolBullet(event) {
  symbols("•");
// Be sure to indicate when the add-in command function is complete
event.completed();
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolSection(event) {
  symbols("§");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolBeta(event) {
  symbols("β");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolDelta(event) {
  symbols("Δ");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolCopyright(event) {
  symbols("©");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolTrademark(event) {
  symbols("™");
// Be sure to indicate when the add-in command function is complete
event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function changeSymbolRegisteredTrademark(event) {
  symbols("®");
// Be sure to indicate when the add-in command function is complete
event.completed();
}


/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function fastFillRight(event) {
  formulas.fastFillRight();
// Be sure to indicate when the add-in command function is complete
event.completed();
}


/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
 async function fastFillDown(event) {
  formulas.fastFillDown();

// Be sure to indicate when the add-in command function is complete
event.completed();
}



function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
g.date = date;
g.text = text;
g.greaterThan = greaterThan;
g.lessThan = lessThan;
g.positivePercent = positivePercent;
g.anyPercent = anyPercent;
g.clear = clear;

g.deleteCommentsAndNotes = deleteCommentsAndNotes;
g.deleteResolvedComments = deleteResolvedComments;
g.resolveComments = resolveComments;
g.reopenComments = reopenComments;
g.removeAuthor = removeAuthor;

g.copyAddress = copyAddress;
g.pasteDuplicate = pasteDuplicate;
g.pasteExact = pasteExact;
g.pasteNumberFormats = pasteNumberFormats; 
g.pasteLinks = pasteLinks;
g.pasteTranspose = pasteTranspose;
g.pasteInsert = pasteInsert;

g.changeSymbolDollar = changeSymbolDollar;
g.changeSymbolPound = changeSymbolPound;
g.changeSymbolEuro = changeSymbolEuro;
g.changeSymbolWon = changeSymbolWon;
g.changeSymbolYenYuan = changeSymbolYenYuan;
g.changeSymbolBitcoin = changeSymbolBitcoin;
g.changeSymbolCent = changeSymbolCent;
g.changeSymbolSubtract = changeSymbolSubtract;
g.changeSymbolMultiply = changeSymbolMultiply;
g.changeSymbolDivide = changeSymbolDivide;
g.changeSymbolEllipsis = changeSymbolEllipsis;
g.changeSymbolBullet = changeSymbolBullet;
g.changeSymbolSection = changeSymbolSection;
g.changeSymbolBeta = changeSymbolBeta;
g.changeSymbolDelta = changeSymbolDelta;
g.changeSymbolCopyright = changeSymbolCopyright;
g.changeSymbolTrademark = changeSymbolTrademark;
g.changeSymbolRegisteredTrademark = changeSymbolRegisteredTrademark;
g.fastFillRight = fastFillRight;
g.fastFillDown = fastFillDown;