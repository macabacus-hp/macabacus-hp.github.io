class Settings {
  constructor() {
    this.settings = JSON.parse(localStorage.getItem("formatSettings")) || {
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
  }
  get() {
    return this.settings;
  }
  set(newSettings) {
    this.settings = newSettings;
  }
  reset() {
    this.settings = {
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
    localStorage.setItem("formatSettings", JSON.stringify(this.settings));
  }
}

export let settings = new Settings();
