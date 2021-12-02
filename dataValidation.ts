/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

export async function number() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.dataValidation.clear();

      range.dataValidation.errorAlert = {
        message: "Enter a number here.",
        showAlert: true, // default is 'true'
        style: "Stop", // other possible values: Warning, Information
        title: "Macabacus"
      };

      range.dataValidation.prompt = {
        message: "Enter a number here.",
        showPrompt: true, // default is 'false'
        title: ""
      };

      range.dataValidation.rule = {
        custom: {
          formula: "=ISNUMBER(INDIRECT(ADDRESS(ROW(), COLUMN())))"
        }
      };
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function date() {
  let strAddress = "INDIRECT(ADDRESS(ROW(), COLUMN()))"
  const intYear = new Date().getFullYear();
  let intYearMin = intYear - 30, intYearMax = intYear + 30
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
       
      range.dataValidation.clear();
      
      range.dataValidation.errorAlert = {
        message: "Enter a date here.",
        showAlert: true, // default is 'true'
        style: "Stop", // other possible values: Warning, Information
        title: "Macabacus"
      };
      range.dataValidation.prompt = {
        message: "Enter a date here.",
        showPrompt: true, // default is 'false'
        title: ""
      };

      range.dataValidation.rule = {
        custom:{
          formula: `=AND(ISNUMBER(${strAddress}),${strAddress}>DATEVALUE(""12/31/${intYearMin.toString()}""),${strAddress}<DATEVALUE(""12/31/${intYearMax.toString()}""))`
          // formula: "=AND(ISNUMBER(" + strAddress + ")," + strAddress + ">DATEVALUE(""12/31/" , intYearMin , """)," , strAddress , "<DATEVALUE(""12/31/" , intYearMax , """))"
        }
        // date: {
        //   formula1: "12/31/1000",
        //   formula2: "12/31/2100",
        //   operator: "Between"
        // }
      };
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function text() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.dataValidation.clear();

      range.dataValidation.errorAlert = {
        message: "Enter text here.",
        showAlert: true, // default is 'true'
        style: "Stop", // other possible values: Warning, Information
        title: "Macabacus"
      };
      range.dataValidation.prompt = {
        message: "Enter text here.",
        showPrompt: true, // default is 'false'
        title: ""
      };

      range.dataValidation.rule = {
        custom: {
          formula: "=ISTEXT(INDIRECT(ADDRESS(ROW(), COLUMN())))"
        }
      };
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function greaterThan() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.dataValidation.clear();

      range.dataValidation.errorAlert = {
        message: "Enter a number greater than or equal to zero.",
        showAlert: true, // default is 'true'
        style: "Stop", // other possible values: Warning, Information
        title: "Macabacus"
      };
      range.dataValidation.prompt = {
        message: "Enter a number greater than or equal to zero.",
        showPrompt: true, // default is 'false'
        title: ""
      };

      range.dataValidation.rule = {
        wholeNumber: {
          formula1: 0,
          operator: "GreaterThanOrEqualTo"
        }
      };
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function lessThan() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.dataValidation.clear();

      range.dataValidation.errorAlert = {
        message: "Enter a number less than or equal to zero.",
        showAlert: true, // default is 'true'
        style: "Stop", // other possible values: Warning, Information
        title: "Macabacus"
      };
      range.dataValidation.prompt = {
        message: "Enter a number less than or equal to zero.",
        showPrompt: true, // default is 'false'
        title: ""
      };

      range.dataValidation.rule = {
        wholeNumber: {
          formula1: 0,
          operator: "LessThanOrEqualTo"
        }
      };
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function positivePercent() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.dataValidation.clear();

      range.dataValidation.errorAlert = {
        message: "Enter a percent between 0% and 100%.",
        showAlert: true, // default is 'true'
        style: "Stop", // other possible values: Warning, Information
        title: "Macabacus"
      };
      range.dataValidation.prompt = {
        message: "Enter a percent between 0% and 100%.",
        showPrompt: true, // default is 'false'
        title: ""
      };

      range.dataValidation.rule = {
        decimal: {
          formula1: 0,
          formula2: 1.0,
          operator: "Between"
        }
      };
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function anyPercent() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.dataValidation.clear();

      range.dataValidation.errorAlert = {
        message: "Enter a percent between -100% and 100%.",
        showAlert: true, // default is 'true'
        style: "Stop", // other possible values: Warning, Information
        title: "Macabacus"
      };
      range.dataValidation.prompt = {
        message: "Enter a percent between -100% and 100%.",
        showPrompt: true, // default is 'false'
        title: ""
      };

      range.dataValidation.rule = {
        decimal: {
          formula1: -1.0,
          formula2: 1.0,
          operator: "Between"
        }
      };
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function clear() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();

      range.dataValidation.clear();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
