/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */


export async function copyAddress() {
  try {
      await Excel.run(async context => {
          console.log("Copy")
          let range = context.workbook.getActiveCell().load("address");
          await context.sync();
          Office.context.document.settings.set("copyAddress", range.address);
          Office.context.document.settings.saveAsync();
      });
  } catch (error) {
      console.error(error);
  }
}

export async function pasteDuplicate() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getActiveCell();
            const copyFromAddress = Office.context.document.settings.get("copyAddress");
            if(copyFromAddress === null ){
                console.log("Error");
            }
            console.log(copyFromAddress)
            range.copyFrom(copyFromAddress);
        });
    } catch (error) {
        console.error(error);
    }
}

export async function pasteExact() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getActiveCell();
            const copyFromAddress = Office.context.document.settings.get("copyAddress");
            if(copyFromAddress === null ){
                console.log("Error");
            }
            console.log(copyFromAddress)
            range.copyFrom(copyFromAddress);
        });
    } catch (error) {
        console.error(error);
    }
}

export async function pasteNumberFormats() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getActiveCell();
            const copyFromAddress = Office.context.document.settings.get("copyAddress");
            if(copyFromAddress === null ){
                console.log("Error");
            }
            console.log(copyFromAddress)
            range.copyFrom(copyFromAddress, "Formats");
        });
    } catch (error) {
        console.error(error);
    }
}

export async function pasteLinks() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getActiveCell();
            const copyFromAddress = Office.context.document.settings.get("copyAddress");
            if(copyFromAddress === null ){
                console.log("Error");
            }
            console.log(copyFromAddress)
            // range.copyFrom(copyFromAddress, );
        });
    } catch (error) {
        console.error(error);
    }
}

export async function pasteTranspose() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getActiveCell();
            const copyFromAddress = Office.context.document.settings.get("copyAddress");
            if(copyFromAddress === null ){
                console.log("Error");
            }
            console.log(copyFromAddress)
            range.copyFrom(copyFromAddress, "All", "true");
        });
    } catch (error) {
        console.error(error);
    }
}

export async function pasteInsert() {
    try {
        await Excel.run(async context => {
            
        });
    } catch (error) {
        console.error(error);
    }
}

export async function pasteDuplicatess() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getSelectedRange();
    let initialRange = range;
    range.format.fill.color = "yellow";
    range.load(["cellCount", "address"]);
    initialRange.load(["addressLocal", "address", "rowIndex"])
    await context.sync();
    const rowIndex = initialRange.rowIndex+1;

    if(range.cellCount === 1){
      let num = 0;
      let autoFillRange = null;
      //get rows above
      while (num < 10) {
        num++;
        range = range.getRowsAbove(1);
        range.load(["valueTypes", "values", "address"]);
        await context.sync();
        
        if(range.valueTypes[0][0] === "Double"){
          autoFillRange = range;
          break;
        }
      }

      if(autoFillRange === null){
        num=0;
        range = initialRange;
        while (num < 10) {
          num++;
          range = range.getRowsBelow(1);
          range.load(["valueTypes", "values", "address"]);
          await context.sync();

          if (range.valueTypes[0][0] === "Double") {
            autoFillRange = range;
            break;
          }
        }
      }
      

      let temp = range.getRangeEdge("Right").load("address");
      await context.sync();
      let addy = range.address.split("!")[1].split(/(\d+)/)[0] + rowIndex + ":" + temp.address.split("!")[1].split(/(\d+)/)[0] + rowIndex
      console.log(initialRange.addressLocal)
      console.log("ADDY:"+ addy)
      initialRange.autoFill(addy)
      console.log(temp.address)
      console.log(range.address + ":" + temp.address.split("!")[1])
      // console.log(range.values[0][0], range.address);
    }
        });
    } catch (error) {
        console.error(error);
    }
}