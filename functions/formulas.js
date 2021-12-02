export async function fastFillRight() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getSelectedRange();
            let initialRange = range;
            range.load(["cellCount", "address"]);
            initialRange.load(["addressLocal", "address", "rowIndex"]);
            await context.sync();
            const rowIndex = initialRange.rowIndex + 1;
            let finalAddress = "";
            let hasSpacer = false;
            let iterations = 0;
            if (range.cellCount === 1) {
        
              //get rows above
              let num = 0;
              let autoFillRange = null;
              while (num < 10) {
                num++;
                range = range.getRowsAbove(1);
                range.load(["valueTypes", "values", "address"]);
                await context.sync();
        
                if (range.valueTypes[0][0] === "Double") {
                  autoFillRange = range;
                  break;
                }
               
              }
              
              //get rows below
              if (autoFillRange === null) {
                num = 0;
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
              const offSetRangeOne = range.getOffsetRange(0, 1).load("values");
              const offSetRangeTwo = range.getOffsetRange(0, 2).load("values");
              range.load("numberFormat")
              await context.sync();
              console.log(offSetRangeTwo.values[0][0], offSetRangeOne.values[0][0])
              //nospacer
              if (offSetRangeOne.values[0][0] !== "") {
                let temp = range.getRangeEdge("Right").load("address");
                await context.sync();
                finalAddress =
                  range.address.split("!")[1].split(/(\d+)/)[0] +
                  rowIndex +
                  ":" +
                  temp.address.split("!")[1].split(/(\d+)/)[0] +
                  rowIndex;
              }
              //spacer
              else if (offSetRangeTwo.values[0][0] != "") {
                hasSpacer = true
                for (let i = 2; i < 100; i += 2) {
                  iterations++;
                  let stepRange = range.getOffsetRange(0, i).load(["text"])
                  await context.sync();
                  console.log(stepRange.text[0][0])
                  if (stepRange.text[0][0] === null || stepRange.text[0][0] === "") {
                    let temp = stepRange.getOffsetRange(0, -2).load("address");
                    await context.sync();
                    finalAddress =
                      range.address.split("!")[1].split(/(\d+)/)[0] +
                      rowIndex +
                      ":" +
                      temp.address.split("!")[1].split(/(\d+)/)[0] +
                      rowIndex;
                    break;
                  }
                }
              }
              console.log(initialRange.addressLocal);
              console.log("finalAddress:" + finalAddress);
        
              if(autoFillRange !== null && !hasSpacer){
                initialRange.autoFill(finalAddress);
              }
              else if(autoFillRange !== null && hasSpacer){
                console.log("h")
                // Application.suspendScreenUpdatingUntilNextSync()
                initialRange.autoFill(finalAddress);
                for(let i = 1; i<iterations*2; i+=2){
                  let spacerRange = initialRange.getOffsetRange(0,i)
                  spacerRange.clear("All")
                }
              }
            }
        });
    } catch (error) {
        console.error(error);
    }
}

export async function fastFillDown() {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getSelectedRange();
            let initialRange = range;
            range.load(["cellCount", "address"]);
            initialRange.load(["addressLocal", "address"]);
            await context.sync();
            let finalAddress = "";
            let hasSpacer = false;
            let iterations = 0;
        
        
            if (range.cellCount === 1) {
              
              //get rows above
              let num = 0;
              let autoFillRange = null;
              while (num < 10) {
                num++;
                range = range.getColumnsBefore(1);
                range.load(["valueTypes", "values", "address"]);
                await context.sync();
        
                if (range.valueTypes[0][0] === "Double") {
                  autoFillRange = range;
                  break;
                }
        
              }
        
              //get rows below
              if (autoFillRange === null) {
                num = 0;
                range = initialRange;
                while (num < 10) {
                  num++;
                  range = range.getColumnsAfter(1);
                  range.load(["valueTypes", "values", "address"]);
                  await context.sync();
        
                  if (range.valueTypes[0][0] === "Double") {
                    autoFillRange = range;
                    break;
                  }
                }
              }
        
              console.log(autoFillRange)
              const offSetRangeOne = range.getOffsetRange(1, 0).load("values");
              const offSetRangeTwo = range.getOffsetRange(2, 0).load("values");
              // range.load("numberFormat")
              await context.sync();
              console.log(offSetRangeTwo.values[0][0], offSetRangeOne.values[0][0])
        
              //nospacer
              if (offSetRangeOne.values[0][0] !== "") {
                let temp = range.getRangeEdge("Down").load("address");
                await context.sync();
                finalAddress =initialRange.address.split("!")[1] + ":" + initialRange.address.split("!")[1].split(/(\d+)/)[0] + temp.address.split("!")[1].split(/(\d+)/)[1]
              }
              
              //spacer
              else if (offSetRangeTwo.values[0][0] != "") {
                hasSpacer = true
                for (let i = 2; i < 100; i += 2) {
                  iterations++;
                  let stepRange = range.getOffsetRange(i , 0).load(["text"])
                  await context.sync();
                  console.log(stepRange.text[0][0])
                  if (stepRange.text[0][0] === null || stepRange.text[0][0] === "") {
                    let temp = stepRange.getOffsetRange(-2, 0).load("address");
                    await context.sync();
                    finalAddress = initialRange.address.split("!")[1] + ":" + initialRange.address.split("!")[1].split(/(\d+)/)[0] + temp.address.split("!")[1].split(/(\d+)/)[1]
                    break;
                  }
                }
              }
        
              console.log("finalAddress:" + finalAddress);
        
              if (autoFillRange !== null && !hasSpacer) {
                initialRange.autoFill(finalAddress);
              }
              else if (autoFillRange !== null && hasSpacer) {
                console.log("h")
                // Application.suspendScreenUpdatingUntilNextSync()
                initialRange.autoFill(finalAddress);
                for (let i = 1; i < iterations * 2; i += 2) {
                  let spacerRange = initialRange.getOffsetRange(i, 0)
                  spacerRange.clear("All")
                }
              }
            }
        });
    } catch (error) {
        console.error(error);
    }
}