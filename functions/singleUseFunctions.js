export async function symbols(symbol) {
    try {
        await Excel.run(async context => {
            let range = context.workbook.getActiveCell();
            range.load("values");
            await context.sync();
            range.values = [[range.values + symbol]]
            // range.values = [[ ]]
        });
    } catch (error) {
        console.error(error);
    }
}

