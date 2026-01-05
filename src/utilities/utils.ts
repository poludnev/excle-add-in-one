export async function createSheetWithName(sheetName: string) {
  return Excel.run(async (context: Excel.RequestContext) => {
    const workbook: Excel.Workbook = context.workbook;

    const sheets = context.workbook.worksheets;
    const sheet = sheets.getItemOrNullObject(sheetName);
    sheet.load("isNullObject");
    const countResult = sheets.getCount(); // ClientResult<number>
    await context.sync();

    if (!sheet.isNullObject) {
      if (countResult.value > 1) {
        sheet.delete();
        await context.sync();
        console.log(`Deleted worksheet "${sheetName}".`);
      } else {
        console.log(`Cannot delete "${sheetName}" because it's the only worksheet.`);
      }
    } else {
      console.log(`Worksheet "${sheetName}" does not exist.`);
    }

    const createdWorksheet: Excel.Worksheet = workbook.worksheets.add(sheetName);

    await context.sync();

    return { success: true, sheetName: name, worksheet: createdWorksheet };
  });
}

export function setCenter(range: Excel.Range) {
  range.format.verticalAlignment = Excel.VerticalAlignment.center;
  range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
}

export const setOuterBorders = (range: Excel.Range) =>
  setOuterBordersWithWeight(range, Excel.BorderWeight.thin);

export const setThickOuterBorders = (range: Excel.Range) =>
  setOuterBordersWithWeight(range, Excel.BorderWeight.thick);
export function setOuterBordersWithWeight(
  range: Excel.Range,
  weight: Excel.BorderWeight = Excel.BorderWeight.thin
) {
  const topItem = range.format.borders.getItemAt(0);
  topItem.style = "Continuous";
  topItem.color = "000000";
  topItem.weight = weight;
  const bottomItem = range.format.borders.getItemAt(1);
  bottomItem.style = "Continuous";
  bottomItem.color = "000000";
  bottomItem.weight = weight;
  const leftItem = range.format.borders.getItemAt(2);
  leftItem.style = "Continuous";
  leftItem.color = "000000";
  leftItem.weight = weight;
  const rightItem = range.format.borders.getItemAt(3);
  rightItem.style = "Continuous";
  rightItem.color = "000000";
  rightItem.weight = weight;
}

export function setAllBorders(range: Excel.Range) {
  const topItem = range.format.borders.getItemAt(0);
  topItem.style = "Continuous";
  topItem.color = "000000";
  topItem.weight = "Thin";
  const bottomItem = range.format.borders.getItemAt(1);
  bottomItem.style = "Continuous";
  bottomItem.color = "000000";
  bottomItem.weight = "Thin";
  const leftItem = range.format.borders.getItemAt(2);
  leftItem.style = "Continuous";
  leftItem.color = "000000";
  leftItem.weight = "Thin";
  const rightItem = range.format.borders.getItemAt(3);
  rightItem.style = "Continuous";
  rightItem.color = "000000";
  rightItem.weight = "Thin";
  const verticalItem = range.format.borders.getItemAt(4);
  verticalItem.style = "Continuous";
  verticalItem.color = "000000";
  verticalItem.weight = "Thin";
  const horizontalItem = range.format.borders.getItemAt(5);
  horizontalItem.style = "Continuous";
  horizontalItem.color = "000000";
  horizontalItem.weight = "Thin";
}
