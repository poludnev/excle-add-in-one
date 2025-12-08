export async function createSheetWithName(name: string) {
  return Excel.run(async (context: Excel.RequestContext) => {
    const workbook: Excel.Workbook = context.workbook;

    const createdWorksheet: Excel.Worksheet = workbook.worksheets.add(name);

    return { success: true, sheetName: name, worksheet: createdWorksheet };
  });
}

export function setCenter(range: Excel.Range) {
  range.format.verticalAlignment = Excel.VerticalAlignment.center;
  range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
}

export function setOuterBorders(range: Excel.Range) {
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
