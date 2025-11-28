import { getColumnLetter } from "../utilities/helpers";
import { setAllBorders } from "../utilities/utils";

export const fillTansitData = async () => {
  await Excel.run(async (context: Excel.RequestContext) => {
    try {
      const { workbook } = context;
      const { worksheets } = workbook;

      const dataSheet = worksheets.getItem("data");
      const targetSheet = worksheets.getItem("transit");

      const usedRange = dataSheet.getUsedRange();
      usedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);

      await context.sync();

      const lastRow = usedRange.rowIndex + usedRange.rowCount;
      const lastColumn = usedRange.columnIndex + usedRange.columnCount;

      const endColumn = getColumnLetter(lastColumn); // A=65 in ASCII
      console.log("end col", endColumn, lastColumn);

      const startColumn = "B";
      const startRow = 3;

      const dynamicRange = dataSheet.getRange(`${startColumn}${startRow}:${endColumn}${lastRow}`);
      // const dynamicRange = sourceSheet.getRange(`${startColumn}${startRow}:${endColumn}${lastRow}`);
      dynamicRange.load(["values"]);

      await context.sync();
      // await context.sync();
      console.log(dynamicRange.values);

      const transitdata = dynamicRange.values.reduce((acc, val) => {
        const hs = val[4];
        const description = val[24].trim();

        if (!acc[hs]) {
          acc[hs] = {};
          // const description = val[24].trim();
          // console.log("desc", hs, description);
        }

        if (!acc[hs][description]) {
          acc[hs][description] = { pcs: 0, nett: 0, gross: 0, amount: 0 };
        }

        acc[hs][description].pcs += val[8];
        acc[hs][description].nett += val[11];
        acc[hs][description].gross += val[12];
        acc[hs][description].amount += val[14];

        // else {
        //   const description = val[24].trim();
        //   // console.log("desc", hs, description);
        //   acc[hs][description] = { pcs: 0, nett: 0, gross: 0, amount: 0 };
        // }
        return acc;
      }, {});

      console.log("transit data", transitdata);

      const transitdataRangeValues = [];

      for (const hs of Object.keys(transitdata)) {
        const descriptions = Object.keys(transitdata[hs]);

        for (const description of descriptions) {
          const n = transitdataRangeValues.length + 1;
          transitdataRangeValues.push([
            n,
            hs,
            description,
            transitdata[hs][description].pcs,
            transitdata[hs][description].nett,
            transitdata[hs][description].gross,
            transitdata[hs][description].amount,
          ]);
        }
      }

      console.log("transitdataRangeValues", transitdataRangeValues);

      transitdataRangeValues.sort((a, b) => {
        return a[0] - b[0];
      });

      const tartgetRange = targetSheet.getRangeByIndexes(1, 0, transitdataRangeValues.length, 7);

      tartgetRange.values = transitdataRangeValues;

      const targetHeadingRange = targetSheet.getRangeByIndexes(0, 0, 1, 7);
      targetHeadingRange.values = [
        ["nn", "HS", "Description", "Pieces", "Nett", "Gross", "Amount"],
      ];

      targetHeadingRange.format.fill.color = "E8E8E8";
      targetHeadingRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      targetSheet.getCell(transitdataRangeValues.length + 1, 2).values = [["Total:"]];
      targetSheet.getCell(transitdataRangeValues.length + 1, 3).formulas = [
        [`=sum(D${2}:D${transitdataRangeValues.length + 1})`],
      ];
      targetSheet.getCell(transitdataRangeValues.length + 1, 4).formulas = [
        [`=sum(E${2}:E${transitdataRangeValues.length + 1})`],
      ];
      targetSheet.getCell(transitdataRangeValues.length + 1, 5).formulas = [
        [`=sum(F${2}:F${transitdataRangeValues.length + 1})`],
      ];
      targetSheet.getCell(transitdataRangeValues.length + 1, 6).formulas = [
        [`=sum(G${2}:G${transitdataRangeValues.length + 1})`],
      ];

      targetSheet.getRangeByIndexes(
        0,
        0,
        transitdataRangeValues.length + 2,
        1
      ).format.horizontalAlignment = Excel.HorizontalAlignment.center;
      targetSheet.getRangeByIndexes(
        0,
        3,
        transitdataRangeValues.length + 2,
        1
      ).format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const decimalsRange = targetSheet.getRangeByIndexes(
        0,
        4,
        transitdataRangeValues.length + 2,
        3
      );

      decimalsRange.numberFormat = [['_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)']];

      targetSheet.getUsedRange().format.autofitColumns();
      setAllBorders(targetSheet.getUsedRange());
      await context.sync();
    } catch (e) {
      console.error("flll transit error", e);
    }
  });
};
