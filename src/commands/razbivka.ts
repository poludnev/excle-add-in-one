import { getColumnLetter } from "../utilities/helpers";
import { setAllBorders } from "../utilities/utils";

type RazbivkaHeading = [
  "No.",
  "NAME OF GOODS",
  "Generic SKU",
  "Full SKU",
  "HS",
  "Description",
  "Origin",
  "EAN",
  "PIECES",
  "Netto KG",
  "Gross KG",
  "Total Netto",
  "Total Gross",
  "Price",
  "Amount",
  "DESC RU",
];

const razbivkaHeading: RazbivkaHeading = [
  "No.",
  "NAME OF GOODS",
  "Generic SKU",
  "Full SKU",
  "HS",
  "Description",
  "Origin",
  "EAN",
  "PIECES",
  "Netto KG",
  "Gross KG",
  "Total Netto",
  "Total Gross",
  "Price",
  "Amount",
  "DESC RU",
];

export async function fillRazbivkaHeading() {
  await Excel.run(async (context: Excel.RequestContext) => {
    try {
      console.log("run fill razbi");
      const startCell = "A12";
      const endCell = getColumnLetter(razbivkaHeading.length) + "12";
      console.log(endCell);
      const targetWorkSheet = context.workbook.worksheets.getItem("razbivka");

      const fillHeadingRange: Excel.Range = targetWorkSheet.getRange(`${startCell}:${endCell}`);
      await context.sync();
      // console.log(fillHeadingRange.addressLocal);
      fillHeadingRange.values = [razbivkaHeading];

      fillHeadingRange.format.fill.color = "E8E8E8";
      fillHeadingRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
      await context.sync();
    } catch (e) {
      console.error("fill razbivka head", e);
    }
  });
}

export async function fillRazbivkaData() {
  await Excel.run(async (context: Excel.RequestContext) => {
    try {
      const workbook = context.workbook;
      const sourceSheet = workbook.worksheets.getItem("data");
      const targetSheet = workbook.worksheets.getItem("razbivka");
      const summarySheet = workbook.worksheets.getItem("summary");

      const usedRange = sourceSheet.getUsedRange();
      usedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
      await context.sync();

      const lastRow = usedRange.rowIndex + usedRange.rowCount;
      const lastColumn = usedRange.columnIndex + usedRange.columnCount;

      const endColumn = getColumnLetter(lastColumn); // A=65 in ASCII
      console.log("end col", endColumn, lastColumn);

      const startColumn = "B";
      const startRow = 3;

      const dynamicRange = sourceSheet.getRange(`${startColumn}${startRow}:${endColumn}${lastRow}`);
      // const dynamicRange = sourceSheet.getRange(`${startColumn}${startRow}:${endColumn}${lastRow}`);
      dynamicRange.load(["values"]);

      await context.sync();
      // await context.sync();
      console.log(dynamicRange);
      console.log("Data:", dynamicRange.values);

      const updatedData = dynamicRange.values.map((values, index) => {
        const nn = index + 1;
        const nameOfGood = values[3];
        const generic = values[22];
        const fullSKU = values[23];

        const hs = values[4];
        const description = values[25];
        const origin = values[6];
        const ean = values[7];
        const quantity = values[8];
        const nettPerPiece = values[9];
        const grossPerPiece = values[10];
        const nettPerEan = values[11];
        const grossPerEan = values[12];
        const price = values[13];
        const amountPerEan = values[14];
        const descriptionRus = values[24];

        return [
          nn,
          nameOfGood,
          generic,
          fullSKU,
          hs,
          description,
          origin,
          ean,
          quantity,
          nettPerPiece,
          grossPerPiece,
          nettPerEan,
          grossPerEan,
          price,
          amountPerEan,
          descriptionRus,
        ];
      });

      console.log("updated data", updatedData);

      // const targetStartCell = "A12";
      // const targetEndCell = getColumnLetter(razbivkaHeading.length) + "12";

      const razbivkaDataStartRow = 13;
      const razbivkaDataEndRow = razbivkaDataStartRow + updatedData.length - 1;

      const targetRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 1,
        0,
        updatedData.length,
        updatedData[0].length
      );

      targetRange.values = updatedData;

      const totalColCell = summarySheet.getRangeByIndexes(0, 1, 25, 1);
      totalColCell.load(["values"]);
      await context.sync();

      const totalQuantityRange = targetSheet.getRangeByIndexes(razbivkaDataEndRow + 1, 8, 2, 1);

      totalQuantityRange.formulas = [
        [
          `=SUM(${getColumnLetter(9)}${razbivkaDataStartRow}:${getColumnLetter(9)}${razbivkaDataEndRow})`,
        ],
        ["pieces"],
      ];

      totalQuantityRange.format.font.bold = true;
      totalQuantityRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      setAllBorders(totalQuantityRange);

      const totalNettRange = targetSheet.getRangeByIndexes(razbivkaDataEndRow + 1, 11, 2, 1);
      totalNettRange.values = [
        [
          `=SUM(${getColumnLetter(12)}${razbivkaDataStartRow}:${getColumnLetter(12)}${razbivkaDataEndRow})`,
        ],
        ["kg netto"],
      ];

      totalNettRange.format.font.bold = true;
      totalNettRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      setAllBorders(totalNettRange);

      const totalGrossRange = targetSheet.getRangeByIndexes(razbivkaDataEndRow + 1, 12, 2, 1);
      totalGrossRange.values = [
        [
          `=SUM(${getColumnLetter(13)}${razbivkaDataStartRow}:${getColumnLetter(13)}${razbivkaDataEndRow})`,
        ],
        ["kg gross"],
      ];

      totalGrossRange.format.font.bold = true;
      totalGrossRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      setAllBorders(totalGrossRange);
      const totalAmountRange = targetSheet.getRangeByIndexes(razbivkaDataEndRow + 1, 14, 2, 1);
      totalAmountRange.values = [
        [
          `=SUM(${getColumnLetter(15)}${razbivkaDataStartRow}:${getColumnLetter(15)}${razbivkaDataEndRow})`,
        ],
        [`Total ${totalColCell.values[9][0]}`],
      ];

      setAllBorders(totalAmountRange);

      totalAmountRange.format.font.bold = true;
      totalAmountRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const razbivkaGeneralDataRange = targetSheet.getRangeByIndexes(
        razbivkaDataEndRow + 4,
        1,
        3,
        2
      );

      console.log(totalColCell.values);
      razbivkaGeneralDataRange.values = [
        ["Total Gross Weight KG:", `=M${razbivkaDataEndRow + 2}`],
        ["Total coll.:", totalColCell.values[1][0]],
        ["Terms of delivery:", totalColCell.values[0][0]],
      ];

      const razbivkaShipperDataRange = targetSheet.getRangeByIndexes(0, 0, 10, 2);
      razbivkaShipperDataRange.values = [
        [totalColCell.values[14][0], null],
        [totalColCell.values[15][0], null],
        [totalColCell.values[16][0], null],
        [null, null],
        ["Order no.", totalColCell.values[2][0]],
        ["Date", totalColCell.values[3][0]],
        ["Invoiced to:", null],
        [totalColCell.values[22][0], null],
        [totalColCell.values[23][0], null],
        [totalColCell.values[24][0], null],
      ];

      const shipperNameRange = targetSheet.getRangeByIndexes(0, 0, 1, updatedData[0].length);
      const shipperAddressRange = targetSheet.getRangeByIndexes(1, 0, 1, updatedData[0].length);
      const shipperTaxNumRange = targetSheet.getRangeByIndexes(2, 0, 1, updatedData[0].length);

      const consigneeTitleRange = targetSheet.getRangeByIndexes(6, 0, 1, updatedData[0].length);
      const consigneeNameRange = targetSheet.getRangeByIndexes(7, 0, 1, updatedData[0].length);
      const consigneeAddressRange = targetSheet.getRangeByIndexes(8, 0, 1, updatedData[0].length);
      const consigneeTaxNumRange = targetSheet.getRangeByIndexes(9, 0, 1, updatedData[0].length);
      shipperNameRange.merge();
      shipperAddressRange.merge();
      shipperTaxNumRange.merge();
      consigneeTitleRange.merge();
      consigneeNameRange.merge();
      consigneeAddressRange.merge();
      consigneeTaxNumRange.merge();

      shipperNameRange.format.font.bold = true;
      shipperAddressRange.format.font.bold = true;
      shipperTaxNumRange.format.font.bold = true;

      consigneeNameRange.format.font.bold = true;
      consigneeAddressRange.format.font.bold = true;
      consigneeTaxNumRange.format.font.bold = true;

      targetSheet.getRange("B5").format.font.bold = true;
      targetSheet.getRange("B5").format.font.size = 14;

      // await context.sync();

      const piecesRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 1,
        8,
        updatedData.length,
        1
      );

      piecesRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const eanRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 1,
        7,
        updatedData.length,
        1
      );

      eanRange.numberFormat = [["0"]];
      eanRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const hsRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 1,
        4,
        updatedData.length,
        1
      );

      hsRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const originRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 1,
        6,
        updatedData.length,
        1
      );

      originRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const decimalsRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 1,
        9,
        updatedData.length + 2,
        6
      );

      decimalsRange.numberFormat = [['_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)']];

      targetSheet.getUsedRange().format.autofitColumns();

      const dataLineNumbersRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 1,
        0,
        updatedData.length,
        1
      );
      dataLineNumbersRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const razbivkaGeneralDataHeadingsRange = targetSheet.getRangeByIndexes(
        razbivkaDataEndRow + 4,
        1,
        3,
        1
      );
      razbivkaGeneralDataHeadingsRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;
      const razbivkaGeneralDataDataRange = targetSheet.getRangeByIndexes(
        razbivkaDataEndRow + 4,
        2,
        3,
        1
      );

      razbivkaGeneralDataDataRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
      razbivkaGeneralDataDataRange.format.font.bold = true;

      const dataWithHeadingRange = targetSheet.getRangeByIndexes(
        razbivkaDataStartRow - 2,
        0,
        updatedData.length + 1,
        16
      );

      setAllBorders(dataWithHeadingRange);

      await context.sync();
    } catch (e) {
      console.error("fill razbivka", e);
    }
  });
}
