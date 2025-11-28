import { getColumnLetter } from "../utilities/helpers";
import { setAllBorders } from "../utilities/utils";
import { fillRazbivkaData } from "./razbivka";

type PackingHeading = [
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
  // "Price",
  // "Amount",
  // "DESC RU",
];

const packingHeading: PackingHeading = [
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
  // "Price",
  // "Amount",
  // "DESC RU",
];

export async function fillPackingHeading() {
  await Excel.run(async (context: Excel.RequestContext) => {
    try {
      console.log("run packing heading");

      const startCell = "A12";
      const endCell = getColumnLetter(packingHeading.length) + "12";

      const targetWorkSheet = context.workbook.worksheets.getItem("packing");

      const fillHeadingRange: Excel.Range = targetWorkSheet.getRange(`${startCell}:${endCell}`);

      await context.sync();

      fillHeadingRange.values = [packingHeading];

      fillHeadingRange.format.fill.color = "E8E8E8";
      fillHeadingRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      await context.sync();
    } catch (e) {
      console.error("packing heading error", e);
    }
  });
}

export async function fillPackingData() {
  await Excel.run(async (context: Excel.RequestContext) => {
    try {
      const { workbook } = context;
      const { worksheets } = workbook;

      const targetSheet = worksheets.getItem("packing");
      const razbivkaSheet = worksheets.getItem("razbivka");
      const dataSheet = worksheets.getItem("data");
      const summarySheet = worksheets.getItem("summary");

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
      console.log(dynamicRange);
      console.log("Data:", dynamicRange.values);

      const targetRangeData = dynamicRange.values.map((values, index) => {
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
          // price,
          // amountPerEan,
          // descriptionRus,
        ];
      });

      console.log("target range data", targetRangeData);

      const dataRowsLength = targetRangeData.length;
      const dataColsLength = targetRangeData[0].length;

      // console.log("length", dataRowsLength, dataColsLength);

      const packingDataStartRow = 13;
      const packingDataEndRow = packingDataStartRow + dataRowsLength - 1;

      const targetRange = targetSheet.getRangeByIndexes(
        packingDataStartRow - 1,
        0,
        dataRowsLength,
        dataColsLength
      );

      targetRange.values = targetRangeData;

      const packingShipperDataRange = targetSheet.getRangeByIndexes(0, 0, 10, 2);
      packingShipperDataRange.values = [
        ["=razbivka!A1", null],
        ["=razbivka!A2", null],
        ["=razbivka!A3", null],
        [null, null],
        ["Order no.", "=razbivka!B5"],
        ["Date", "=razbivka!B6"],
        ["Invoiced to:", null],
        ["=razbivka!A8", null],
        ["=razbivka!A9", null],
        ["=razbivka!A10", null],
      ];

      const shipperNameRange = targetSheet.getRangeByIndexes(0, 0, 1, dataColsLength);
      const shipperAddressRange = targetSheet.getRangeByIndexes(1, 0, 1, dataColsLength);
      const shipperTaxNumRange = targetSheet.getRangeByIndexes(2, 0, 1, dataColsLength);

      const consigneeTitleRange = targetSheet.getRangeByIndexes(6, 0, 1, dataColsLength);
      const consigneeNameRange = targetSheet.getRangeByIndexes(7, 0, 1, dataColsLength);
      const consigneeAddressRange = targetSheet.getRangeByIndexes(8, 0, 1, dataColsLength);
      const consigneeTaxNumRange = targetSheet.getRangeByIndexes(9, 0, 1, dataColsLength);
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

      const totalQuantityRange = targetSheet.getRangeByIndexes(packingDataEndRow + 1, 8, 2, 1);

      totalQuantityRange.formulas = [
        [
          `=SUM(${getColumnLetter(9)}${packingDataStartRow}:${getColumnLetter(9)}${packingDataEndRow})`,
        ],
        ["pieces"],
      ];

      totalQuantityRange.format.font.bold = true;
      totalQuantityRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      setAllBorders(totalQuantityRange);

      const totalNettRange = targetSheet.getRangeByIndexes(packingDataEndRow + 1, 11, 2, 1);
      totalNettRange.values = [
        [
          `=SUM(${getColumnLetter(12)}${packingDataStartRow}:${getColumnLetter(12)}${packingDataEndRow})`,
        ],
        ["kg netto"],
      ];

      totalNettRange.format.font.bold = true;
      totalNettRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      setAllBorders(totalNettRange);

      const totalGrossRange = targetSheet.getRangeByIndexes(packingDataEndRow + 1, 12, 2, 1);
      totalGrossRange.values = [
        [
          `=SUM(${getColumnLetter(13)}${packingDataStartRow}:${getColumnLetter(13)}${packingDataEndRow})`,
        ],
        ["kg gross"],
      ];

      totalGrossRange.format.font.bold = true;
      totalGrossRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      setAllBorders(totalGrossRange);

      const packingGeneralDataRange = targetSheet.getRangeByIndexes(packingDataEndRow + 4, 1, 3, 2);

      packingGeneralDataRange.values = [
        ["Total Gross Weight KG:", `=M${packingDataEndRow + 2}`],
        ["Total coll.:", `=razbivka!C${packingDataEndRow + 4 + 2}`],
        ["Terms of delivery:", `=razbivka!C${packingDataEndRow + 4 + 3}`],
      ];

      const piecesRange = targetSheet.getRangeByIndexes(
        packingDataStartRow - 1,
        8,
        dataRowsLength,
        1
      );

      piecesRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const eanRange = targetSheet.getRangeByIndexes(packingDataStartRow - 1, 7, dataRowsLength, 1);

      eanRange.numberFormat = [["0"]];
      eanRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const hsRange = targetSheet.getRangeByIndexes(packingDataStartRow - 1, 4, dataRowsLength, 1);

      hsRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const originRange = targetSheet.getRangeByIndexes(
        packingDataStartRow - 1,
        6,
        dataRowsLength,
        1
      );

      originRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const dataLineNumbersRange = targetSheet.getRangeByIndexes(
        packingDataStartRow - 1,
        0,
        dataRowsLength,
        1
      );
      dataLineNumbersRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      const decimalsRange = targetSheet.getRangeByIndexes(
        packingDataStartRow - 1,
        9,
        dataRowsLength + 2,
        6
      );

      decimalsRange.numberFormat = [['_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)']];

      targetSheet.getUsedRange().format.autofitColumns();

      const razbivkaGeneralDataHeadingsRange = targetSheet.getRangeByIndexes(
        packingDataEndRow + 4,
        1,
        3,
        1
      );
      razbivkaGeneralDataHeadingsRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;
      const razbivkaGeneralDataDataRange = targetSheet.getRangeByIndexes(
        packingDataEndRow + 4,
        2,
        3,
        1
      );

      razbivkaGeneralDataDataRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
      razbivkaGeneralDataDataRange.format.font.bold = true;

      const dataWithHeadingRange = targetSheet.getRangeByIndexes(
        packingDataStartRow - 2,
        0,
        dataRowsLength + 1, //header
        dataColsLength
      );

      setAllBorders(dataWithHeadingRange);

      const targetSheetLayout = targetSheet.pageLayout;

      targetSheetLayout.orientation = Excel.PageOrientation.landscape;
      targetSheetLayout.paperSize = Excel.PaperType.a4;

      targetSheetLayout.zoom = { horizontalFitToPages: 1 };

      await context.sync();
    } catch (e) {
      console.error("fill packing data error", e);
    }
  });
}
