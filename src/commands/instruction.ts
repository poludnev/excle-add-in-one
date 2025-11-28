import { getColumnLetter } from "../utilities/helpers";
type InstructionHeadings = [
  "FAKTURA",
  "KUPAC",
  "No.",
  "NAME OF GOODS",
  "HS",
  // "Origin",
  "Origin ENG",
  "EAN",
  "KOLICINA",
  "Netto KG ",
  "Gross KG",
  "Total Netto",
  "Total Gross",
  "IZLAZNA RSD",
  "IZLAZNA UKUPNO RSD",
  "VALUTA",
  "Dobavljac",
  "SD",
  "Ulazna faktura",
  "ULAZ CENA EUR",
  "ULAZ CENA EUR UKUPNO ",
  "VALUTA",
  // "Generic SKU",
  // "Full SKU",
  // "Description",
  // "Description ENG",
];
const instructionHeadings: InstructionHeadings = [
  "FAKTURA",
  "KUPAC",
  "No.",
  "NAME OF GOODS",
  "HS",
  // "Origin",
  "Origin ENG",
  "EAN",
  "KOLICINA",
  "Netto KG ",
  "Gross KG",
  "Total Netto",
  "Total Gross",
  "IZLAZNA RSD",
  "IZLAZNA UKUPNO RSD",
  "VALUTA",
  "Dobavljac",
  "SD",
  "Ulazna faktura",
  "ULAZ CENA EUR",
  "ULAZ CENA EUR UKUPNO ",
  "VALUTA",
  // "Generic SKU",
  // "Full SKU",
  // "Description",
  // "Description ENG",
];
export async function fillInstuctionHeading() {
  console.log("fill header");
  await Excel.run(async (context: Excel.RequestContext) => {
    const startCell = "B2";
    const workbook = context.workbook;
    const targetWorksheet: Excel.Worksheet = workbook.worksheets.getItem("instruction");

    // const headingRange: Excel.Range = targetWorksheet.getRange(startCell);
    const endCell = getColumnLetter(instructionHeadings.length + 1) + "2";
    console.log(endCell);

    const fullHeaderRange: Excel.Range = targetWorksheet.getRange(`${startCell}:${endCell}`);
    fullHeaderRange.values = [instructionHeadings];

    fullHeaderRange.format.fill.color = "E8E8E8";

    await context.sync();
  });
}

export async function fillInstructionData() {
  await Excel.run(async (context: Excel.RequestContext) => {
    try {
      console.log("run fill");
      const workbook = context.workbook;
      const sourceSheet = workbook.worksheets.getItem("data");
      const targetSheet = workbook.worksheets.getItem("instruction");

      const usedRange = sourceSheet.getUsedRange();
      usedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
      await context.sync();

      console.log("run fill 2", usedRange);
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
      const modifiedArray = dynamicRange.values.map((value, index) => {
        const res = value.slice(0, -4);
        res.splice(5, 1);

        res[10] = `=${getColumnLetter(10 - 1)}${index + startRow}*${getColumnLetter(10)}${index + startRow}`;
        res[11] = `=${getColumnLetter(10 - 1)}${index + startRow}*${getColumnLetter(11)}${index + startRow}`;
        res[13] = `=${getColumnLetter(10 - 1)}${index + startRow}*${getColumnLetter(14)}${index + startRow}`;
        res[19] = `=${getColumnLetter(10 - 1)}${index + startRow}*${getColumnLetter(20)}${index + startRow}`;
        return res;
      });

      const targetDataStartRow = 2;

      console.log(modifiedArray);
      const targetRange = targetSheet.getRangeByIndexes(
        targetDataStartRow,
        1,
        modifiedArray.length,
        modifiedArray[0].length
      );
      targetRange.values = modifiedArray;

      const quantityTotalCell = targetSheet.getCell(0, 8);

      quantityTotalCell.formulas = [
        [
          `=SUM(${getColumnLetter(9)}${targetDataStartRow + 1}:${getColumnLetter(9)}${modifiedArray.length + targetDataStartRow})`,
        ],
      ];

      const totalNetCell = targetSheet.getCell(0, 11);
      totalNetCell.formulas = [
        [
          `=SUM(${getColumnLetter(12)}${targetDataStartRow + 1}:${getColumnLetter(12)}${modifiedArray.length + targetDataStartRow})`,
        ],
      ];

      const totalGrossCell = targetSheet.getCell(0, 12);
      totalGrossCell.formulas = [
        [
          `=SUM(${getColumnLetter(13)}${targetDataStartRow + 1}:${getColumnLetter(13)}${modifiedArray.length + targetDataStartRow})`,
        ],
      ];

      const totalExportAmount = targetSheet.getCell(0, 14);
      totalExportAmount.formulas = [
        [
          `=SUM(${getColumnLetter(15)}${targetDataStartRow + 1}:${getColumnLetter(15)}${modifiedArray.length + targetDataStartRow})`,
        ],
      ];

      const totalImportAmount = targetSheet.getCell(0, 20);
      totalImportAmount.formulas = [
        [
          `=SUM(${getColumnLetter(21)}${targetDataStartRow + 1}:${getColumnLetter(21)}${modifiedArray.length + targetDataStartRow})`,
        ],
      ];

      const accountingFormatRange1 = targetSheet.getRange(
        `${getColumnLetter(10)}:${getColumnLetter(15)}`
      );

      accountingFormatRange1.numberFormat = [
        ['_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'],
      ];
      const accountingFormatRange2 = targetSheet.getRange(
        `${getColumnLetter(20)}:${getColumnLetter(21)}`
      );
      accountingFormatRange2.numberFormat = [
        ['_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'],
      ];

      const EANcodeColumn = targetSheet.getRange(`${getColumnLetter(8)}:${getColumnLetter(8)}`);
      EANcodeColumn.numberFormat = [["#0"]];

      targetSheet.getUsedRange().format.autofitColumns();

      await context.sync();
    } catch (e) {
      console.error("error", e);
    }
  });
}
