import { getColumnLetter } from "../utilities/helpers";

type SummaryHeadings = [
  "FAKTURA",
  "KUPAC",
  "No.",
  "NAME OF GOODS",
  "HS",
  "Origin",
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
  "Generic SKU",
  "Full SKU",
  "Description",
  "Description ENG",
];
const summaryHeadings: SummaryHeadings = [
  "FAKTURA",
  "KUPAC",
  "No.",
  "NAME OF GOODS",
  "HS",
  "Origin",
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
  "Generic SKU",
  "Full SKU",
  "Description",
  "Description ENG",
];

export async function createSheetWithName(name: string) {
  return Excel.run(async (context: Excel.RequestContext) => {
    try {
      const workbook: Excel.Workbook = context.workbook;
      console.log("Creating sheet with name:1", name);

      const createdWorksheet: Excel.Worksheet = workbook.worksheets.add(name);
      console.log("Creating sheet with name:2 ", name);

      createdWorksheet.activate();

      console.log("Creating sheet with name:3 ", name);
      await context.sync();
      console.log("Creating sheet with name:4 ", name);

      return { success: true, sheetName: name, worksheet: createdWorksheet };
    } catch (error) {
      if (
        error instanceof OfficeExtension.Error &&
        error.code === Excel.ErrorCodes.itemAlreadyExists
      ) {
        if (error.code === Excel.ErrorCodes.itemAlreadyExists) {
          console.warn(`Sheet with name "${name}" already exists. Activating existing sheet.`);
          const workbook: Excel.Workbook = context.workbook;
          const existingWorksheet: Excel.Worksheet = workbook.worksheets.getItem(name);
          existingWorksheet.activate();
          await context.sync();
          return {
            success: false,
            error: new Error(`Sheet with name "${name}" already exists.`),
            sheetName: name,
            worksheet: existingWorksheet,
          };
        }
      }
      console.error("Error creating sheet with name:", name, error);
      return {
        success: false,
        error: error instanceof Error ? error : new Error("Unknown error"),
      };
    }
  });
}

export async function fillDataHeading() {
  console.log("fill header");
  await Excel.run(async (context: Excel.RequestContext) => {
    const startCell = "B2";
    const workbook = context.workbook;
    const targetWorksheet: Excel.Worksheet = workbook.worksheets.getItem("data");

    // const headingRange: Excel.Range = targetWorksheet.getRange(startCell);
    const endCell = getColumnLetter(summaryHeadings.length + 1) + "2";
    console.log(endCell);

    const fullHeaderRange: Excel.Range = targetWorksheet.getRange(`${startCell}:${endCell}`);
    fullHeaderRange.values = [summaryHeadings];

    fullHeaderRange.format.fill.color = "E8E8E8";

    await context.sync();
  });
}

export const updateDataFormats = async () => {
  console.log("rus update formats");
  Excel.run(async (context: Excel.RequestContext) => {
    const { workbook } = context;
    const { worksheets } = workbook;

    const targetWorksheet = worksheets.getItem("data");

    const usedRange = targetWorksheet.getUsedRange();
    usedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
    console.log(usedRange);

    await context.sync();

    const lastRow = usedRange.rowIndex + usedRange.rowCount;

    targetWorksheet.getRange("J1").formulas = [[`=sum(J${3}:J${lastRow})`]];
    targetWorksheet.getRange("M1").formulas = [[`=sum(M${3}:M${lastRow})`]];
    targetWorksheet.getRange("N1").formulas = [[`=sum(N${3}:N${lastRow})`]];
    targetWorksheet.getRange("P1").formulas = [[`=sum(P${3}:P${lastRow})`]];
    targetWorksheet.getRange("V1").formulas = [[`=sum(V${3}:V${lastRow})`]];

    usedRange.format.autofitColumns();
  });
};
