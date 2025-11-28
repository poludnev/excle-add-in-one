type SummaryHeadings = [
  "Terms of delivery",
  "Total col",
  "FAKTURA",
  "Date",
  "KUPAC",
  "QUANTITY",
  "Total Netto",
  "Total Gross",
  "EXPORT TOTAL",
  "EXPORT CURRENCY",
  "IMPORT INVOICE NUM",
  "IMPORT TOTAL",
  "IMPORT CURRENCY",
  "SHIPPER INFO",

  "Shipper name",
  "Shipper address",
  "Shipper tax details",
  "SHIPPER REGISTRATON INFO",

  "Shipper reg",
  "Shipper tax",
  "Shipper IBAN",
  "CONSIGNEE INFO",

  "Consignee name",
  "Consignee address",
  "Consignee VAT",
];

const summaryHeadings: SummaryHeadings = [
  "Terms of delivery",
  "Total col",
  "FAKTURA",
  "Date",
  "KUPAC",
  "QUANTITY",
  "Total Netto",
  "Total Gross",
  "EXPORT TOTAL",
  "EXPORT CURRENCY",
  "IMPORT INVOICE NUM",
  "IMPORT TOTAL",
  "IMPORT CURRENCY",
  "SHIPPER INFO",
  "Shipper name",
  "Shipper address",
  "Shipper tax details",
  "SHIPPER REGISTRATON INFO",
  "Shipper reg",
  "Shipper tax",
  "Shipper IBAN",
  "CONSIGNEE INFO",
  "Consignee name",
  "Consignee address",
  "Consignee VAT",
];

const shipperDefaultValues = {
  name: "SAVIMPEX doo",
  address: "Gogoljeva 7, 21000 Novi Sad, Srbija/ Serbia",
  taxNumbers: "PIB:  113113438 Maticni broj : 21804495",
  regNumber: "Reg. number: 21804495",
  taxNumber: "Tax number: 113113438",
  iban: "Bank account: IBAN: RS35375111120000153237",
};

const consigneeDefaultValues = {
  name: "Limited Liability Company Ursus Trade",
  address: "Russia, Moscow, Hlebniy pereulok 19A, 121069",
  vat: "VAT ID: 7735189429",
};

export const insertSummaryHeaders = async () => {
  console.log("run summary headers");
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const targetWorksheet = worksheets.getItem("summary");
    const headingRange = targetWorksheet.getRangeByIndexes(0, 0, 25, 1);
    console.log("summaryHeading leng", summaryHeadings.length);

    const headingsToColumns = summaryHeadings.map((val) => [val]);

    headingRange.values = headingsToColumns;
  });
};

export const fillSummaryDefaultValues = async () =>
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const worksheet = worksheets.getItem("summary");

    const exportInvoiceRow = summaryHeadings.indexOf("FAKTURA");
    worksheet.getCell(exportInvoiceRow, 1).formulas = [["=data!B3"]];

    const clientNameRow = summaryHeadings.indexOf("KUPAC");
    worksheet.getCell(clientNameRow, 1).formulas = [["=data!C3"]];

    const quantityRow = summaryHeadings.indexOf("QUANTITY");
    worksheet.getCell(quantityRow, 1).formulas = [["=data!J1"]];

    const nettRow = summaryHeadings.indexOf("Total Netto");

    worksheet.getCell(nettRow, 1).formulas = [["=data!M1"]];

    const grossRow = summaryHeadings.indexOf("Total Gross");
    worksheet.getCell(grossRow, 1).formulas = [["=data!N1"]];

    const exportTotalRow = summaryHeadings.indexOf("EXPORT TOTAL");
    worksheet.getCell(exportTotalRow, 1).formulas = [["=data!P1"]];

    const exportCurrencyRow = summaryHeadings.indexOf("EXPORT CURRENCY");
    worksheet.getCell(exportCurrencyRow, 1).formulas = [["=data!Q3"]];

    const shipperNameRow = summaryHeadings.indexOf("Shipper name");

    const shipperAddressRow = summaryHeadings.indexOf("Shipper address");
    const shipperTaxDetailsRow = summaryHeadings.indexOf("Shipper tax details");

    worksheet.getCell(shipperNameRow, 1).values = [[shipperDefaultValues.name]];
    worksheet.getCell(shipperAddressRow, 1).values = [[shipperDefaultValues.address]];
    worksheet.getCell(shipperTaxDetailsRow, 1).values = [[shipperDefaultValues.taxNumbers]];

    const shipperRegRow = summaryHeadings.indexOf("Shipper reg");
    const shipperTaxNumRow = summaryHeadings.indexOf("Shipper tax");
    const shipperIbanRow = summaryHeadings.indexOf("Shipper IBAN");

    worksheet.getCell(shipperRegRow, 1).values = [[shipperDefaultValues.regNumber]];
    worksheet.getCell(shipperTaxNumRow, 1).values = [[shipperDefaultValues.taxNumber]];
    worksheet.getCell(shipperIbanRow, 1).values = [[shipperDefaultValues.iban]];

    const consigneeNameRow = summaryHeadings.indexOf("Consignee name");
    const consigneAddressRow = summaryHeadings.indexOf("Consignee address");
    const consigneeVatRow = summaryHeadings.indexOf("Consignee VAT");

    worksheet.getCell(consigneeNameRow, 1).values = [[consigneeDefaultValues.name]];
    worksheet.getCell(consigneAddressRow, 1).values = [[consigneeDefaultValues.address]];
    worksheet.getCell(consigneeVatRow, 1).values = [[consigneeDefaultValues.vat]];
  });
