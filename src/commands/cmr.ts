import { setAllBorders, setOuterBorders } from "../utilities/utils";

export const fillCMRTemplate = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const targetSheet: Excel.Worksheet = worksheets.getItem("cmr");

    targetSheet.getRange("1:2").format.rowHeight = 15;
    targetSheet.getRange("3:4").format.rowHeight = 9;
    targetSheet.getRange("5:5").format.rowHeight = 2.25;
    targetSheet.getRange("6:9").format.rowHeight = 10.5;
    targetSheet.getRange("10:11").format.rowHeight = 9;
    targetSheet.getRange("12:15").format.rowHeight = 10.5;
    targetSheet.getRange("16:17").format.rowHeight = 9;
    targetSheet.getRange("18:19").format.rowHeight = 15.5;
    targetSheet.getRange("20:21").format.rowHeight = 9;
    targetSheet.getRange("22:23").format.rowHeight = 12;
    targetSheet.getRange("24:25").format.rowHeight = 10.5;
    targetSheet.getRange("26:27").format.rowHeight = 21.75;
    targetSheet.getRange("28:29").format.rowHeight = 9.0;
    targetSheet.getRange("30:31").format.rowHeight = 6.7;
    targetSheet.getRange("32:36").format.rowHeight = 11.75;
    targetSheet.getRange("37:40").format.rowHeight = 9.75;
    targetSheet.getRange("41:42").format.rowHeight = 11.75;
    targetSheet.getRange("43:44").format.rowHeight = 10.5;
    targetSheet.getRange("45:60").format.rowHeight = 7.5;
    targetSheet.getRange("61:64").format.rowHeight = 10.5;
    targetSheet.getRange("65:66").format.rowHeight = 8.25;
    targetSheet.getRange("67:67").format.rowHeight = 3;
    targetSheet.getRange("68:69").format.rowHeight = 7.5;
    targetSheet.getRange("70:71").format.rowHeight = 8.25;
    targetSheet.getRange("72:75").format.rowHeight = 7.5;

    targetSheet.getRange("A:B").format.columnWidth = 7.5; // 0.83
    targetSheet.getRange("C:C").format.columnWidth = 18; // 2.71
    targetSheet.getRange("D:D").format.columnWidth = 36.75; // 6.29
    targetSheet.getRange("E:E").format.columnWidth = 25.5; // 4.14
    targetSheet.getRange("F:F").format.columnWidth = 8.25; // 0.92
    targetSheet.getRange("G:G").format.columnWidth = 14.25; // 2
    targetSheet.getRange("H:I").format.columnWidth = 4.5; // 0.5
    // targetSheet.getRange("I:I").format.rowHeight = 0.83;
    targetSheet.getRange("J:J").format.columnWidth = 56.25; // 10
    targetSheet.getRange("K:K").format.columnWidth = 3.75; // 0.42
    targetSheet.getRange("L:L").format.columnWidth = 14.25; // 2
    targetSheet.getRange("M:M").format.columnWidth = 22.5; // 3.57
    targetSheet.getRange("N:N").format.columnWidth = 16.5; // 2.43
    targetSheet.getRange("O:O").format.columnWidth = 12; // 1.57
    targetSheet.getRange("P:P").format.columnWidth = 14.25; // 2
    targetSheet.getRange("Q:Q").format.columnWidth = 15; // 2.14
    targetSheet.getRange("R:R").format.columnWidth = 45; // 7.86
    targetSheet.getRange("S:S").format.columnWidth = 14.25; // 2
    targetSheet.getRange("T:T").format.columnWidth = 18.75; // 2.86
    targetSheet.getRange("U:U").format.columnWidth = 19.5; // 3
    targetSheet.getRange("V:V").format.columnWidth = 14.25; // 2
    targetSheet.getRange("W:W").format.columnWidth = 18.75; // 2.86
    targetSheet.getRange("X:X").format.columnWidth = 28.5; // 4.71
    targetSheet.getRange("Y:Y").format.columnWidth = 14.25; // 2
    targetSheet.getRange("Z:Z").format.columnWidth = 18.75; // 2.86
    targetSheet.getRange("AA:AA").format.columnWidth = 27; // 4.43
    targetSheet.getRange("AB:AB").format.columnWidth = 5.25; // 0.58
    targetSheet.getRange("AC:AC").format.columnWidth = 7.5; // 0.83

    targetSheet.getRange("A1:B2").merge();
    targetSheet.getRange("A1").values = [[1]];
    targetSheet.getRange("A1").format.font.size = 18;
    targetSheet.getRange("A1").format.verticalAlignment = Excel.VerticalAlignment.center;
    targetSheet.getRange("A1").format.horizontalAlignment = Excel.HorizontalAlignment.center;

    targetSheet.getRange("C1:I1").merge();
    targetSheet.getRange("C1:C1").values = [["Exemplar für den Absender"]];
    targetSheet.getRange("C2:I2").merge();
    targetSheet.getRange("C2:C2").values = [["Copy for sender"]];

    targetSheet.getRange("A1:AC2").format.fill.color = "000000";

    targetSheet.getRange("A1:AC2").format.font.color = "ffffff";

    targetSheet.getRange("C3:C4").merge();
    targetSheet.getRange("C3:C3").values = [[1]];

    targetSheet.getRange("D3:P3").merge();
    targetSheet.getRange("D3:D3").values = [["Absender (Name, Adresse, Land)"]];

    targetSheet.getRange("D4:P4").merge();
    targetSheet.getRange("D4:D4").values = [["Sender (name, address, country)"]];

    targetSheet.getRange("R3:V3").merge();
    targetSheet.getRange("R3:V3").format.font.size = 6;
    targetSheet.getRange("R3:R3").values = [["INTERNATIONALER FRACHTBRIEF"]];
    targetSheet.getRange("R4:V4").merge();
    targetSheet.getRange("R4:V4").format.font.size = 6;
    targetSheet.getRange("R4:R4").values = [["INTERNATIONAL CONSIGNEMENT NOTE"]];

    targetSheet.getRange("X3:Z4").merge();
    targetSheet.getRange("X3:X3").values = [["E204-210"]];
    targetSheet.getRange("AA3:AA4").merge();
    targetSheet.getRange("AA3:AA3").values = [["CMR"]];

    targetSheet.getRange("C6:P6").merge();
    targetSheet.getRange("C7:P7").merge();
    targetSheet.getRange("C8:P8").merge();
    targetSheet.getRange("C9:P9").merge();

    targetSheet.getRange("R6:V9").merge();
    targetSheet.getRange("R6:V9").format.font.size = 5;
    targetSheet.getRange("R6:V9").format.verticalAlignment = Excel.VerticalAlignment.center;
    targetSheet.getRange("R6:V9").format.horizontalAlignment = Excel.HorizontalAlignment.left;
    targetSheet.getRange("R6:V9").format.wrapText = true;
    targetSheet.getRange("R6:R6").values = [
      [
        "Diese Beförderung unterliegt, unbeschadet anders lautender Bestimmungen, dem Übereinkommen über den Vertrag über den internationalen Güterkraftverkehr (CMR).",
      ],
    ];
    targetSheet.getRange("W6:AA9").merge();
    targetSheet.getRange("W6:AA9").format.font.size = 5;
    targetSheet.getRange("W6:AA9").format.wrapText = true;
    targetSheet.getRange("W6:AA9").format.verticalAlignment = Excel.VerticalAlignment.center;
    targetSheet.getRange("W6:AA9").format.horizontalAlignment = Excel.HorizontalAlignment.left;
    targetSheet.getRange("W6:W6").values = [
      [
        "This carriage is subject, notwithstanding any clause to the contrary, to the Convention on the Contract for the international Carriage of goods by road (CMR).",
      ],
    ];

    setOuterBorders(targetSheet.getRange("C3:P9"));
    setOuterBorders(targetSheet.getRange("Q3:AA9"));

    targetSheet.getRange("C75:AA75").merge();
    targetSheet.getRange("A75:AC75").format.fill.color = "000000";

    targetSheet.getRange("A75:AC75").format.font.color = "ffffff";
    targetSheet.getRange("c76:AA76").merge();
    targetSheet.getRange("A76:AC76").format.fill.color = "000000";
    targetSheet.getRange("A76:AC76").format.font.color = "ffffff";

    // targetSheet.getRange("A:B");

    // const column = targetSheet.getRange("AB:AB");

    // Load column width
    // column.load("width");

    targetSheet.getUsedRange().format.font.name = "Arial";

    await context.sync();

    // console.log(column.width);
  });
};
