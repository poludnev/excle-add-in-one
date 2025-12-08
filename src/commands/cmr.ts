import { setAllBorders, setCenter, setOuterBorders } from "../utilities/utils";

const formatHeights = (targetSheet: Excel.Worksheet) => {
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
};

const mergeRanges = (targetSheet: Excel.Worksheet) => {
  targetSheet.getRange("A1:B2").merge();
  targetSheet.getRange("C1:I1").merge();
  targetSheet.getRange("C2:I2").merge();
  targetSheet.getRange("C3:C4").merge();
  targetSheet.getRange("D3:P3").merge();
  targetSheet.getRange("D4:P4").merge();

  targetSheet.getRange("R3:V3").merge();

  targetSheet.getRange("R4:V4").merge();
  targetSheet.getRange("X3:Z4").merge();
  targetSheet.getRange("AA3:AA4").merge();

  targetSheet.getRange("C6:P6").merge();
  targetSheet.getRange("C7:P7").merge();
  targetSheet.getRange("C8:P8").merge();
  targetSheet.getRange("C9:P9").merge();

  targetSheet.getRange("R6:V9").merge();

  targetSheet.getRange("W6:AA9").merge();

  // 2
  targetSheet.getRange("C10:C11").merge();
  targetSheet.getRange("D10:P10").merge();
  targetSheet.getRange("D11:P11").merge();

  targetSheet.getRange("C12:P12").merge();
  targetSheet.getRange("C13:P13").merge();
  targetSheet.getRange("C14:P14").merge();
  targetSheet.getRange("C15:P15").merge();

  // 3
  targetSheet.getRange("C16:C17").merge();
  targetSheet.getRange("D16:P16").merge();
  targetSheet.getRange("D17:P17").merge();

  targetSheet.getRange("C18:P19").merge();

  // 4
  targetSheet.getRange("C20:C21").merge();
  targetSheet.getRange("D20:P20").merge();
  targetSheet.getRange("D21:P21").merge();
  targetSheet.getRange("C22:P23").merge();

  // 5
  targetSheet.getRange("C24:C25").merge();

  targetSheet.getRange("D24:P24").merge();
  targetSheet.getRange("D25:P25").merge();

  targetSheet.getRange("C26:P27").merge();

  // 6

  targetSheet.getRange("C28:C29").merge();

  targetSheet.getRange("D28:F28").merge();
  targetSheet.getRange("D29:F29").merge();

  targetSheet.getRange("C30:F40").merge();

  targetSheet.getRange("D41:D42").merge();

  targetSheet.getRange("E41:F41").merge();
  targetSheet.getRange("E42:F42").merge();

  // 7
  targetSheet.getRange("G28:G29").merge();

  targetSheet.getRange("H28:K28").merge();
  targetSheet.getRange("H29:K29").merge();

  targetSheet.getRange("G30:K40").merge();

  // 8
  targetSheet.getRange("L28:L29").merge();
  targetSheet.getRange("M28:O28").merge();
  targetSheet.getRange("M29:O29").merge();

  targetSheet.getRange("L30:O40").merge();
  targetSheet.getRange("L41:O42").merge();

  // 9
  targetSheet.getRange("P28:P29").merge();

  targetSheet.getRange("Q28:R28").merge();
  targetSheet.getRange("Q29:R29").merge();

  targetSheet.getRange("P30:R40").merge();

  targetSheet.getRange("P41:Q41").merge();
  targetSheet.getRange("P42:Q42").merge();

  // 10
  targetSheet.getRange("S28:S29").merge();

  targetSheet.getRange("T28:U28").merge();

  targetSheet.getRange("T29:U29").merge();

  targetSheet.getRange("S30:U42").merge();

  // 11
  targetSheet.getRange("V28:V29").merge();

  targetSheet.getRange("W28:X28").merge();

  targetSheet.getRange("W29:X29").merge();

  targetSheet.getRange("V30:X40").merge();
  targetSheet.getRange("V41:x41").merge();
  targetSheet.getRange("V42:x42").merge();

  // 12
  targetSheet.getRange("Y28:Y29").merge();

  targetSheet.getRange("Z28:AA28").merge();

  targetSheet.getRange("Z29:AA29").merge();

  targetSheet.getRange("Y30:AA42").merge();

  //13
  targetSheet.getRangeByIndexes(42, 2, 2, 1).merge();
  targetSheet.getRangeByIndexes(42, 3, 1, 13).merge();
  targetSheet.getRangeByIndexes(43, 3, 1, 13).merge();
  targetSheet.getRangeByIndexes(44, 2, 14, 14).merge();

  // 14
  targetSheet.getRangeByIndexes(58, 2, 2, 1).merge();

  targetSheet.getRangeByIndexes(58, 3, 1, 2).merge();
  targetSheet.getRangeByIndexes(59, 3, 1, 2).merge();

  targetSheet.getRangeByIndexes(58, 5, 2, 22).merge();

  // 15

  targetSheet.getRangeByIndexes(60, 2, 2, 1).merge();

  targetSheet.getRangeByIndexes(60, 3, 1, 13).merge();
  targetSheet.getRangeByIndexes(61, 3, 1, 13).merge();

  targetSheet.getRangeByIndexes(62, 2, 1, 3).merge();
  targetSheet.getRangeByIndexes(63, 2, 1, 3).merge();

  targetSheet.getRangeByIndexes(62, 5, 1, 11).merge();
  targetSheet.getRangeByIndexes(63, 5, 1, 11).merge();

  // 16
  targetSheet.getRange("Q10:Q11").merge();
  // targetSheet.getRange("Q10:Q10").values = [["16"]];

  targetSheet.getRange("R10:AA10").merge();
  targetSheet.getRange("R11:AA11").merge();
  targetSheet.getRange("Q12:AA15").merge();

  // 17
  targetSheet.getRange("Q16:Q17").merge();
  targetSheet.getRange("R16:AA16").merge();
  targetSheet.getRange("R17:AA17").merge();
  targetSheet.getRange("Q18:AA19").merge();

  // 18
  targetSheet.getRange("Q20:Q21").merge();
  targetSheet.getRange("R20:AA20").merge();
  targetSheet.getRange("R21:AA21").merge();
  targetSheet.getRange("Q22:AA27").merge();

  // 19
  targetSheet.getRangeByIndexes(42, 16, 2, 1).merge();

  targetSheet.getRangeByIndexes(42, 18, 1, 3).merge();
  targetSheet.getRangeByIndexes(43, 18, 1, 3).merge();

  targetSheet.getRangeByIndexes(42, 21, 1, 3).merge();
  targetSheet.getRangeByIndexes(43, 21, 1, 3).merge();

  targetSheet.getRangeByIndexes(42, 24, 1, 3).merge();
  targetSheet.getRangeByIndexes(43, 24, 1, 3).merge();

  targetSheet.getRangeByIndexes(44, 16, 1, 2).merge();
  targetSheet.getRangeByIndexes(45, 16, 1, 2).merge();

  targetSheet.getRangeByIndexes(44, 18, 2, 3).merge();
  targetSheet.getRangeByIndexes(44, 21, 2, 2).merge();
  targetSheet.getRangeByIndexes(44, 23, 2, 1).merge();
  targetSheet.getRangeByIndexes(44, 24, 2, 3).merge();

  targetSheet.getRangeByIndexes(46, 16, 1, 2).merge();
  targetSheet.getRangeByIndexes(47, 16, 1, 2).merge();

  targetSheet.getRangeByIndexes(46, 18, 2, 3).merge();
  targetSheet.getRangeByIndexes(46, 21, 2, 2).merge();
  targetSheet.getRangeByIndexes(46, 23, 2, 1).merge();
  targetSheet.getRangeByIndexes(46, 24, 2, 3).merge();

  targetSheet.getRangeByIndexes(48, 16, 1, 2).merge();
  targetSheet.getRangeByIndexes(49, 16, 1, 2).merge();

  targetSheet.getRangeByIndexes(48, 18, 2, 3).merge();
  targetSheet.getRangeByIndexes(48, 21, 2, 2).merge();
  targetSheet.getRangeByIndexes(48, 23, 2, 1).merge();
  targetSheet.getRangeByIndexes(48, 24, 2, 3).merge();

  targetSheet.getRangeByIndexes(50, 16, 1, 2).merge();
  targetSheet.getRangeByIndexes(51, 16, 1, 2).merge();

  targetSheet.getRangeByIndexes(50, 18, 2, 3).merge();
  targetSheet.getRangeByIndexes(50, 21, 2, 2).merge();
  targetSheet.getRangeByIndexes(50, 23, 2, 1).merge();
  targetSheet.getRangeByIndexes(50, 24, 2, 3).merge();

  targetSheet.getRangeByIndexes(52, 16, 1, 2).merge();
  targetSheet.getRangeByIndexes(53, 16, 1, 2).merge();

  targetSheet.getRangeByIndexes(52, 18, 2, 3).merge();
  targetSheet.getRangeByIndexes(52, 21, 2, 2).merge();
  targetSheet.getRangeByIndexes(52, 23, 2, 1).merge();
  targetSheet.getRangeByIndexes(52, 24, 2, 3).merge();

  targetSheet.getRangeByIndexes(54, 16, 1, 2).merge();
  targetSheet.getRangeByIndexes(55, 16, 1, 2).merge();

  targetSheet.getRangeByIndexes(54, 18, 2, 3).merge();
  targetSheet.getRangeByIndexes(54, 21, 2, 2).merge();
  targetSheet.getRangeByIndexes(54, 23, 2, 1).merge();
  targetSheet.getRangeByIndexes(54, 24, 2, 3).merge();

  targetSheet.getRangeByIndexes(56, 16, 1, 2).merge();
  targetSheet.getRangeByIndexes(57, 16, 1, 2).merge();

  targetSheet.getRangeByIndexes(56, 18, 2, 3).merge();
  targetSheet.getRangeByIndexes(56, 21, 2, 2).merge();
  targetSheet.getRangeByIndexes(56, 23, 2, 1).merge();
  targetSheet.getRangeByIndexes(56, 24, 2, 3).merge();

  //20
  targetSheet.getRangeByIndexes(60, 16, 2, 1).merge();

  targetSheet.getRangeByIndexes(60, 17, 1, 10).merge();
  targetSheet.getRangeByIndexes(61, 17, 1, 10).merge();

  targetSheet.getRangeByIndexes(62, 16, 4, 11).merge();

  // 21
  targetSheet.getRangeByIndexes(64, 2, 2, 1).merge();

  targetSheet.getRangeByIndexes(64, 3, 1, 2).merge();
  targetSheet.getRangeByIndexes(65, 3, 1, 2).merge();

  targetSheet.getRangeByIndexes(64, 5, 2, 6).merge();

  targetSheet.getRangeByIndexes(64, 12, 2, 4).merge();

  // 22

  targetSheet.getRangeByIndexes(67, 2, 2, 1).merge();

  targetSheet.getRangeByIndexes(69, 2, 5, 1).merge();

  targetSheet.getRangeByIndexes(67, 3, 5, 8).merge();
  targetSheet.getRangeByIndexes(72, 3, 1, 8).merge();
  targetSheet.getRangeByIndexes(73, 3, 1, 8).merge();

  // 23
  targetSheet.getRangeByIndexes(67, 11, 2, 1).merge();
  targetSheet.getRangeByIndexes(69, 11, 5, 1).merge();

  targetSheet.getRangeByIndexes(67, 12, 5, 7).merge();

  targetSheet.getRangeByIndexes(72, 12, 1, 7).merge();
  targetSheet.getRangeByIndexes(73, 12, 1, 7).merge();

  // 24
  targetSheet.getRangeByIndexes(67, 19, 2, 1).merge();
  targetSheet.getRangeByIndexes(69, 19, 5, 1).merge();

  targetSheet.getRangeByIndexes(67, 20, 1, 3).merge();
  targetSheet.getRangeByIndexes(68, 20, 1, 3).merge();
  targetSheet.getRangeByIndexes(67, 23, 2, 4).merge();

  targetSheet.getRangeByIndexes(69, 21, 2, 2).merge();

  targetSheet.getRangeByIndexes(69, 24, 2, 3).merge();

  targetSheet.getRangeByIndexes(71, 20, 1, 7).merge();
  targetSheet.getRangeByIndexes(72, 20, 1, 7).merge();
  targetSheet.getRangeByIndexes(73, 20, 1, 7).merge();

  targetSheet.getRange("C75:AA75").merge();
  targetSheet.getRange("c76:AA76").merge();

  targetSheet.getRangeByIndexes(4, 0, 27, 1).merge();
  targetSheet.getRangeByIndexes(4, 1, 27, 1).merge();

  targetSheet.getRangeByIndexes(31, 0, 5, 2).merge();

  targetSheet.getRangeByIndexes(36, 0, 3, 1).merge();
  targetSheet.getRangeByIndexes(36, 1, 3, 1).merge();

  targetSheet.getRangeByIndexes(39, 0, 2, 2).merge();

  targetSheet.getRangeByIndexes(41, 0, 15, 1).merge();
  targetSheet.getRangeByIndexes(41, 1, 15, 1).merge();
};

const textFormats = (targetSheet: Excel.Worksheet) => {
  targetSheet.getRange("A1").format.font.size = 16;
  targetSheet.getRange("A1").format.verticalAlignment = Excel.VerticalAlignment.center;
  targetSheet.getRange("A1").format.horizontalAlignment = Excel.HorizontalAlignment.center;

  targetSheet.getRange("C1:D2").format.font.size = 8;
  targetSheet.getRange("A1:AC2").format.fill.color = "000000";
  targetSheet.getRange("A1:AC2").format.font.color = "ffffff";
  targetSheet.getRange("A75:AC76").format.fill.color = "000000";
  targetSheet.getRange("A75:AC76").format.font.color = "ffffff";

  const c3 = targetSheet.getRange("C3");
  setCenter(c3);

  const c10 = targetSheet.getRange("C10");
  setCenter(c10);

  const c16 = targetSheet.getRange("C16");
  setCenter(c16);

  const c20 = targetSheet.getRange("C20");
  setCenter(c20);

  const c24 = targetSheet.getRange("C24");
  setCenter(c24);
};

const fillInTemplateData = (targetSheet: Excel.Worksheet) => {
  targetSheet.getRange("A1").values = [["1"]];
  targetSheet.getRange("C1:C1").values = [["Exemplar für den Absender"]];
  targetSheet.getRange("C2:C2").values = [["Copy for sender"]];

  targetSheet.getRange("C3:C3").values = [["1"]];

  targetSheet.getRange("D3:D3").values = [["Absender (Name, Adresse, Land)"]];

  targetSheet.getRange("D4:D4").values = [["Sender (name, address, country)"]];

  targetSheet.getRange("R3:R3").values = [["INTERNATIONALER FRACHTBRIEF"]];
  targetSheet.getRange("R4:R4").values = [["INTERNATIONAL CONSIGNEMENT NOTE"]];

  targetSheet.getRange("X3:X3").values = [["E204-210"]];
  targetSheet.getRange("AA3:AA3").values = [["CMR"]];

  targetSheet.getRange("R6:R6").values = [
    [
      "Diese Beförderung unterliegt, unbeschadet anders lautender Bestimmungen, dem Übereinkommen über den Vertrag über den internationalen Güterkraftverkehr (CMR).",
    ],
  ];
  targetSheet.getRange("W6:W6").values = [
    [
      "This carriage is subject, notwithstanding any clause to the contrary, to the Convention on the Contract for the international Carriage of goods by road (CMR).",
    ],
  ];

  // 2
  targetSheet.getRange("C10:C10").values = [["2"]];
  targetSheet.getRange("D10:D10").values = [["Empfänger (Name, Adresse, Land)"]];
  targetSheet.getRange("D11:D11").values = [["Consignee (name, address, country)"]];

  // 3
  targetSheet.getRange("C16:C16").values = [["3"]];
  targetSheet.getRange("D16:D16").values = [["Auslieferort des Gutes (Ort, Land)"]];
  targetSheet.getRange("D17:D17").values = [["Place of delivery of the goods (place, country)"]];

  // 4
  targetSheet.getRange("C20:C20").values = [["4"]];
  targetSheet.getCell(19, 3).values = [
    ["Ort und Datum der Übernahme des Gutes (Ort, Land, Datum)"],
  ];
  targetSheet.getCell(20, 3).values = [
    ["Place and date of taking over the goods (place, country, date)"],
  ];
  targetSheet.getCell(21, 2).values = [["Beograd, 11.03.2025"]];

  // 5
  targetSheet.getRange("C24:C24").values = [["5"]];
  targetSheet.getCell(23, 3).values = [["Beigefügte Dokumente"]];
  targetSheet.getCell(24, 3).values = [["Documents attached"]];
  targetSheet.getCell(25, 2).values = [["Invoice № UT-EX-91-TE  dated 10.03.2025"]];

  // 6

  targetSheet.getRange("C28:C28").values = [["6"]];

  targetSheet.getRange("D28:D28").values = [["Kennzeichen u. Nummern"]];
  targetSheet.getRange("D29:D29").values = [["Marks and Nos"]];

  targetSheet.getRange("C41:C41").values = [["UN-Nr."]];

  targetSheet.getRange("E41:E41").values = [["Ben. s. Nr. 9"]];
  targetSheet.getRange("E42:E42").values = [["name s. nr. 9"]];

  // 7
  targetSheet.getRange("G28:G28").values = [["7"]];

  targetSheet.getRange("H28:H28").values = [["Anzahl der Pakete"]];
  targetSheet.getRange("H29:H29").values = [["Number of pakages"]];

  targetSheet.getRange("G30:G30").values = [["3 PAL"]];

  targetSheet.getRange("J41:J41").values = [["Gefahrzettelmuster-Nr."]];
  targetSheet.getRange("J42:J42").values = [["Hazard label sample no."]];

  // 8
  targetSheet.getRange("L28:L28").values = [["8"]];
  targetSheet.getRange("M28:M28").values = [["Art der Verpackung"]];
  targetSheet.getRange("M29:M29").values = [["Method of packing"]];

  targetSheet.getRange("L30:L30").values = [["Total 3 coll"]];

  // 9
  targetSheet.getRange("P28:P28").values = [["9"]];

  targetSheet.getRange("Q28:Q28").values = [["Bezeichnung des Gutes*"]];
  targetSheet.getRange("Q29:Q29").values = [["Nature of the goods*"]];

  targetSheet.getRange("P30:P30").values = [["Consumer electronic goods"]];

  targetSheet.getRange("P41:P41").values = [["Verp.-Grp."]];
  targetSheet.getRange("P42:P42").values = [["Pack. group"]];

  // 10
  targetSheet.getRange("S28:S28").values = [["10"]];

  targetSheet.getRange("T28:T28").values = [["Statistiknr."]];

  targetSheet.getRange("T29:T29").values = [["Statistical nr."]];

  targetSheet.getRange("S30:S30").values = [["See packing list"]];

  // 11
  targetSheet.getRange("V28:V28").values = [["11"]];

  targetSheet.getRange("W28:W28").values = [["Bruttogew. kg"]];

  targetSheet.getRange("W29:W29").values = [["Gross weight kg"]];

  targetSheet.getRange("V30:V30").values = [["285.50 KG"]];
  targetSheet.getRange("V41:V41").values = [["Total Gross:"]];

  targetSheet.getRange("V42:V42").values = [["10285.50 KG"]];

  // 12
  targetSheet.getRange("Y28:Y28").values = [["12"]];

  targetSheet.getRange("Z28:Z28").values = [["Volumen in m3"]];

  targetSheet.getRange("Z29:Z29").values = [["Volume in m3"]];

  //13

  targetSheet.getCell(42, 2).values = [["13"]];
  targetSheet.getCell(42, 3).values = [
    ["Anweisungen des Absenders (Zoll-, amtl. Behandlungen, Sondervorschriften, etc.)"],
  ];
  targetSheet.getCell(43, 3).values = [["Sender's instructions"]];
  targetSheet.getCell(44, 2).values = [
    [
      'T/P "Akulovskiy" Code 10013010 SVH OOO "Crocus Interservice" 143002 Moskovskaya Obl, Odintsovskiy r-n, S. Akulovo, Ul. Novaya, D. 137      Lic. 10013/200111/10118/11 from 10.10.2024',
    ],
  ];

  // 14
  targetSheet.getCell(58, 2).values = [["14"]];

  targetSheet.getCell(58, 3).values = [["Rückerstattung"]];
  targetSheet.getCell(59, 3).values = [["Cash on delivery"]];
  targetSheet.getCell(58, 5).values = [["DAP MOSCOW"]];

  // 15

  targetSheet.getCell(60, 2).values = [["15"]];

  targetSheet.getCell(60, 3).values = [["Frachtzahlungsanweisungen"]];
  targetSheet.getCell(61, 3).values = [["Instruction as to payement carriage"]];

  targetSheet.getCell(62, 2).values = [["Frei/Carriage paid"]];
  targetSheet.getCell(63, 2).values = [["Unfrei/Carriage forward"]];

  // 16

  targetSheet.getCell(9, 16).values = [["16"]];
  targetSheet.getCell(9, 17).values = [["Frachtführer (Name, Adresse, Land)"]];
  targetSheet.getCell(10, 17).values = [["Carrier (name, address, country)"]];
  targetSheet.getCell(11, 16).values = [["BG 2699-OI"]];

  // 17
  targetSheet.getCell(15, 16).values = [["17"]];
  targetSheet.getCell(15, 17).values = [["Nachfolgender Frachtführer (Name, Adresse, Land)"]];
  targetSheet.getCell(16, 17).values = [["Successive carriers (name, address, country)"]];
  targetSheet.getCell(17, 16).values = [["empty"]];

  // 18
  targetSheet.getCell(19, 16).values = [["18"]];
  targetSheet.getCell(19, 17).values = [["Vorbehalte und Bemerkungen der Frachtführer"]];
  targetSheet.getCell(20, 17).values = [["Carrier's reservations and observations"]];
  targetSheet.getCell(21, 16).values = [["em[ty"]];

  // 19
  targetSheet.getCell(42, 16).values = [["19"]];

  targetSheet.getCell(42, 17).values = [["Zu bezahlen vom"]];
  targetSheet.getCell(43, 17).values = [["To be paid by"]];

  targetSheet.getCell(42, 18).values = [["Absender"]];
  targetSheet.getCell(43, 18).values = [["Sender"]];

  targetSheet.getCell(42, 21).values = [["Währung"]];
  targetSheet.getCell(43, 21).values = [["Currency"]];
  targetSheet.getCell(42, 24).values = [["Empfänger"]];
  targetSheet.getCell(43, 24).values = [["Consignee"]];

  targetSheet.getCell(44, 16).values = [["Fracht"]];
  targetSheet.getCell(45, 16).values = [["Carriage"]];
  targetSheet.getCell(46, 16).values = [["Ermäßigung"]];
  targetSheet.getCell(47, 16).values = [["Reductions"]];
  targetSheet.getCell(48, 16).values = [["Zwischensumme"]];
  targetSheet.getCell(49, 16).values = [["Balance"]];
  targetSheet.getCell(50, 16).values = [["Zuschläge"]];
  targetSheet.getCell(51, 16).values = [["Supplement charges"]];
  targetSheet.getCell(52, 16).values = [["Nebengebühren"]];
  targetSheet.getCell(53, 16).values = [["Additional charges"]];
  targetSheet.getCell(54, 16).values = [["Sonstiges"]];
  targetSheet.getCell(55, 16).values = [["Miscellaneous"]];
  targetSheet.getCell(56, 16).values = [["Gesamtbetrag"]];
  targetSheet.getCell(57, 16).values = [["Total to be paid"]];

  //20
  targetSheet.getCell(60, 16).values = [["20"]];

  targetSheet.getCell(60, 17).values = [["Besondere Vereinbarungen"]];
  targetSheet.getCell(61, 17).values = [["Special agreements"]];

  // 21
  targetSheet.getCell(64, 2).values = [["21"]];

  targetSheet.getCell(64, 3).values = [["Ausgefertigt in"]];
  targetSheet.getCell(65, 3).values = [["Established in"]];

  targetSheet.getCell(64, 5).values = [["Belgrade"]];

  targetSheet.getCell(64, 11).values = [["am"]];
  targetSheet.getCell(65, 11).values = [["on"]];

  targetSheet.getCell(64, 12).values = [["11.03.2025"]];

  // 22

  targetSheet.getCell(67, 2).values = [["22"]];

  targetSheet.getCell(67, 3).values = [
    [
      "SAVIMPEX DOO                                          Novi Sad, Gogoljeva 7, Republic of Serbia       PIB: 113113438",
    ],
  ];

  targetSheet.getCell(72, 3).values = [["Signatur und Stempel des Absenders"]];
  targetSheet.getCell(73, 3).values = [["Signature and stamp of the sender"]];

  // 23
  targetSheet.getCell(67, 11).values = [["23"]];

  targetSheet.getCell(67, 12).values = [
    [
      "SAVIMPEX DOO                                          Novi Sad, Gogoljeva 7, Republic of Serbia       PIB: 113113438",
    ],
  ];

  targetSheet.getCell(72, 12).values = [["Unterschrift und Stempel des Frachtführers"]];
  targetSheet.getCell(73, 12).values = [["Signature and stamp of the carrier"]];

  // 24
  targetSheet.getCell(67, 19).values = [["24"]];

  targetSheet.getCell(67, 20).values = [["Gut empfangen"]];
  targetSheet.getCell(68, 20).values = [["Goods received"]];

  targetSheet.getCell(69, 20).values = [["Ort"]];
  targetSheet.getCell(70, 20).values = [["Place"]];

  targetSheet.getCell(69, 23).values = [["am"]];
  targetSheet.getCell(70, 23).values = [["on"]];

  targetSheet.getCell(72, 20).values = [["Unterschrift und Stempel des Empfängers"]];
  targetSheet.getCell(73, 20).values = [["Signature and stamp of the consignee"]];

  targetSheet.getCell(74, 2).values = [
    [
      "Das CMR/IRU/Polen-Modell von 1976 für den internationalen Straßenverkehr entspricht den Regelungen der Internationalen Straßenverkehrsunion/IRU/.",
    ],
  ];
  targetSheet.getCell(75, 2).values = [
    [
      "The 1976 CMR/IRU/Poland model for international road transport complies with the rules of the International Road Transport Union/IRU/.",
    ],
  ];

  targetSheet.getCell(4, 0).values = [
    [
      "Die mit fett gedruckten Linien eingerahmten Rubriken müssen vom Frachtführer ausgefüllt werden.",
    ],
  ];
  targetSheet.getCell(4, 1).values = [
    ["The spaces framed with heavy lines must filied in by the carrier."],
  ];

  targetSheet.getCell(31, 0).values = [["19+20+21+22"]];

  targetSheet.getCell(36, 0).values = [["einschließlich"]];
  targetSheet.getCell(36, 1).values = [["including"]];

  targetSheet.getCell(39, 0).values = [["1 - 15"]];

  targetSheet.getCell(41, 0).values = [["Auszufüllen auf Verantwortung des Absenders"]];
  targetSheet.getCell(41, 1).values = [["To be completed on sender's responsability"]];
};

export const fillCMRTemplate = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const targetSheet: Excel.Worksheet = worksheets.getItem("cmr");

    targetSheet.getRange("A1:AD43").delete("Up");

    // targetSheet.getRange("1:2").format.rowHeight = 15;
    // targetSheet.getRange("3:4").format.rowHeight = 9;
    // targetSheet.getRange("5:5").format.rowHeight = 2.25;
    // targetSheet.getRange("6:9").format.rowHeight = 10.5;
    // targetSheet.getRange("10:11").format.rowHeight = 9;
    // targetSheet.getRange("12:15").format.rowHeight = 10.5;
    // targetSheet.getRange("16:17").format.rowHeight = 9;
    // targetSheet.getRange("18:19").format.rowHeight = 15.5;
    // targetSheet.getRange("20:21").format.rowHeight = 9;
    // targetSheet.getRange("22:23").format.rowHeight = 12;
    // targetSheet.getRange("24:25").format.rowHeight = 10.5;
    // targetSheet.getRange("26:27").format.rowHeight = 21.75;
    // targetSheet.getRange("28:29").format.rowHeight = 9.0;
    // targetSheet.getRange("30:31").format.rowHeight = 6.7;
    // targetSheet.getRange("32:36").format.rowHeight = 11.75;
    // targetSheet.getRange("37:40").format.rowHeight = 9.75;
    // targetSheet.getRange("41:42").format.rowHeight = 11.75;
    // targetSheet.getRange("43:44").format.rowHeight = 10.5;
    // targetSheet.getRange("45:60").format.rowHeight = 7.5;
    // targetSheet.getRange("61:64").format.rowHeight = 10.5;
    // targetSheet.getRange("65:66").format.rowHeight = 8.25;
    // targetSheet.getRange("67:67").format.rowHeight = 3;
    // targetSheet.getRange("68:69").format.rowHeight = 7.5;
    // targetSheet.getRange("70:71").format.rowHeight = 8.25;
    // targetSheet.getRange("72:75").format.rowHeight = 7.5;

    // targetSheet.getRange("A:B").format.columnWidth = 7.5; // 0.83
    // targetSheet.getRange("C:C").format.columnWidth = 18; // 2.71
    // targetSheet.getRange("D:D").format.columnWidth = 36.75; // 6.29
    // targetSheet.getRange("E:E").format.columnWidth = 25.5; // 4.14
    // targetSheet.getRange("F:F").format.columnWidth = 8.25; // 0.92
    // targetSheet.getRange("G:G").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("H:I").format.columnWidth = 4.5; // 0.5
    // // targetSheet.getRange("I:I").format.rowHeight = 0.83;
    // targetSheet.getRange("J:J").format.columnWidth = 56.25; // 10
    // targetSheet.getRange("K:K").format.columnWidth = 3.75; // 0.42
    // targetSheet.getRange("L:L").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("M:M").format.columnWidth = 22.5; // 3.57
    // targetSheet.getRange("N:N").format.columnWidth = 16.5; // 2.43
    // targetSheet.getRange("O:O").format.columnWidth = 12; // 1.57
    // targetSheet.getRange("P:P").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("Q:Q").format.columnWidth = 15; // 2.14
    // targetSheet.getRange("R:R").format.columnWidth = 45; // 7.86
    // targetSheet.getRange("S:S").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("T:T").format.columnWidth = 18.75; // 2.86
    // targetSheet.getRange("U:U").format.columnWidth = 19.5; // 3
    // targetSheet.getRange("V:V").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("W:W").format.columnWidth = 18.75; // 2.86
    // targetSheet.getRange("X:X").format.columnWidth = 28.5; // 4.71
    // targetSheet.getRange("Y:Y").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("Z:Z").format.columnWidth = 18.75; // 2.86
    // targetSheet.getRange("AA:AA").format.columnWidth = 27; // 4.43
    // targetSheet.getRange("AB:AB").format.columnWidth = 5.25; // 0.58
    // targetSheet.getRange("AC:AC").format.columnWidth = 7.5; // 0.83

    const cmrRange = targetSheet.getRangeByIndexes(0, 0, 75, 28);
    cmrRange.format.font.name = "Arial";
    cmrRange.format.font.size = 5;
    formatHeights(targetSheet);
    fillInTemplateData(targetSheet);
    textFormats(targetSheet);
    mergeRanges(targetSheet);

    await context.sync();
  });
};
export const fillCMRTemplateC = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const targetSheet: Excel.Worksheet = worksheets.getItem("cmr");

    targetSheet.getRange("A1:AD43").delete("Up");

    // targetSheet.getRange("1:2").format.rowHeight = 15;
    // targetSheet.getRange("3:4").format.rowHeight = 9;
    // targetSheet.getRange("5:5").format.rowHeight = 2.25;
    // targetSheet.getRange("6:9").format.rowHeight = 10.5;
    // targetSheet.getRange("10:11").format.rowHeight = 9;
    // targetSheet.getRange("12:15").format.rowHeight = 10.5;
    // targetSheet.getRange("16:17").format.rowHeight = 9;
    // targetSheet.getRange("18:19").format.rowHeight = 15.5;
    // targetSheet.getRange("20:21").format.rowHeight = 9;
    // targetSheet.getRange("22:23").format.rowHeight = 12;
    // targetSheet.getRange("24:25").format.rowHeight = 10.5;
    // targetSheet.getRange("26:27").format.rowHeight = 21.75;
    // targetSheet.getRange("28:29").format.rowHeight = 9.0;
    // targetSheet.getRange("30:31").format.rowHeight = 6.7;
    // targetSheet.getRange("32:36").format.rowHeight = 11.75;
    // targetSheet.getRange("37:40").format.rowHeight = 9.75;
    // targetSheet.getRange("41:42").format.rowHeight = 11.75;
    // targetSheet.getRange("43:44").format.rowHeight = 10.5;
    // targetSheet.getRange("45:60").format.rowHeight = 7.5;
    // targetSheet.getRange("61:64").format.rowHeight = 10.5;
    // targetSheet.getRange("65:66").format.rowHeight = 8.25;
    // targetSheet.getRange("67:67").format.rowHeight = 3;
    // targetSheet.getRange("68:69").format.rowHeight = 7.5;
    // targetSheet.getRange("70:71").format.rowHeight = 8.25;
    // targetSheet.getRange("72:75").format.rowHeight = 7.5;

    // targetSheet.getRange("A:B").format.columnWidth = 7.5; // 0.83
    // targetSheet.getRange("C:C").format.columnWidth = 18; // 2.71
    // targetSheet.getRange("D:D").format.columnWidth = 36.75; // 6.29
    // targetSheet.getRange("E:E").format.columnWidth = 25.5; // 4.14
    // targetSheet.getRange("F:F").format.columnWidth = 8.25; // 0.92
    // targetSheet.getRange("G:G").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("H:I").format.columnWidth = 4.5; // 0.5
    // // targetSheet.getRange("I:I").format.rowHeight = 0.83;
    // targetSheet.getRange("J:J").format.columnWidth = 56.25; // 10
    // targetSheet.getRange("K:K").format.columnWidth = 3.75; // 0.42
    // targetSheet.getRange("L:L").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("M:M").format.columnWidth = 22.5; // 3.57
    // targetSheet.getRange("N:N").format.columnWidth = 16.5; // 2.43
    // targetSheet.getRange("O:O").format.columnWidth = 12; // 1.57
    // targetSheet.getRange("P:P").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("Q:Q").format.columnWidth = 15; // 2.14
    // targetSheet.getRange("R:R").format.columnWidth = 45; // 7.86
    // targetSheet.getRange("S:S").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("T:T").format.columnWidth = 18.75; // 2.86
    // targetSheet.getRange("U:U").format.columnWidth = 19.5; // 3
    // targetSheet.getRange("V:V").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("W:W").format.columnWidth = 18.75; // 2.86
    // targetSheet.getRange("X:X").format.columnWidth = 28.5; // 4.71
    // targetSheet.getRange("Y:Y").format.columnWidth = 14.25; // 2
    // targetSheet.getRange("Z:Z").format.columnWidth = 18.75; // 2.86
    // targetSheet.getRange("AA:AA").format.columnWidth = 27; // 4.43
    // targetSheet.getRange("AB:AB").format.columnWidth = 5.25; // 0.58
    // targetSheet.getRange("AC:AC").format.columnWidth = 7.5; // 0.83

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

    // 2
    targetSheet.getRange("C10:C11").merge();
    targetSheet.getRange("C10:C10").values = [["2"]];
    targetSheet.getRange("D10:P10").merge();
    targetSheet.getRange("D10:D10").values = [["Empfänger (Name, Adresse, Land)"]];
    targetSheet.getRange("D11:P11").merge();
    targetSheet.getRange("D11:D11").values = [["Consignee (name, address, country)"]];

    targetSheet.getRange("C12:P12").merge();
    targetSheet.getRange("C13:P13").merge();
    targetSheet.getRange("C14:P14").merge();
    targetSheet.getRange("C15:P15").merge();

    setOuterBorders(targetSheet.getRange("C10:P15"));

    // 3
    targetSheet.getRange("C16:C17").merge();
    targetSheet.getRange("C16:C16").values = [["3"]];
    targetSheet.getRange("D16:P16").merge();
    targetSheet.getRange("D16:D16").values = [["Auslieferort des Gutes (Ort, Land)"]];
    targetSheet.getRange("D17:P17").merge();
    targetSheet.getRange("D17:D17").values = [["Place of delivery of the goods (place, country)"]];

    targetSheet.getRange("C18:P19").merge();

    setOuterBorders(targetSheet.getRange("C16:P19"));

    // 4
    targetSheet.getRange("C20:C21").merge();
    targetSheet.getRange("C20:C20").values = [["4"]];
    targetSheet.getRange("D20:P20").merge();
    targetSheet.getRange("D21:P21").merge();

    targetSheet.getRange("C22:P23").merge();

    setOuterBorders(targetSheet.getRange("C20:P23"));

    // 5
    targetSheet.getRange("C24:C25").merge();
    targetSheet.getRange("C24:C24").values = [["5"]];

    targetSheet.getRange("D24:P24").merge();
    targetSheet.getRange("D25:P25").merge();

    targetSheet.getRange("C26:P27").merge();

    setOuterBorders(targetSheet.getRange("C24:P27"));

    // 6

    targetSheet.getRange("C28:C29").merge();
    targetSheet.getRange("C28:C28").values = [["6"]];

    targetSheet.getRange("D28:F28").merge();
    targetSheet.getRange("D28:D28").values = [["Kennzeichen u. Nummern"]];
    targetSheet.getRange("D29:F29").merge();
    targetSheet.getRange("D29:D29").values = [["Marks and Nos"]];

    targetSheet.getRange("C30:F40").merge();

    targetSheet.getRange("D41:D42").merge();
    targetSheet.getRange("C41:C41").values = [["UN-Nr."]];

    targetSheet.getRange("E41:F41").merge();
    targetSheet.getRange("E42:F42").merge();
    targetSheet.getRange("E41:E41").values = [["Ben. s. Nr. 9"]];
    targetSheet.getRange("E42:E42").values = [["name s. nr. 9"]];

    targetSheet.getRange("C28:AA42").format.font.size = 5;
    targetSheet.getRange("C28:AA42").format.font.name = "Arial";

    // 7
    targetSheet.getRange("G28:G29").merge();
    targetSheet.getRange("G28:G28").values = [["7"]];

    targetSheet.getRange("H28:K28").merge();
    targetSheet.getRange("H28:H28").values = [["Anzahl der Pakete"]];
    targetSheet.getRange("H29:K29").merge();
    targetSheet.getRange("H29:H29").values = [["Number of pakages"]];

    targetSheet.getRange("G30:K40").merge();
    targetSheet.getRange("G30:G30").values = [["3 PAL"]];

    targetSheet.getRange("J41:J41").values = [["Gefahrzettelmuster-Nr."]];
    targetSheet.getRange("J42:J42").values = [["Hazard label sample no."]];

    // 8
    targetSheet.getRange("L28:L29").merge();
    targetSheet.getRange("L28:L28").values = [["8"]];
    targetSheet.getRange("M28:O28").merge();
    targetSheet.getRange("M28:M28").values = [["Art der Verpackung"]];
    targetSheet.getRange("M29:O29").merge();
    targetSheet.getRange("M29:M29").values = [["Method of packing"]];

    targetSheet.getRange("L30:O40").merge();
    targetSheet.getRange("L30:L30").values = [["Total 3 coll"]];
    targetSheet.getRange("L41:O42").merge();

    // 9
    targetSheet.getRange("P28:P29").merge();
    targetSheet.getRange("P28:P28").values = [["9"]];

    targetSheet.getRange("Q28:R28").merge();
    targetSheet.getRange("Q28:Q28").values = [["Bezeichnung des Gutes*"]];
    targetSheet.getRange("Q29:R29").merge();
    targetSheet.getRange("Q29:Q29").values = [["Nature of the goods*"]];

    targetSheet.getRange("P30:R40").merge();
    targetSheet.getRange("P30:P30").values = [["Consumer electronic goods"]];

    targetSheet.getRange("P41:Q41").merge();
    targetSheet.getRange("P41:P41").values = [["Verp.-Grp."]];
    targetSheet.getRange("P42:Q42").merge();
    targetSheet.getRange("P42:P42").values = [["Pack. group"]];

    setOuterBorders(targetSheet.getRange("C28:R40"));
    setOuterBorders(targetSheet.getRange("C41:R42"));

    // 10
    targetSheet.getRange("S28:S29").merge();
    targetSheet.getRange("S28:S28").values = [["10"]];

    targetSheet.getRange("T28:U28").merge();

    targetSheet.getRange("T28:T28").values = [["Statistiknr."]];

    targetSheet.getRange("T29:U29").merge();
    targetSheet.getRange("T29:T29").values = [["Statistical nr."]];

    targetSheet.getRange("S30:U42").merge();
    targetSheet.getRange("S30:S30").values = [["See packing list"]];

    setOuterBorders(targetSheet.getRange("S28:U42"));

    // 11
    targetSheet.getRange("V28:V29").merge();
    targetSheet.getRange("V28:V28").values = [["11"]];

    targetSheet.getRange("W28:X28").merge();
    targetSheet.getRange("W28:W28").values = [["Bruttogew. kg"]];

    targetSheet.getRange("W29:X29").merge();
    targetSheet.getRange("W29:W29").values = [["Gross weight kg"]];

    targetSheet.getRange("V30:X40").merge();
    targetSheet.getRange("V41:x41").merge();
    targetSheet.getRange("V30:V30").values = [["285.50 KG"]];
    targetSheet.getRange("V42:x42").merge();
    targetSheet.getRange("V41:V41").values = [["Total Gross:"]];

    targetSheet.getRange("V42:V42").values = [["10285.50 KG"]];

    setOuterBorders(targetSheet.getRange("V28:X40"));
    setOuterBorders(targetSheet.getRange("V41:X42"));

    // 12
    targetSheet.getRange("Y28:Y29").merge();
    targetSheet.getRange("Y28:Y28").values = [["12"]];

    targetSheet.getRange("Z28:AA28").merge();

    targetSheet.getRange("Z28:Z28").values = [["Volumen in m3"]];

    targetSheet.getRange("Z29:AA29").merge();
    targetSheet.getRange("Z29:Z29").values = [["Volume in m3"]];

    targetSheet.getRange("Y30:AA42").merge();

    setOuterBorders(targetSheet.getRange("Y28:AA42"));

    setOuterBorders(targetSheet.getRange("C28:AA42"));

    //13
    targetSheet.getRangeByIndexes(42, 2, 2, 1).merge();
    targetSheet.getRangeByIndexes(42, 3, 1, 13).merge();
    targetSheet.getRangeByIndexes(43, 3, 1, 13).merge();
    targetSheet.getRangeByIndexes(44, 2, 14, 14).merge();

    targetSheet.getCell(42, 2).values = [["13"]];
    targetSheet.getCell(42, 3).values = [
      ["Anweisungen des Absenders (Zoll-, amtl. Behandlungen, Sondervorschriften, etc.)"],
    ];
    targetSheet.getCell(43, 3).values = [["Sender's instructions"]];
    targetSheet.getCell(44, 2).values = [
      [
        'T/P "Akulovskiy" Code 10013010 SVH OOO "Crocus Interservice" 143002 Moskovskaya Obl, Odintsovskiy r-n, S. Akulovo, Ul. Novaya, D. 137      Lic. 10013/200111/10118/11 from 10.10.2024',
      ],
    ];

    // 14
    targetSheet.getRangeByIndexes(58, 2, 2, 1).merge();
    targetSheet.getCell(58, 2).values = [["14"]];

    targetSheet.getRangeByIndexes(58, 3, 1, 2).merge();
    targetSheet.getRangeByIndexes(59, 3, 1, 2).merge();

    targetSheet.getCell(58, 3).values = [["Rückerstattung"]];
    targetSheet.getCell(59, 3).values = [["Cash on delivery"]];

    targetSheet.getRangeByIndexes(58, 5, 2, 22).merge();
    targetSheet.getCell(58, 5).values = [["DAP MOSCOW"]];

    // 15

    targetSheet.getRangeByIndexes(60, 2, 2, 1).merge();
    targetSheet.getCell(60, 2).values = [["15"]];

    targetSheet.getRangeByIndexes(60, 3, 1, 13).merge();
    targetSheet.getRangeByIndexes(61, 3, 1, 13).merge();

    targetSheet.getCell(60, 3).values = [["Frachtzahlungsanweisungen"]];
    targetSheet.getCell(61, 3).values = [["Instruction as to payement carriage"]];

    targetSheet.getRangeByIndexes(62, 2, 1, 3).merge();
    targetSheet.getRangeByIndexes(63, 2, 1, 3).merge();

    targetSheet.getCell(62, 2).values = [["Frei/Carriage paid"]];
    targetSheet.getCell(63, 2).values = [["Unfrei/Carriage forward"]];

    targetSheet.getRangeByIndexes(62, 5, 1, 11).merge();
    targetSheet.getRangeByIndexes(63, 5, 1, 11).merge();

    // 16
    targetSheet.getRange("Q10:Q11").merge();
    // targetSheet.getRange("Q10:Q10").values = [["16"]];

    targetSheet.getRange("R10:AA10").merge();
    targetSheet.getRange("R11:AA11").merge();
    targetSheet.getRange("Q12:AA15").merge();

    targetSheet.getCell(9, 16).values = [["16"]];
    targetSheet.getCell(9, 17).values = [["Frachtführer (Name, Adresse, Land)"]];
    targetSheet.getCell(10, 17).values = [["Carrier (name, address, country)"]];
    targetSheet.getCell(11, 16).values = [["BG 2699-OI"]];

    // 17
    targetSheet.getRange("Q16:Q17").merge();
    targetSheet.getRange("R16:AA16").merge();
    targetSheet.getRange("R17:AA17").merge();
    targetSheet.getRange("Q18:AA19").merge();

    targetSheet.getCell(15, 16).values = [["17"]];
    targetSheet.getCell(15, 17).values = [["Nachfolgender Frachtführer (Name, Adresse, Land)"]];
    targetSheet.getCell(16, 17).values = [["Successive carriers (name, address, country)"]];
    targetSheet.getCell(17, 16).values = [["empty"]];

    // 18
    targetSheet.getRange("Q20:Q21").merge();
    targetSheet.getRange("R20:AA20").merge();
    targetSheet.getRange("R21:AA21").merge();
    targetSheet.getRange("Q22:AA27").merge();

    targetSheet.getCell(19, 16).values = [["18"]];
    targetSheet.getCell(19, 17).values = [["Vorbehalte und Bemerkungen der Frachtführer"]];
    targetSheet.getCell(20, 17).values = [["Carrier's reservations and observations"]];
    targetSheet.getCell(21, 16).values = [["em[ty"]];

    // 19
    targetSheet.getRangeByIndexes(42, 16, 2, 1).merge();
    targetSheet.getCell(42, 16).values = [["19"]];

    targetSheet.getCell(42, 17).values = [["Zu bezahlen vom"]];
    targetSheet.getCell(43, 17).values = [["To be paid by"]];

    targetSheet.getRangeByIndexes(42, 18, 1, 3).merge();
    targetSheet.getRangeByIndexes(43, 18, 1, 3).merge();
    targetSheet.getCell(42, 18).values = [["Absender"]];
    targetSheet.getCell(43, 18).values = [["Sender"]];

    targetSheet.getRangeByIndexes(42, 21, 1, 3).merge();
    targetSheet.getRangeByIndexes(43, 21, 1, 3).merge();
    targetSheet.getCell(42, 21).values = [["Währung"]];
    targetSheet.getCell(43, 21).values = [["Currency"]];

    targetSheet.getRangeByIndexes(42, 24, 1, 3).merge();
    targetSheet.getRangeByIndexes(43, 24, 1, 3).merge();
    targetSheet.getCell(42, 24).values = [["Empfänger"]];
    targetSheet.getCell(43, 24).values = [["Consignee"]];

    targetSheet.getRangeByIndexes(44, 16, 1, 2).merge();
    targetSheet.getRangeByIndexes(45, 16, 1, 2).merge();
    targetSheet.getCell(44, 16).values = [["Fracht"]];
    targetSheet.getCell(45, 16).values = [["Carriage"]];
    targetSheet.getCell(46, 16).values = [["Ermäßigung"]];
    targetSheet.getCell(47, 16).values = [["Reductions"]];
    targetSheet.getCell(48, 16).values = [["Zwischensumme"]];
    targetSheet.getCell(49, 16).values = [["Balance"]];
    targetSheet.getCell(50, 16).values = [["Zuschläge"]];
    targetSheet.getCell(51, 16).values = [["Supplement charges"]];
    targetSheet.getCell(52, 16).values = [["Nebengebühren"]];
    targetSheet.getCell(53, 16).values = [["Additional charges"]];
    targetSheet.getCell(54, 16).values = [["Sonstiges"]];
    targetSheet.getCell(55, 16).values = [["Miscellaneous"]];
    targetSheet.getCell(56, 16).values = [["Gesamtbetrag"]];
    targetSheet.getCell(57, 16).values = [["Total to be paid"]];

    targetSheet.getRangeByIndexes(44, 18, 2, 3).merge();
    targetSheet.getRangeByIndexes(44, 21, 2, 2).merge();
    targetSheet.getRangeByIndexes(44, 23, 2, 1).merge();
    targetSheet.getRangeByIndexes(44, 24, 2, 3).merge();

    targetSheet.getRangeByIndexes(46, 16, 1, 2).merge();
    targetSheet.getRangeByIndexes(47, 16, 1, 2).merge();

    targetSheet.getRangeByIndexes(46, 18, 2, 3).merge();
    targetSheet.getRangeByIndexes(46, 21, 2, 2).merge();
    targetSheet.getRangeByIndexes(46, 23, 2, 1).merge();
    targetSheet.getRangeByIndexes(46, 24, 2, 3).merge();

    targetSheet.getRangeByIndexes(48, 16, 1, 2).merge();
    targetSheet.getRangeByIndexes(49, 16, 1, 2).merge();

    targetSheet.getRangeByIndexes(48, 18, 2, 3).merge();
    targetSheet.getRangeByIndexes(48, 21, 2, 2).merge();
    targetSheet.getRangeByIndexes(48, 23, 2, 1).merge();
    targetSheet.getRangeByIndexes(48, 24, 2, 3).merge();

    targetSheet.getRangeByIndexes(50, 16, 1, 2).merge();
    targetSheet.getRangeByIndexes(51, 16, 1, 2).merge();

    targetSheet.getRangeByIndexes(50, 18, 2, 3).merge();
    targetSheet.getRangeByIndexes(50, 21, 2, 2).merge();
    targetSheet.getRangeByIndexes(50, 23, 2, 1).merge();
    targetSheet.getRangeByIndexes(50, 24, 2, 3).merge();

    targetSheet.getRangeByIndexes(52, 16, 1, 2).merge();
    targetSheet.getRangeByIndexes(53, 16, 1, 2).merge();

    targetSheet.getRangeByIndexes(52, 18, 2, 3).merge();
    targetSheet.getRangeByIndexes(52, 21, 2, 2).merge();
    targetSheet.getRangeByIndexes(52, 23, 2, 1).merge();
    targetSheet.getRangeByIndexes(52, 24, 2, 3).merge();

    targetSheet.getRangeByIndexes(54, 16, 1, 2).merge();
    targetSheet.getRangeByIndexes(55, 16, 1, 2).merge();

    targetSheet.getRangeByIndexes(54, 18, 2, 3).merge();
    targetSheet.getRangeByIndexes(54, 21, 2, 2).merge();
    targetSheet.getRangeByIndexes(54, 23, 2, 1).merge();
    targetSheet.getRangeByIndexes(54, 24, 2, 3).merge();

    targetSheet.getRangeByIndexes(56, 16, 1, 2).merge();
    targetSheet.getRangeByIndexes(57, 16, 1, 2).merge();

    targetSheet.getRangeByIndexes(56, 18, 2, 3).merge();
    targetSheet.getRangeByIndexes(56, 21, 2, 2).merge();
    targetSheet.getRangeByIndexes(56, 23, 2, 1).merge();
    targetSheet.getRangeByIndexes(56, 24, 2, 3).merge();

    //20
    targetSheet.getRangeByIndexes(60, 16, 2, 1).merge();
    targetSheet.getCell(60, 16).values = [["20"]];

    targetSheet.getRangeByIndexes(60, 17, 1, 10).merge();
    targetSheet.getRangeByIndexes(61, 17, 1, 10).merge();

    targetSheet.getCell(60, 17).values = [["Besondere Vereinbarungen"]];
    targetSheet.getCell(61, 17).values = [["Special agreements"]];

    targetSheet.getRangeByIndexes(62, 16, 4, 11).merge();

    // 21
    targetSheet.getRangeByIndexes(64, 2, 2, 1).merge();
    targetSheet.getCell(64, 2).values = [["21"]];

    targetSheet.getRangeByIndexes(64, 3, 1, 2).merge();
    targetSheet.getRangeByIndexes(65, 3, 1, 2).merge();

    targetSheet.getCell(64, 3).values = [["Ausgefertigt in"]];
    targetSheet.getCell(65, 3).values = [["Established in"]];

    targetSheet.getRangeByIndexes(64, 5, 2, 6).merge();

    targetSheet.getCell(64, 5).values = [["Belgrade"]];

    targetSheet.getCell(64, 11).values = [["am"]];
    targetSheet.getCell(65, 11).values = [["on"]];

    targetSheet.getRangeByIndexes(64, 12, 2, 4).merge();

    targetSheet.getCell(64, 12).values = [["11.03.2025"]];

    // 22

    targetSheet.getRangeByIndexes(67, 2, 2, 1).merge();
    targetSheet.getCell(67, 2).values = [["22"]];

    targetSheet.getRangeByIndexes(69, 2, 5, 1).merge();

    targetSheet.getRangeByIndexes(67, 3, 5, 8).merge();
    targetSheet.getCell(67, 3).values = [
      [
        "SAVIMPEX DOO                                          Novi Sad, Gogoljeva 7, Republic of Serbia       PIB: 113113438",
      ],
    ];

    targetSheet.getRangeByIndexes(72, 3, 1, 8).merge();
    targetSheet.getRangeByIndexes(73, 3, 1, 8).merge();

    targetSheet.getCell(72, 3).values = [["Signatur und Stempel des Absenders"]];
    targetSheet.getCell(73, 3).values = [["Signature and stamp of the sender"]];

    // 23
    targetSheet.getRangeByIndexes(67, 11, 2, 1).merge();
    targetSheet.getCell(67, 11).values = [["23"]];
    targetSheet.getRangeByIndexes(69, 11, 5, 1).merge();

    targetSheet.getRangeByIndexes(67, 12, 5, 7).merge();

    targetSheet.getCell(67, 12).values = [
      [
        "SAVIMPEX DOO                                          Novi Sad, Gogoljeva 7, Republic of Serbia       PIB: 113113438",
      ],
    ];

    targetSheet.getRangeByIndexes(72, 12, 1, 7).merge();
    targetSheet.getRangeByIndexes(73, 12, 1, 7).merge();

    targetSheet.getCell(72, 12).values = [["Unterschrift und Stempel des Frachtführers"]];
    targetSheet.getCell(73, 12).values = [["Signature and stamp of the carrier"]];

    // 24
    targetSheet.getRangeByIndexes(67, 19, 2, 1).merge();
    targetSheet.getCell(67, 19).values = [["24"]];
    targetSheet.getRangeByIndexes(69, 19, 5, 1).merge();

    targetSheet.getRangeByIndexes(67, 20, 1, 3).merge();
    targetSheet.getRangeByIndexes(68, 20, 1, 3).merge();
    targetSheet.getRangeByIndexes(67, 23, 2, 4).merge();

    targetSheet.getCell(67, 20).values = [["Gut empfangen"]];
    targetSheet.getCell(68, 20).values = [["Goods received"]];

    targetSheet.getCell(69, 20).values = [["Ort"]];
    targetSheet.getCell(70, 20).values = [["Place"]];

    targetSheet.getRangeByIndexes(69, 21, 2, 2).merge();

    targetSheet.getCell(69, 23).values = [["am"]];
    targetSheet.getCell(70, 23).values = [["on"]];

    targetSheet.getRangeByIndexes(69, 24, 2, 3).merge();

    targetSheet.getRangeByIndexes(71, 20, 1, 7).merge();
    targetSheet.getRangeByIndexes(72, 20, 1, 7).merge();
    targetSheet.getRangeByIndexes(73, 20, 1, 7).merge();

    targetSheet.getCell(72, 20).values = [["Unterschrift und Stempel des Empfängers"]];
    targetSheet.getCell(73, 20).values = [["Signature and stamp of the consignee"]];

    targetSheet.getRange("C75:AA75").merge();
    targetSheet.getRange("A75:AC75").format.fill.color = "000000";
    targetSheet.getRange("A75:AC75").format.font.color = "ffffff";

    targetSheet.getCell(74, 2).values = [
      [
        "Das CMR/IRU/Polen-Modell von 1976 für den internationalen Straßenverkehr entspricht den Regelungen der Internationalen Straßenverkehrsunion/IRU/.",
      ],
    ];
    targetSheet.getCell(75, 2).values = [
      [
        "The 1976 CMR/IRU/Poland model for international road transport complies with the rules of the International Road Transport Union/IRU/.",
      ],
    ];

    targetSheet.getRange("c76:AA76").merge();
    targetSheet.getRange("A76:AC76").format.fill.color = "000000";
    targetSheet.getRange("A76:AC76").format.font.color = "ffffff";

    targetSheet.getRangeByIndexes(4, 0, 27, 1).merge();
    targetSheet.getRangeByIndexes(4, 1, 27, 1).merge();

    targetSheet.getRangeByIndexes(4, 0, 27, 1).format.textOrientation = 90;
    targetSheet.getRangeByIndexes(4, 1, 27, 1).format.textOrientation = 90;

    targetSheet.getCell(4, 0).values = [
      [
        "Die mit fett gedruckten Linien eingerahmten Rubriken müssen vom Frachtführer ausgefüllt werden.",
      ],
    ];
    targetSheet.getCell(4, 1).values = [
      ["The spaces framed with heavy lines must filied in by the carrier."],
    ];

    targetSheet.getRangeByIndexes(31, 0, 5, 2).merge();
    targetSheet.getRangeByIndexes(31, 0, 5, 2).format.textOrientation = 90;
    targetSheet.getCell(31, 0).values = [["19+20+21+22"]];

    targetSheet.getRangeByIndexes(36, 0, 3, 1).merge();
    targetSheet.getRangeByIndexes(36, 0, 3, 1).format.textOrientation = 90;
    targetSheet.getRangeByIndexes(36, 1, 3, 1).merge();
    targetSheet.getRangeByIndexes(36, 1, 3, 1).format.textOrientation = 90;

    targetSheet.getCell(36, 0).values = [["einschließlich"]];
    targetSheet.getCell(36, 1).values = [["including"]];

    targetSheet.getRangeByIndexes(39, 0, 2, 2).merge();
    targetSheet.getRangeByIndexes(39, 0, 2, 2).format.textOrientation = 90;
    targetSheet.getCell(39, 0).values = [["1 - 15"]];

    targetSheet.getRangeByIndexes(41, 0, 15, 1).merge();
    targetSheet.getRangeByIndexes(41, 0, 15, 1).format.textOrientation = 90;
    targetSheet.getRangeByIndexes(41, 1, 15, 1).merge();
    targetSheet.getRangeByIndexes(41, 1, 15, 1).format.textOrientation = 90;
    targetSheet.getCell(41, 0).values = [["Auszufüllen auf Verantwortung des Absenders"]];
    targetSheet.getCell(41, 1).values = [["To be completed on sender's responsability"]];

    targetSheet.getUsedRange().format.font.name = "Arial";

  
  

    await context.sync();
  });
};
