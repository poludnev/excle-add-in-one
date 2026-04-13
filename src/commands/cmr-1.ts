import {
  consignees,
  places_of_delivery,
  senders_instructions,
  shippers,
} from "../constants/cmrConstants";
import { getColumnLetter } from "../utilities/helpers";
import {
  createSheetWithName,
  setAllBorders,
  setCenter,
  setOuterBorders,
  setThickOuterBorders,
} from "../utilities/utils";

const formatFieldNumberRange = (range: Excel.Range) => {
  range.format.font.size = 7;
  range.format.font.bold = true;
  setCenter(range);
};

const formatFieldDescriptionRange = (range: Excel.Range) => {
  // range.format.font.size = 6;
  range.format.horizontalAlignment = Excel.HorizontalAlignment.left;
  range.format.verticalAlignment = Excel.VerticalAlignment.center;
};

const formatDataRange = (range: Excel.Range, horCenter?: boolean) => {
  range.format.font.size = 8;
  range.format.font.bold = true;
  range.format.horizontalAlignment = horCenter
    ? Excel.HorizontalAlignment.center
    : Excel.HorizontalAlignment.left;
  range.format.verticalAlignment = Excel.VerticalAlignment.center;
};
const formatDataRangeLeft = (range: Excel.Range) => formatDataRange(range, false);
const formatDataRangeCenter = (range: Excel.Range) => formatDataRange(range, true);

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
  targetSheet.getRange("72:76").format.rowHeight = 7.5;

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
  targetSheet.getRange("W3:W4").merge();
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

  targetSheet.getRangeByIndexes(62, 16, 5, 11).merge();

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

  targetSheet.getRangeByIndexes(19, 27, 55, 1).merge();
  targetSheet.getRangeByIndexes(19, 28, 55, 1).merge();
};

const textFormats = (targetSheet: Excel.Worksheet) => {
  targetSheet.getRange("A1:AC76").format.font.name = "Arial";
  targetSheet.getRange("A1:AC76").format.font.size = 5;

  targetSheet.getRange("A1").format.font.size = 16;
  targetSheet.getRange("A1").format.verticalAlignment = Excel.VerticalAlignment.center;
  targetSheet.getRange("A1").format.horizontalAlignment = Excel.HorizontalAlignment.center;

  targetSheet.getRange("C1:D2").format.font.size = 8;
  targetSheet.getRange("A1:AC2").format.fill.color = "000000";
  targetSheet.getRange("A1:AC2").format.font.color = "ffffff";
  targetSheet.getRange("A75:AC76").format.fill.color = "000000";
  targetSheet.getRange("A75:AC76").format.font.color = "ffffff";

  targetSheet.getRange("A75:AA76").format.font.size = 6;
  targetSheet.getRange("A75:AA76").format.horizontalAlignment = Excel.HorizontalAlignment.left;
  targetSheet.getRange("A75:AA76").format.verticalAlignment = Excel.VerticalAlignment.center;

  targetSheet.getRange("A5:B42").format.textOrientation = 90;
  targetSheet.getRange("AB20:AC20").format.textOrientation = 90;

  const r3r4 = targetSheet.getRange("R3:R4");
  r3r4.format.font.size = 6;
  r3r4.format.verticalAlignment = Excel.VerticalAlignment.center;

  // targetSheet.getRange("A32:A32").format.textOrientation = 90;

  const c3 = targetSheet.getRange("C3");
  const w3 = targetSheet.getRange("w3");
  const x3 = targetSheet.getRange("x3");
  const aa3 = targetSheet.getRange("AA3");
  // c3.format.font.size = 7;
  // c3.format.font.bold = true;
  // setCenter(c3);

  const c10 = targetSheet.getRange("C10");
  // c10.format.font.size = 7;
  // c10.format.font.bold = true;
  // setCenter(c10);

  const c16 = targetSheet.getRange("C16");
  // c16.format.font.size = 7;
  // c16.format.font.bold = true;
  // setCenter(c16);

  const c20 = targetSheet.getRange("C20");
  // c20.format.font.size = 7;
  // c20.format.font.bold = true;
  // setCenter(c20);

  const c24 = targetSheet.getRange("C24");
  // c24.format.font.size = 7;
  // c24.format.font.bold = true;
  // setCenter(c24);
  const c28 = targetSheet.getRange("C28");
  // c28.format.font.size = 7;
  // c28.format.font.bold = true;
  // setCenter(c28);

  const g28 = targetSheet.getRange("G28");
  // g28.format.font.size = 7;
  // g28.format.font.bold = true;
  // setCenter(g28);

  const l28 = targetSheet.getRange("L28");
  // l28.format.font.size = 7;
  // l28.format.font.bold = true;
  // setCenter(l28);
  const p28 = targetSheet.getRange("p28");
  // p28.format.font.size = 7;
  // p28.format.font.bold = true;
  // setCenter(p28);

  const s28 = targetSheet.getRange("s28");
  // s28.format.font.size = 7;
  // s28.format.font.bold = true;
  // setCenter(s28);

  const v28 = targetSheet.getRange("v28");
  // v28.format.font.size = 7;
  // v28.format.font.bold = true;
  // setCenter(v28);

  const y28 = targetSheet.getRange("y28");
  // y28.format.font.size = 7;
  // y28.format.font.bold = true;
  // setCenter(y28);

  const c43 = targetSheet.getRange("c43");
  // c43.format.font.size = 7;
  // c43.format.font.bold = true;
  // setCenter(c43);

  const q10 = targetSheet.getRange("q10");
  const q16 = targetSheet.getRange("q16");
  const q20 = targetSheet.getRange("q20");
  const q43 = targetSheet.getRange("q43");

  // setCenter(q43);

  const c59 = targetSheet.getRange("c59");
  // setCenter(c59);

  const f59 = targetSheet.getRange("f59");
  // setCenter(f59);

  const c61 = targetSheet.getRange("c61");
  // setCenter(c61);

  const q61 = targetSheet.getRange("q61");
  // setCenter(q61);

  const c65 = targetSheet.getRange("c65");
  // setCenter(c65);

  const f65 = targetSheet.getRange("f65");
  // setCenter(f65);

  const m65 = targetSheet.getRange("m65");
  // setCenter(m65);

  const c68 = targetSheet.getRange("c68");
  // setCenter(c68);

  const l68 = targetSheet.getRange("l68");
  // setCenter(l68);

  const t68 = targetSheet.getRange("t68");
  // setCenter(t68);

  const fieldNumberRanges: Excel.Range[] = [
    c3,
    w3,
    x3,
    aa3,
    c10,
    c16,
    c20,
    c24,
    c28,
    g28,
    l28,
    p28,
    s28,
    v28,
    y28,
    c43,
    q10,
    q16,
    q20,
    q43,
    c59,
    c61,
    q61,
    c65,
    c68,
    l68,
    t68,
  ];

  fieldNumberRanges.forEach(formatFieldNumberRange);

  const d3d4 = targetSheet.getRange("D3:D4");
  const d10d11 = targetSheet.getRange("D10:D11");
  const r10r11 = targetSheet.getRange("R10:R11");
  const d16d17 = targetSheet.getRange("D16:D17");
  const r16r17 = targetSheet.getRange("R16:R17");
  const d20d21 = targetSheet.getRange("D20:D21");
  const r20r21 = targetSheet.getRange("R20:R21");
  const d24d25 = targetSheet.getRange("D24:D25");
  const d28d29 = targetSheet.getRange("D28:D29");
  const h28h29 = targetSheet.getRange("H28:H29");
  const m28m29 = targetSheet.getRange("M28:M29");
  const q28q29 = targetSheet.getRange("Q28:Q29");
  const t28t29 = targetSheet.getRange("T28:T29");
  const w28w29 = targetSheet.getRange("W28:W29");
  const z28z29 = targetSheet.getRange("Z28:Z29");
  const d43d44 = targetSheet.getRange("D43:D44");
  const r43aa44 = targetSheet.getRange("R43:AA44");
  const q45q58 = targetSheet.getRange("Q45:Q58");
  const d59q60 = targetSheet.getRange("D59:Q60");
  const d61p62 = targetSheet.getRange("D61:P62");
  const r61aa62 = targetSheet.getRange("R61:AA62");
  const c63c64 = targetSheet.getRange("C63:C64");
  const d65d66 = targetSheet.getRange("D65:D66");
  const c73c74 = targetSheet.getRange("C73:C74");
  const m73m74 = targetSheet.getRange("M73:M74");
  const u68aa74 = targetSheet.getRange("U68:AA74");

  const fieldDescriptionRanges: Excel.Range[] = [
    d10d11,
    r10r11,
    r16r17,
    r20r21,
    d3d4,
    d16d17,
    d20d21,
    d24d25,
    d28d29,
    h28h29,
    m28m29,
    q28q29,
    t28t29,
    w28w29,
    z28z29,
    d43d44,
    r43aa44,
    q45q58,
    d59q60,
    d61p62,
    r61aa62,
    c63c64,
    d65d66,
    c73c74,
    m73m74,
    u68aa74,
  ];
  fieldDescriptionRanges.forEach(formatFieldDescriptionRange);

  targetSheet.getRange("A1:AC76").format.wrapText = true;

  const shipperRange = targetSheet.getRange("C6:P9");
  const consigneeRange = targetSheet.getRange("C12:P15");
  const deliveryPlaceRange = targetSheet.getRange("C18:P19");
  const takingOverPlaceRange = targetSheet.getRange("C22:P23");
  const attachedDocumentsRange = targetSheet.getRange("C26:P27");
  const carrierRange = targetSheet.getRange("Q12:AA15");
  const successiveCarriersRange = targetSheet.getRange("Q18:AA19");
  const carrierReservartionRange = targetSheet.getRange("Q22:AA27");

  const numberOfPackagesRange = targetSheet.getRange("G30:K40");
  const methodOfPackingRange = targetSheet.getRange("L30:O40");
  const natureOfGoodsRange = targetSheet.getRange("P30:R40");
  const statisticalNumberRange = targetSheet.getRange("S30:U42");
  const grossWeightRange = targetSheet.getRange("V30:X40");
  const grossWeightTotalRange = targetSheet.getRange("V41:X42");

  const senderInstructionsRange = targetSheet.getRange("C45:C58");
  const cashOnDeliveryRange = targetSheet.getRange("F59:AA60");
  const establishedInRange = targetSheet.getRange("F65:K66");
  const establishedOnRange = targetSheet.getRange("M65:P66");
  const senderSignatureRange = targetSheet.getRange("D68:K72");
  const carrierSignatureRange = targetSheet.getRange("M68:S72");
  // shipperRange.format.font.size = 8;
  const dataRanges = [
    shipperRange,
    consigneeRange,
    deliveryPlaceRange,
    takingOverPlaceRange,
    attachedDocumentsRange,

    successiveCarriersRange,
    carrierReservartionRange,
    senderInstructionsRange,
    senderSignatureRange,
    carrierSignatureRange,
  ];

  dataRanges.forEach(formatDataRangeLeft);

  const dataRangesCenter = [
    carrierRange,
    numberOfPackagesRange,
    methodOfPackingRange,
    natureOfGoodsRange,
    statisticalNumberRange,
    grossWeightRange,
    grossWeightTotalRange,
    establishedInRange,
    establishedOnRange,
    cashOnDeliveryRange,
  ];
  dataRangesCenter.forEach(formatDataRangeCenter);
};

const bordersFormats = (targetSheet: Excel.Worksheet) => {
  const c3aa74 = targetSheet.getRange("C3:AA74");
  const c3p9 = targetSheet.getRange("C3:P9");
  const c10p15 = targetSheet.getRange("C10:P15");
  const c16p19 = targetSheet.getRange("C16:P19");
  const c20p23 = targetSheet.getRange("c20:P23");
  const c24p27 = targetSheet.getRange("C24:P27");

  const q3aa9 = targetSheet.getRange("Q3:AA9");
  const q10aa15 = targetSheet.getRange("Q10:AA15");
  const q16aa19 = targetSheet.getRange("Q16:AA19");
  const q20aa27 = targetSheet.getRange("Q20:AA27");
  const c28aa42 = targetSheet.getRange("C28:AA42");
  const c28r42 = targetSheet.getRange("C28:R42");
  const c28r40 = targetSheet.getRange("C28:R40");
  const c41r42 = targetSheet.getRange("C41:R42");
  const s28u42 = targetSheet.getRange("S28:U42");
  const v28x42 = targetSheet.getRange("V28:X42");
  const v41x42 = targetSheet.getRange("V31:X42");
  const y28aa42 = targetSheet.getRange("Y28:AA42");

  const c43aa58 = targetSheet.getRange("C43:AA58");
  // const q43aa58 = targetSheet.getRange("Q43:AA58");
  const q43r58 = targetSheet.getRange("Q43:R58");
  const s43u58 = targetSheet.getRange("S43:U58");
  const v43x58 = targetSheet.getRange("V43:X58");
  const y43aa58 = targetSheet.getRange("Y43:AA58");
  const q43aaq44 = targetSheet.getRange("Q43:AA44");
  const q45aa48 = targetSheet.getRange("Q45:AA48");
  const q49aa56 = targetSheet.getRange("Q49:AA56");
  const q57aa58 = targetSheet.getRange("Q57:AA58");

  const s45aa58 = targetSheet.getRange("S45:AA58");

  setAllBorders(s45aa58);

  // thinBorderRange(s45aa58);
  //
  [
    c3aa74,
    c3p9,
    c10p15,
    c16p19,
    c20p23,
    c24p27,
    q3aa9,
    q10aa15,
    q16aa19,
    q20aa27,
    c28aa42,
    c28r42,
    c28r40,
    c41r42,
    s28u42,
    v28x42,
    v41x42,
    y28aa42,
    c43aa58,
    q43r58,
    s43u58,
    v43x58,
    y43aa58,
    q43aaq44,
    q45aa48,
    q49aa56,
    q57aa58,
  ].forEach(setOuterBorders);

  const ac59aa60 = targetSheet.getRange("AC59:AA60");
  const c61p64 = targetSheet.getRange("C61:P64");
  const c65p67 = targetSheet.getRange("C65:P67");
  const q61aa67 = targetSheet.getRange("Q61:AA67");
  const c68k74 = targetSheet.getRange("C68:K74");
  const l68s74 = targetSheet.getRange("L68:S74");
  const t68aa74 = targetSheet.getRange("T68:AA74");

  [ac59aa60, c61p64, c65p67, q61aa67, c68k74, t68aa74].forEach(setOuterBorders);

  setThickOuterBorders(l68s74);

  const q43aa58 = targetSheet.getRange("Q43:AA58");
  setThickOuterBorders(q43aa58);
};

function fillInTemplateData(targetSheet: Excel.Worksheet) {
  targetSheet.getRange("A1").values = [["1"]];
  targetSheet.getRange("C1:C1").values = [["Exemplar für den Absender"]];
  targetSheet.getRange("C2:C2").values = [["Copy for sender"]];

  targetSheet.getRange("C3:C3").values = [["1"]];

  targetSheet.getRange("D3:D3").values = [["Absender (Name, Adresse, Land)"]];

  targetSheet.getRange("D4:D4").values = [["Sender (name, address, country)"]];

  targetSheet.getRange("R3:R3").values = [["INTERNATIONALER FRACHTBRIEF"]];
  targetSheet.getRange("R4:R4").values = [["INTERNATIONAL CONSIGNEMENT NOTE"]];

  targetSheet.getRange("W3:W3").values = [["№"]];
  // targetSheet.getRange("X3:X3").values = [["E204-210"]];
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
  // targetSheet.getCell(21, 2).values = [["Beograd, 11.03.2025"]];

  // 5
  targetSheet.getRange("C24:C24").values = [["5"]];
  targetSheet.getCell(23, 3).values = [["Beigefügte Dokumente"]];
  targetSheet.getCell(24, 3).values = [["Documents attached"]];
  // targetSheet.getCell(25, 2).values = [["Invoice № UT-EX-91-TE  dated 10.03.2025"]];

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

  // targetSheet.getRange("G30:G30").values = [["3 PAL"]];

  targetSheet.getRange("J41:J41").values = [["Gefahrzettelmuster-Nr."]];
  targetSheet.getRange("J42:J42").values = [["Hazard label sample no."]];

  // 8
  targetSheet.getRange("L28:L28").values = [["8"]];
  targetSheet.getRange("M28:M28").values = [["Art der Verpackung"]];
  targetSheet.getRange("M29:M29").values = [["Method of packing"]];

  // targetSheet.getRange("L30:L30").values = [["Total 3 coll"]];

  // 9
  targetSheet.getRange("P28:P28").values = [["9"]];

  targetSheet.getRange("Q28:Q28").values = [["Bezeichnung des Gutes*"]];
  targetSheet.getRange("Q29:Q29").values = [["Nature of the goods*"]];

  // targetSheet.getRange("P30:P30").values = [["Consumer electronic goods"]];

  targetSheet.getRange("P41:P41").values = [["Verp.-Grp."]];
  targetSheet.getRange("P42:P42").values = [["Pack. group"]];

  // 10
  targetSheet.getRange("S28:S28").values = [["10"]];

  targetSheet.getRange("T28:T28").values = [["Statistiknr."]];

  targetSheet.getRange("T29:T29").values = [["Statistical nr."]];

  // targetSheet.getRange("S30:S30").values = [["See packing list"]];

  // 11
  targetSheet.getRange("V28:V28").values = [["11"]];

  targetSheet.getRange("W28:W28").values = [["Bruttogew. kg"]];

  targetSheet.getRange("W29:W29").values = [["Gross weight kg"]];

  // targetSheet.getRange("V30:V30").values = [["285.50 KG"]];
  targetSheet.getRange("V41:V41").values = [["Total Gross:"]];

  // targetSheet.getRange("V42:V42").values = [["10285.50 KG"]];

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
  // targetSheet.getCell(44, 2).values = [
  //   [
  //     'T/P "Akulovskiy" Code 10013010 SVH OOO "Crocus Interservice" 143002 Moskovskaya Obl, Odintsovskiy r-n, S. Akulovo, Ul. Novaya, D. 137      Lic. 10013/200111/10118/11 from 10.10.2024',
  //   ],
  // ];

  // 14
  targetSheet.getCell(58, 2).values = [["14"]];

  targetSheet.getCell(58, 3).values = [["Rückerstattung"]];
  targetSheet.getCell(59, 3).values = [["Cash on delivery"]];
  // targetSheet.getCell(58, 5).values = [["DAP MOSCOW"]];

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
  // targetSheet.getCell(11, 16).values = [["BG 2699-OI"]];

  // 17
  targetSheet.getCell(15, 16).values = [["17"]];
  targetSheet.getCell(15, 17).values = [["Nachfolgender Frachtführer (Name, Adresse, Land)"]];
  targetSheet.getCell(16, 17).values = [["Successive carriers (name, address, country)"]];
  // targetSheet.getCell(17, 16).values = [["empty"]];

  // 18
  targetSheet.getCell(19, 16).values = [["18"]];
  targetSheet.getCell(19, 17).values = [["Vorbehalte und Bemerkungen der Frachtführer"]];
  targetSheet.getCell(20, 17).values = [["Carrier's reservations and observations"]];
  // targetSheet.getCell(21, 16).values = [["em[ty"]];

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

  // targetSheet.getCell(64, 5).values = [["Belgrade"]];

  targetSheet.getCell(64, 11).values = [["am"]];
  targetSheet.getCell(65, 11).values = [["on"]];

  // targetSheet.getCell(64, 12).values = [["11.03.2025"]];

  // 22
  targetSheet.getCell(67, 2).values = [["22"]];

  // targetSheet.getCell(67, 3).values = [
  //   [
  //     "SAVIMPEX DOO                                          Novi Sad, Gogoljeva 7, Republic of Serbia       PIB: 113113438",
  //   ],
  // ];

  targetSheet.getCell(72, 3).values = [["Signatur und Stempel des Absenders"]];
  targetSheet.getCell(73, 3).values = [["Signature and stamp of the sender"]];

  // 23
  targetSheet.getCell(67, 11).values = [["23"]];

  // targetSheet.getCell(67, 12).values = [
  //   [
  //     "SAVIMPEX DOO                                          Novi Sad, Gogoljeva 7, Republic of Serbia       PIB: 113113438",
  //   ],
  // ];

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

  targetSheet.getCell(19, 28).values = [
    [
      "*Bei gefährlichen Gütern  sind neben der möglichen Zertifizierung in der letzten Zeile der Spalte die Angaben zur Klasse, der UN-Nummer und ggf. der Verpackungsgruppe anzugeben.",
    ],
  ];
}

const fillCMRDataHandler = (targetSheet: Excel.Worksheet) => {
  const cmrData: {
    cmrNumber: string;
    shipper: [string, string | null, string | null, string | null];
    consignee: string[];
  } = {
    cmrNumber: "E204-210",
    shipper: ["SAVIMPEX DOO", "Novi Sad, Gogoljeva 7, Republic of Serbia", "PIB: 113113438", null],
    consignee: [],
  };
  const { cmrNumber } = cmrData;

  const cmrNumberRange = targetSheet.getRange("X3:X3");
  cmrNumberRange.values = [[cmrNumber]];
  // 1
  const shipperRange = targetSheet.getRange("C6:C9");

  // shipperRange.clear(Excel.ClearApplyTo.contents); // clears values and formulas
  // await context.sync();
  shipperRange.values = [
    ["SAVIMPEX DOO"],
    ["Novi Sad, Gogoljeva 7, Republic of Serbia"],
    ["PIB: 113113438"],
    [null],
  ];
  // const shipper1 = targetSheet.getRange("C6:C6");
  // const shipper2 = targetSheet.getRange("C7:C7");
  // const shipper3 = targetSheet.getRange("C8:C8");
  // const shipper4 = targetSheet.getRange("C9:C9");

  // shipper1.values = [["SAVIMPEX DOO"]];
  // shipper2.values = [["Novi Sad, Gogoljeva 7, Republic of Serbia"]];
  // shipper3.values = [["PIB: 113113438"]];
  // shipper4.values = [["PIB: 113113438"]];

  // 2
  const consignee1 = targetSheet.getRange("C12:C12");
  const consignee2 = targetSheet.getRange("C13:C13");
  const consignee3 = targetSheet.getRange("C14:C14");
  const consignee4 = targetSheet.getRange("C15:C15");

  consignee1.values = [["URSUS TRADE LLC"]];
  consignee2.values = [["Russia, Moscow, Hlebniy pereulok 19A, 121069"]];
  consignee3.values = [["TIN 7735189429"]];
  consignee4.values = [["TIN 7735189429"]];

  // 3

  const deliveryPlace = targetSheet.getRange("C18:C18");
  deliveryPlace.values = [
    [
      'LLC "URSUS TRADE" / SVH OOO "Crocus Interservice" 143002 Moskovskaya Obl, Odintsovskiy r-n, S. Akulovo, Ul. Novaya, D. 137',
    ],
  ];

  // 4
  const takingOverPlace = targetSheet.getRange("C22:C22");
  takingOverPlace.values = [["Beograd, 11.03.2025"]];

  // 5
  const documentAttached = targetSheet.getRange("C26:C26");
  documentAttached.values = [["Invoice № UT-EX-91-TE  dated 10.03.2025"]];

  const carrierName = targetSheet.getRange("Q12:Q12");
  carrierName.values = [["BG 2699-OI"]];

  const numberOfPackages = targetSheet.getRange("G30:G30");
  numberOfPackages.values = [["3 PAL"]];

  const methodOfPacking = targetSheet.getRange("L30:L30");
  methodOfPacking.values = [["Total 3 coll"]];

  const natureOfGoods = targetSheet.getRange("P30:P30");
  natureOfGoods.values = [["Consumer electronic goods"]];

  const statisticalNumber = targetSheet.getRange("S30:S30");
  statisticalNumber.values = [["See packing list"]];

  const grossWeight = targetSheet.getRange("V30:V30");
  grossWeight.values = [["285.50 KG"]];

  const grossWeightTotal = targetSheet.getRange("V42:V42");
  grossWeightTotal.values = [["10285.50 KG"]];

  const senderInstructions = targetSheet.getRange("C45:C45");
  senderInstructions.values = [
    [
      'T/P "Akulovskiy" Code 10013010 SVH OOO "Crocus Interservice" 143002 Moskovskaya Obl, Odintsovskiy r-n, S. Akulovo, Ul. Novaya, D. 137      Lic. 10132/200111/10118/12 from 01.07.2025',
    ],
  ];

  const cashOnDelivery = targetSheet.getRange("F59:F59");
  cashOnDelivery.values = [["DAP MOSCOW"]];

  const establishedIn = targetSheet.getRange("F65:F65");
  establishedIn.values = [["Belgrade"]];

  const establishedOn = targetSheet.getRange("M65:M65");
  establishedOn.values = [["11.03.2025"]];

  const senderSignature = targetSheet.getRange("D68:D68");
  senderSignature.values = [
    ["SAVIMPEX DOO Novi Sad, Gogoljeva 7, Republic of Serbia       PIB: 113113438"],
  ];
};

export const fillCMRTemplate = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const targetSheet: Excel.Worksheet = worksheets.getItem("cmr");

    const cmrRange = targetSheet.getRangeByIndexes(0, 0, 76, 26);
    cmrRange.format.font.name = "Arial";

    formatHeights(targetSheet);
    fillInTemplateData(targetSheet);
    mergeRanges(targetSheet);
    textFormats(targetSheet);
    // fillCMRDataHandler(targetSheet);
    bordersFormats(targetSheet);

    await context.sync();
  });
};

export const fillCMRTemplateBySheet = async (
  targetSheet: Excel.Worksheet,
  context: Excel.RequestContext
) => {
  const cmrRange = targetSheet.getRangeByIndexes(0, 0, 76, 26);
  cmrRange.format.font.name = "Arial";

  formatHeights(targetSheet);
  fillInTemplateData(targetSheet);
  mergeRanges(targetSheet);
  textFormats(targetSheet);
  // fillCMRDataHandler(targetSheet);
  bordersFormats(targetSheet);

  await context.sync();
};

type TCMRDataFileds = [
  "cmr_num",
  "shipper_1",
  "shipper_2",
  "shipper_3",
  "shipper_4",
  "consignee_1",
  "consignee_2",
  "consignee_3",
  "consignee_4",
  "place_of_delivery",
  "take_over",
  "attachment",
  "marks",
  "number_of_packs",
  "method_of_packing",
  "nature_of_goods",
  "statistical_num",
  "gross",
  "total_gross",
  "volume",
  "senders_instruct",
  "cash_on_delivery",
  "prepayment_instruction",
  "carrier",
  "successive_carriers",
  "carrier_reservations",
  "to_be_paid",
  "special_agreement",
  "established_in",
  "established_on",
  "sender_stamp",
  "carrier_stamp",
];

const cmrDataFields: TCMRDataFileds = [
  "cmr_num",
  "shipper_1",
  "shipper_2",
  "shipper_3",
  "shipper_4",
  "consignee_1",
  "consignee_2",
  "consignee_3",
  "consignee_4",
  "place_of_delivery",
  "take_over",
  "attachment",
  "marks",
  "number_of_packs",
  "method_of_packing",
  "nature_of_goods",
  "statistical_num",
  "gross",
  "total_gross",
  "volume",
  "senders_instruct",
  "cash_on_delivery",
  "prepayment_instruction",
  "carrier",
  "successive_carriers",
  "carrier_reservations",
  "to_be_paid",
  "special_agreement",
  "established_in",
  "established_on",
  "sender_stamp",
  "carrier_stamp",
];

export const fillCMRData = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const targetSheet: Excel.Worksheet = worksheets.getItem("cmr");
    const dataSourcceSheet: Excel.Worksheet = worksheets.getItem("cmr_data");

    const dataRange = dataSourcceSheet.getRange("B1:C32");

    dataRange.load("values");

    await context.sync();

    console.log(dataRange.values);
    const dataObj: { [key: string]: string } = {};
    dataRange.values.forEach((row) => {
      const key = row[0];
      const value = row[1];
      dataObj[key] = value;
    });

    targetSheet.getRange("X3:Z4").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C6:P9").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C12:P15").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C18:P19").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C22:P23").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C26:P27").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C30:R40").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("S30:U42").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("V30:X40").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("V42:X42").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C45:P58").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("F59:AA60").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("Q12:AA15").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("F65:K66").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("M65:P66").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("D68:K72").clear(Excel.ClearApplyTo.contents);

    await context.sync();

    const cmrNumberRange = targetSheet.getRange("X3:X3");
    cmrNumberRange.values = [[dataObj["cmr_num"]]];
    // 1
    const shipperRange = targetSheet.getRange("C6:C9");
    shipperRange.values = [
      [dataObj["shipper_1"]],
      [dataObj["shipper_2"]],
      [dataObj["shipper_3"]],
      [dataObj["shipper_4"]],
    ];

    // 2
    const consigneeRange = targetSheet.getRange("C12:C15");
    consigneeRange.values = [
      [dataObj["consignee_1"]],
      [dataObj["consignee_2"]],
      [dataObj["consignee_3"]],
      [dataObj["consignee_4"]],
    ];

    // 3
    const deliveryPlaceRange = targetSheet.getRange("C18:C18");
    deliveryPlaceRange.values = [[dataObj["place_of_delivery"]]];
    // 4
    const takingOverPlaceRange = targetSheet.getRange("C22:C22");
    takingOverPlaceRange.values = [[dataObj["take_over"]]];

    // 5
    const attachedDocumentsRange = targetSheet.getRange("C26:C26");
    attachedDocumentsRange.values = [[dataObj["attachment"]]];
    // 6

    // 7
    const numberOfPackagesRange = targetSheet.getRange("G30:G30");
    numberOfPackagesRange.values = [[dataObj["number_of_packs"]]];

    // 8
    const methodOfPackingRange = targetSheet.getRange("L30:L30");
    methodOfPackingRange.values = [[dataObj["method_of_packing"]]];
    // 9
    const natureOfGoodsRange = targetSheet.getRange("P30:P30");
    natureOfGoodsRange.values = [[dataObj["nature_of_goods"]]];

    // 10
    const statisticalNumberRange = targetSheet.getRange("S30:S30");
    statisticalNumberRange.values = [[dataObj["statistical_num"]]];
    // 11
    const grossWeightRange = targetSheet.getRange("V30:V30");
    grossWeightRange.values = [[dataObj["gross"]]];
    const grossWeightTotalRange = targetSheet.getRange("V42:V42");
    grossWeightTotalRange.values = [[dataObj["total_gross"]]];

    // 12

    // 13
    const senderInstructionsRange = targetSheet.getRange("C45:C45");
    senderInstructionsRange.values = [[dataObj["senders_instruct"]]];

    // 14
    const cashOnDeliveryRange = targetSheet.getRange("F59:F59");
    cashOnDeliveryRange.values = [[dataObj["cash_on_delivery"]]];

    // 16
    const carrierRange = targetSheet.getRange("Q12:Q12");
    carrierRange.values = [[dataObj["carrier"]]];
    // 17
    const successiveCarriersRange = targetSheet.getRange("Q18:Q18");
    successiveCarriersRange.values = [[dataObj["successive_carriers"]]];

    // 18
    const carrierReservationRange = targetSheet.getRange("Q22:Q22");
    carrierReservationRange.values = [[dataObj["carrier_reservation"]]];
    // 21
    const establishedInRange = targetSheet.getRange("F65:F65");
    establishedInRange.values = [[dataObj["established_in"]]];

    const establishedOnRange = targetSheet.getRange("M65:M65");
    establishedOnRange.values = [[dataObj["established_on"]]];

    // 22
    const senderSignatureRange = targetSheet.getRange("D68:D68");
    senderSignatureRange.values = [[dataObj["sender_stamp"]]];
  });
};

const checkCMRDataValidity = async (sourceSheetName: string, context: Excel.RequestContext) => {
  const {
    workbook: { worksheets },
  } = context;
  const dataSourcceSheet: Excel.Worksheet = worksheets.getItem(sourceSheetName);

  const dataRange = dataSourcceSheet.getRange("B1:C32");

  dataRange.load("values");

  await context.sync();

  console.log(dataRange.values);
  const dataObj: { [key: string]: string } = {};
  dataRange.values.forEach((row) => {
    const key = row[0];
    const value = row[1];
    if (value.startsWith("DEFAULT_"))
      throw Error(`Sheet ${sourceSheetName} has Invalid data value for key ${key}  : ${value}`);
    dataObj[key] = value;
  });
};

export const fillCMRDataBySheetNames = async (targetSheetName: string, sourceSheetName: string) => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    const targetSheet: Excel.Worksheet = worksheets.getItem(targetSheetName);

    targetSheet.getRange("X3:Z4").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C6:P9").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C12:P15").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C18:P19").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C22:P23").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C26:P27").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C30:R40").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("S30:U42").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("V30:X40").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("V42:X42").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("C45:P58").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("F59:AA60").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("Q12:AA15").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("F65:K66").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("M65:P66").clear(Excel.ClearApplyTo.contents);
    targetSheet.getRange("D68:K72").clear(Excel.ClearApplyTo.contents);

    await context.sync();

    const dataSourcceSheet: Excel.Worksheet = worksheets.getItem(sourceSheetName);

    const dataRange = dataSourcceSheet.getRange("B1:C32");

    dataRange.load("values");

    await context.sync();

    console.log(dataRange.values);
    const dataObj: { [key: string]: string } = {};
    dataRange.values.forEach((row) => {
      const key = row[0];
      const value = row[1];
      if (value.startsWith("DEFAULT_"))
        throw Error(`Sheet ${sourceSheetName} has Invalid data value for key ${key}  : ${value}`);
      dataObj[key] = value;
    });

    const cmrNumberRange = targetSheet.getRange("X3:X3");
    cmrNumberRange.values = [[dataObj["cmr_num"]]];
    // 1
    const shipperRange = targetSheet.getRange("C6:C9");
    shipperRange.values = [
      [dataObj["shipper_1"]],
      [dataObj["shipper_2"]],
      [dataObj["shipper_3"]],
      [dataObj["shipper_4"]],
    ];

    // 2
    const consigneeRange = targetSheet.getRange("C12:C15");
    consigneeRange.values = [
      [dataObj["consignee_1"]],
      [dataObj["consignee_2"]],
      [dataObj["consignee_3"]],
      [dataObj["consignee_4"]],
    ];

    // 3
    const deliveryPlaceRange = targetSheet.getRange("C18:C18");
    deliveryPlaceRange.values = [[dataObj["place_of_delivery"]]];
    // 4
    const takingOverPlaceRange = targetSheet.getRange("C22:C22");
    takingOverPlaceRange.values = [[dataObj["take_over"]]];

    // 5
    const attachedDocumentsRange = targetSheet.getRange("C26:C26");
    attachedDocumentsRange.values = [[dataObj["attachment"]]];
    // 6

    // 7
    const numberOfPackagesRange = targetSheet.getRange("G30:G30");
    numberOfPackagesRange.values = [[dataObj["number_of_packs"]]];

    // 8
    const methodOfPackingRange = targetSheet.getRange("L30:L30");
    methodOfPackingRange.values = [[dataObj["method_of_packing"]]];
    // 9
    const natureOfGoodsRange = targetSheet.getRange("P30:P30");
    natureOfGoodsRange.values = [[dataObj["nature_of_goods"]]];

    // 10
    const statisticalNumberRange = targetSheet.getRange("S30:S30");
    statisticalNumberRange.values = [[dataObj["statistical_num"]]];
    // 11
    const grossWeightRange = targetSheet.getRange("V30:V30");
    grossWeightRange.values = [[dataObj["gross"]]];
    const grossWeightTotalRange = targetSheet.getRange("V42:V42");
    grossWeightTotalRange.values = [[dataObj["total_gross"]]];

    // 12

    // 13
    const senderInstructionsRange = targetSheet.getRange("C45:C45");
    senderInstructionsRange.values = [[dataObj["senders_instruct"]]];

    // 14
    const cashOnDeliveryRange = targetSheet.getRange("F59:F59");
    cashOnDeliveryRange.values = [[dataObj["cash_on_delivery"]]];

    // 16
    const carrierRange = targetSheet.getRange("Q12:Q12");
    carrierRange.values = [[dataObj["carrier"]]];
    // 17
    const successiveCarriersRange = targetSheet.getRange("Q18:Q18");
    successiveCarriersRange.values = [[dataObj["successive_carriers"]]];

    // 18
    const carrierReservationRange = targetSheet.getRange("Q22:Q22");
    carrierReservationRange.values = [[dataObj["carrier_reservation"]]];
    // 21
    const establishedInRange = targetSheet.getRange("F65:F65");
    establishedInRange.values = [[dataObj["established_in"]]];

    const establishedOnRange = targetSheet.getRange("M65:M65");
    establishedOnRange.values = [[dataObj["established_on"]]];

    // 22
    const senderSignatureRange = targetSheet.getRange("D68:D68");
    senderSignatureRange.values = [[dataObj["sender_stamp"]]];
  });
};

const fillCMRDataConstants = async (
  dataSourcceSheet: Excel.Worksheet,
  context: Excel.RequestContext
) => {
  console.log("fillCMRDataConstants");
  const dataRange = dataSourcceSheet.getRange("A1:C32");

  const getFieldnumber = (index: number): number => {
    if (index === 0) return 0;
    if (index < 5) return 1;
    if (index < 9) return 2;
    if (index < 18) return index - 6;
    if (index < 29) return index - 7;
    return index - 8;
  };
  console.log("fillCMRDataConstants 2");

  const defaultShipper = shippers["SAVIMPEX"];
  const defaultConsignee = consignees["URSUS_TRADE"];
  const defaultPlaceOfDelivery = places_of_delivery["CROCUS"];
  const defaultSendersInstructions = senders_instructions["AKULOVO"];

  dataRange.values = cmrDataFields.map((field, index) => {
    if (field === "cmr_num") {
      return [getFieldnumber(index), field, "DEFAULT_CMR_NUMBER"];
    }
    if (field.startsWith("shipper")) {
      const lineIndex = parseInt(field.split("_")[1], 10) - 1;
      const lineValue = defaultShipper[lineIndex] || "";
      return [getFieldnumber(index), field, lineValue];
    }
    if (field.startsWith("consignee")) {
      const lineIndex = parseInt(field.split("_")[1], 10) - 1;
      const lineValue = defaultConsignee[lineIndex] || "";
      return [getFieldnumber(index), field, lineValue];
    }
    if (field === "place_of_delivery") {
      return [getFieldnumber(index), field, defaultPlaceOfDelivery];
    }
    if (field === "take_over") {
      return [getFieldnumber(index), field, '=TEXTJOIN(", ",1,C29,C30)'];
    }
    if (field === "attachment") {
      return [getFieldnumber(index), field, "DEFAULT_INVOICE_NUMBER_DATE"];
    }
    if (field === "number_of_packs") {
      return [getFieldnumber(index), field, "DEFAULT_QUANTITY_OF_PALS"];
    }
    if (field === "method_of_packing") {
      return [getFieldnumber(index), field, "DEFAULT_TOTAL_QUANTITY_OF_COLLS"];
    }
    if (field === "nature_of_goods") {
      return [getFieldnumber(index), field, "DEFAULT_TYPE_OF_GOODS"];
    }
    if (field === "statistical_num") {
      return [getFieldnumber(index), field, "See packing list"];
    }
    if (field === "gross") {
      return [getFieldnumber(index), field, "DEFAULT_GROSS_WEIGHT_KG"];
    }
    if (field === "total_gross") {
      return [getFieldnumber(index), field, "DEFAULT_TOTAL_GROSS_WEIGHT_KG"];
    }
    if (field === "senders_instruct") {
      return [getFieldnumber(index), field, defaultSendersInstructions];
    }
    if (field === "cash_on_delivery") {
      return [getFieldnumber(index), field, "DEFAULT_DAP_MOSCOW"];
    }
    if (field === "carrier") {
      return [getFieldnumber(index), field, "DEFAULT_TRUCK_NUMBER"];
    }

    if (field === "established_in") {
      return [getFieldnumber(index), field, "DEFAULT_PLACE"];
    }
    if (field === "established_on") {
      return [getFieldnumber(index), field, "DEFAULT_DATE"];
    }

    if (field === "sender_stamp") {
      return [getFieldnumber(index), field, defaultShipper.slice(0, 3).join(" ")];
    }
    return [getFieldnumber(index), field, ""];
  });

  const valuesRange = dataSourcceSheet.getRange("C1:C32");

  const cf = valuesRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);

  cf.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "DEFAULT",
  };
  cf.textComparison.format.fill.color = "#FFFF00"; // yellow

  await context.sync();
};

export const fillCMR_data_constants = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;
    const dataSourcceSheet: Excel.Worksheet = worksheets.getItem("cmr_data");
    const dataRange = dataSourcceSheet.getRange("A1:C32");

    const getFieldnumber = (index: number): number => {
      if (index === 0) return 0;
      if (index < 5) return 1;
      if (index < 9) return 2;
      if (index < 18) return index - 6;
      if (index < 29) return index - 7;
      return index - 8;
    };

    const defaultShipper = shippers["SAVIMPEX"];
    const defaultConsignee = consignees["URSUS_TRADE"];
    const defaultPlaceOfDelivery = places_of_delivery["CROCUS"];
    const defaultSendersInstructions = senders_instructions["AKULOVO"];

    dataRange.values = cmrDataFields.map((field, index) => {
      if (field === "cmr_num") {
        return [getFieldnumber(index), field, "DEFAULT_CMR_NUMBER"];
      }
      if (field.startsWith("shipper")) {
        const lineIndex = parseInt(field.split("_")[1], 10) - 1;
        const lineValue = defaultShipper[lineIndex] || "";
        return [getFieldnumber(index), field, lineValue];
      }
      if (field.startsWith("consignee")) {
        const lineIndex = parseInt(field.split("_")[1], 10) - 1;
        const lineValue = defaultConsignee[lineIndex] || "";
        return [getFieldnumber(index), field, lineValue];
      }
      if (field === "place_of_delivery") {
        return [getFieldnumber(index), field, defaultPlaceOfDelivery];
      }
      if (field === "take_over") {
        return [getFieldnumber(index), field, '=TEXTJOIN(", ",1,C29,C30)'];
      }
      if (field === "attachment") {
        return [getFieldnumber(index), field, "DEFAULT_INVOICE_NUMBER_DATE"];
      }
      if (field === "number_of_packs") {
        return [getFieldnumber(index), field, "DEFAULT_QUANTITY_OF_PALS"];
      }
      if (field === "method_of_packing") {
        return [getFieldnumber(index), field, "DEFAULT_TOTAL_QUANTITY_OF_COLLS"];
      }
      if (field === "nature_of_goods") {
        return [getFieldnumber(index), field, "DEFAULT_TYPE_OF_GOODS"];
      }
      if (field === "statistical_num") {
        return [getFieldnumber(index), field, "See packing list"];
      }
      if (field === "gross") {
        return [getFieldnumber(index), field, "DEFAULT_GROSS_WEIGHT_KG"];
      }
      if (field === "total_gross") {
        return [getFieldnumber(index), field, "DEFAULT_TOTAL_GROSS_WEIGHT_KG"];
      }
      if (field === "senders_instruct") {
        return [getFieldnumber(index), field, defaultSendersInstructions];
      }
      if (field === "cash_on_delivery") {
        return [getFieldnumber(index), field, "DEFAULT_DAP_MOSCOW"];
      }
      if (field === "carrier") {
        return [getFieldnumber(index), field, "DEFAULT_TRUCK_NUMBER"];
      }

      if (field === "established_in") {
        return [getFieldnumber(index), field, "DEFAULT_PLACE"];
      }
      if (field === "established_on") {
        return [getFieldnumber(index), field, "DEFAULT_DATE"];
      }

      if (field === "sender_stamp") {
        return [getFieldnumber(index), field, defaultShipper.slice(0, 3).join(" ")];
      }
      return [getFieldnumber(index), field, ""];
    });

    await context.sync();
  });
};

// interface Icmr {
//   number: string;
//   invoices: { number: string; date: string }[];
//   grossWeightKg: number;
// }
const readInstructionsSheetData = async (
  sourceSheet: Excel.Worksheet,
  context: Excel.RequestContext
) => {
  // const sourceSheet: Excel.Worksheet = worksheets.getItem("instruction");

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

  interface Icmr {
    number: string;
    invoices: { number: string; date: string }[];
    grossWeightKg: number;
  }
  const cmrDataList: Icmr[] = dynamicRange.values.reduce((acc, row) => {
    // console.log("ROW:", row);
    const cmrNumber = row[21];
    if (!cmrNumber) throw Error("CMR number is missing in instruction sheet data");
    const invoiceNumber = row[0];
    const invoiceDate = row[22];
    if (!invoiceDate) throw Error("Invoice date is missing in instruction sheet data");
    const grossWeightKg = row[11];
    const accItem = acc.find((item) => item.number === cmrNumber);
    if (accItem) {
      const accInvoice = accItem.invoices.find((inv) => inv.number === invoiceNumber);
      if (!accInvoice) {
        accItem.invoices.push({ number: invoiceNumber, date: invoiceDate });
      }
      accItem.grossWeightKg += Number(grossWeightKg);
    } else {
      acc.push({
        number: cmrNumber,
        invoices: [{ number: invoiceNumber, date: invoiceDate }],
        grossWeightKg: Number(grossWeightKg),
      });
    }
    return acc;
  }, [] as Icmr[]);

  for (let i = 0; i < cmrDataList.length - 1; i += 1) {
    const cmrData = cmrDataList[i];
    const invocies = cmrData.invoices;
    // console.log("Processing CMR:", cmrData.number, invocies);
    invocies.forEach((inv) => {
      // console.log(" Checking invoice:", inv.number);
      const invoiceNumber = inv.number;
      // console.log(" Checking invoice number:", invoiceNumber);
      for (let j = i + 1; j < cmrDataList.length; j += 1) {
        // console.log("  Comparing with CMR index:", j, i);
        const nextCmr = cmrDataList[j];
        // console.log("  Against CMR:", nextCmr);
        const nextCmrInvocies = nextCmr.invoices;
        // console.log("   Next CMR invoices:", nextCmrInvocies);
        const invoiceNumberIndex = nextCmrInvocies.findIndex(
          (invItem) => invItem.number === invoiceNumber
        );
        if (invoiceNumberIndex !== -1) {
          throw new Error(`Duplicate invoice number ${invoiceNumber} found in CMRs
            ${cmrData.number} and ${nextCmr.number}`);
        }
      }
    });
  }
  console.log("CMR DATA LIST:", cmrDataList);
  return cmrDataList;
};

export const fillCMR_data_values = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;
    // const targetSheet: Excel.Worksheet = worksheets.getItem("cmr_data");

    worksheets.load("items/name");
    await context.sync();

    worksheets.items
      .filter((s) => s.name.startsWith("cmr_"))
      .forEach((sheet) => {
        console.log("Deleting existing CMR sheet:", sheet.name);
        sheet.delete();
      });
    await context.sync();

    const sourceSheet: Excel.Worksheet = worksheets.getItem("instruction");

    // console.log("SOURCE SHEET:", sourceSheet.name);
    const cmrData = await readInstructionsSheetData(sourceSheet, context);
    console.log("CMR DATA FROM FUNCTION:", cmrData);
    // const cmrDataSheetNames = cmrData.map((cmr) => "cmr_data_" + cmr.number);

    // Promise.all(
    //   cmrDataSheetNames.map(async (sheetName) => {
    //     try {
    //       await createSheetWithName(sheetName);
    //     } catch (error) {
    //       console.error(`Error creating sheet ${sheetName}:`, error);
    //     } finally {
    //       return;
    //     }
    //   })
    // ).then(() => {
    //   console.log("All sheets created");
    // });

    cmrData.forEach(async (cmr) => {
      const cmrSheetName = "cmr_data_" + cmr.number;
      await createSheetWithName(cmrSheetName);
      const cmrSheet: Excel.Worksheet = worksheets.getItem(cmrSheetName);
      await fillCMRDataConstants(cmrSheet, context);
      const invoiceNames =
        "Invoice № " + cmr.invoices.map((inv) => inv.number + " dated " + inv.date).join(", ");
      console.log("Filling CMR sheet:", cmrSheetName, " with invoices:", invoiceNames);

      const cmrNumberRange = cmrSheet.getRange("C1:C1");
      cmrNumberRange.values = [[cmr.number]];
      const cmrAttachmentRange = cmrSheet.getRange("C12:C12");

      console.log("cmrAttachmentRange:", cmrAttachmentRange);
      cmrAttachmentRange.values = [[invoiceNames]];
      const grossWeightRange = cmrSheet.getRange("C18:C18");
      grossWeightRange.values = [[`${cmr.grossWeightKg.toFixed(2)} KG`]];
      const grossTotalWeightRange = cmrSheet.getRange("C19:C19");
      grossTotalWeightRange.values = [[`${cmr.grossWeightKg.toFixed(2)} KG`]];

      const dataRange = cmrSheet.getRange("B:B");
      dataRange.format.autofitColumns();

      cmrSheet.activate();
      await context.sync();
    });

    await context.sync();
  });
};

export const makeCMRs = async () => {
  Excel.run(async (context: Excel.RequestContext) => {
    const {
      workbook: { worksheets },
    } = context;

    worksheets.load("items/name");
    await context.sync();

    console.log(
      "Worksheets loaded:",
      worksheets.items.map((s) => s.name)
    );

    const cmrDataSheetsNames = worksheets.items
      .map((s) => s.name)
      .filter((name) => name.startsWith("cmr_data_"));
    console.log("All sheets:", cmrDataSheetsNames);

    if (cmrDataSheetsNames.length === 0) {
      throw Error("No CMR data sheets found ");
    }

    for (let i = 0; i < cmrDataSheetsNames.length; i += 1) {
      await checkCMRDataValidity(cmrDataSheetsNames[i], context);
    }

    worksheets.items
      .filter((s) => s.name.startsWith("cmr_") && !s.name.includes("cmr_data_"))
      .forEach((sheet) => {
        console.log("Deleting existing CMR sheet:", sheet.name);
        sheet.delete();
      });
    await context.sync();

    const cmrSheetsNames = cmrDataSheetsNames.map((name) => name.replace("cmr_data_", "cmr_"));

    console.log("CMR data sheets names:", cmrDataSheetsNames);
    console.log("CMR sheets names to create:", cmrSheetsNames);

    const { worksheet: firstCMRSheet } = await createSheetWithName(cmrSheetsNames[0]);

    // console.log("Creating CMR from sheet:", firstCMRSheet);

    // // await fillCMRTemplateBySheet(firstCMRSheet, context);

    const targetSheet: Excel.Worksheet = worksheets.getItem(cmrSheetsNames[0]);

    formatHeights(targetSheet);
    fillInTemplateData(targetSheet);
    mergeRanges(targetSheet);
    textFormats(targetSheet);
    bordersFormats(targetSheet);

    await context.sync();

    for (let i = 1; i < cmrSheetsNames.length; i += 1) {
      targetSheet.copy(Excel.WorksheetPositionType.after, worksheets.getLast());
      await context.sync();

      const copied = worksheets.getLast();
      // console.log("Renaming copied sheet to:", cmrSheetsNames[i]);
      copied.name = cmrSheetsNames[i];
      await context.sync();
    }

    // // cmrSheetsNames.forEach(async (sheetName, index) => {
    // //   if (index === 0) return;
    // //   targetSheet.copy(Excel.WorksheetPositionType.after, targetSheet);
    // //   await context.sync();
    // //   const copied = worksheets.getLast();

    // //   console.log("Renaming copied sheet to:", sheetName);
    // //   copied.name = sheetName;
    // //   await context.sync();
    // // });

    for (let i = 0; i < cmrSheetsNames.length; i += 1) {
      await fillCMRDataBySheetNames(cmrSheetsNames[i], cmrDataSheetsNames[i]);
    }

    // // Promise.all(
    // //   cmrSheetsNames.map(async (sheetName) => {
    // //     try {

    // //     } catch (error) {
    // //       console.error(`Error creating CMR from sheet ${sheetName}:`, error);
    // //     } finally {
    // //       return;
    // //     }
    // //   })
    // // ).then(() => {
    // //   console.log("All CMRs created");
    // // });
    await context.sync();
  });
};
