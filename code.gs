function onInstall() {
  onOpen();
}
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("DataMate")
    .addItem("Save Record", "save")
    .addSeparator()
    .addItem("Reset Input", "copyInput1")
    .addSeparator()
    .addItem("Reset View_Print", "view")
    .addSeparator()
    .addItem("New Dataset", "newfile")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("AddressBlock").addItem("Add Contact Sheets", "contacts")
    )
    .addItem("New Contact", "NewContact")
    .addItem("AddressBlock Name", "EditAddressSheet")
    .addItem("AddressBlock Company", "EditAddressSheet1");
    
  menu.addToUi();
}



function edit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const addressSheet = ss.getSheetByName("Address");

  addressSheet.getRange('F2').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 44, FALSE)');
  addressSheet.getRange('F3').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 52, FALSE)');
  addressSheet.getRange('F4').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 56, FALSE) & ", " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 57, FALSE) & " " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 58, FALSE)');

  addressSheet.getRange('F5').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getRange("F5").setFormula('=HYPERLINK(VLOOKUP(' + lookupValue + ', contacts!A:CJ, 16, FALSE))');

  addressSheet.getRange('F6').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 40, FALSE)');
  addressSheet.getRange('F7').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 42, FALSE)');
  addressSheet.getRange('F8').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 26, FALSE)');
  addressSheet.getRange('F9').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 30, FALSE) & ", " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 31, FALSE) & " " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 32, FALSE)');
  addressSheet.getRange('F10').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 63, FALSE)');
  addressSheet.getRange('F11').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 67, FALSE) & ", " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 68, FALSE) & " " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 69, FALSE)');
  addressSheet.getRange('F12').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 5, FALSE)');
  addressSheet.getRange('F13').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 20, FALSE)');
  addressSheet.getRange('F14').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 22, FALSE)');
}

function newcontact() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newContact = ss.getSheetByName("NewContact");
  const contactsSheet = ss.getSheetByName("contacts");

   contactsSheet.insertRowAfter(1);

newContact.getRange('B1').copyTo(contactsSheet.getRange('contacts!B2'), { contentsOnly: true });
newContact.getRange('B2').copyTo(contactsSheet.getRange('contacts!C2'), { contentsOnly: true });
newContact.getRange('B3').copyTo(contactsSheet.getRange('contacts!D2'), { contentsOnly: true });
newContact.getRange('B4').copyTo(contactsSheet.getRange('contacts!AR2'), { contentsOnly: true });
newContact.getRange('B5').copyTo(contactsSheet.getRange('contacts!AZ2'), { contentsOnly: true });
newContact.getRange('B6').copyTo(contactsSheet.getRange('contacts!BD2'), { contentsOnly: true });
newContact.getRange('B7').copyTo(contactsSheet.getRange('contacts!BE2'), { contentsOnly: true });
newContact.getRange('B8').copyTo(contactsSheet.getRange('contacts!BF2'), { contentsOnly: true });
newContact.getRange('B9').copyTo(contactsSheet.getRange('contacts!P2'), { contentsOnly: true });
newContact.getRange('B10').copyTo(contactsSheet.getRange('contacts!AN2'), { contentsOnly: true });
newContact.getRange('B11').copyTo(contactsSheet.getRange('contacts!AP2'), { contentsOnly: true });
newContact.getRange('B12').copyTo(contactsSheet.getRange('contacts!Z2'), { contentsOnly: true });
newContact.getRange('B13').copyTo(contactsSheet.getRange('contacts!AD2'), { contentsOnly: true });
newContact.getRange('B14').copyTo(contactsSheet.getRange('contacts!AE2'), { contentsOnly: true });
newContact.getRange('B15').copyTo(contactsSheet.getRange('contacts!AF2'), { contentsOnly: true });
newContact.getRange('B16').copyTo(contactsSheet.getRange('contacts!BK2'), { contentsOnly: true });
newContact.getRange('B17').copyTo(contactsSheet.getRange('contacts!BO2'), { contentsOnly: true });
newContact.getRange('B18').copyTo(contactsSheet.getRange('contacts!BP2'), { contentsOnly: true });
newContact.getRange('B19').copyTo(contactsSheet.getRange('contacts!BQ2'), { contentsOnly: true });
newContact.getRange('B20').copyTo(contactsSheet.getRange('contacts!E2'), { contentsOnly: true });
newContact.getRange('B21').copyTo(contactsSheet.getRange('contacts!T2'), { contentsOnly: true });
newContact.getRange('B22').copyTo(contactsSheet.getRange('contacts!V2'), { contentsOnly: true });

contactsSheet.getRange('A2').activate();
contactsSheet.getCurrentCell().setFormula('=CONCATENATE(B2," ",C2," ",D2)');
contactsSheet.getRange('A:A').activate();
contactsSheet.getRange('A1').copyTo(contactsSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

contactsSheet.getRange('A1').activate();
contactsSheet.getRange('A1').getFilter().sort(1, false);

newContact.getRange('B1:B22').activate();
newContact.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
}

function copyInput1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Sheet1");
  var targetSheet = ss.getSheetByName("Input");
  
  // Define the range to copy
  var copyRange = sourceSheet.getRange("A3:Q48");
  var targetRange = targetSheet.getRange("A3:Q48");
  
  // Copy everything from source to target
  copyRange.copyTo(targetRange); // This will copy values, formats, and formulas
  
  // Copy column widths
  var sourceColWidths = [];
  var lastColumnSource = sourceSheet.getLastColumn();
  var lastColumnTarget = targetSheet.getLastColumn();
  
  // Ensure we only consider columns up to the last column in both sheets
  var columnsToCopy = Math.min(lastColumnSource, lastColumnTarget, 17); // A to Q = 17 columns
  
  for (var i = 1; i <= columnsToCopy; i++) {
    sourceColWidths.push(sourceSheet.getColumnWidth(i));
  }
  
  // Set column widths in target sheet, but only for existing columns
  for (var j = 1; j <= columnsToCopy; j++) {
    targetSheet.setColumnWidth(j, sourceColWidths[j - 1]);
  }
  
  // Select cell C4 in the target sheet
  targetSheet.activate();
  targetSheet.getRange("C4").activate();
  
  // Show a message box to notify the user
  SpreadsheetApp.getUi().alert("Sheet copied from 'Sheet1' to 'Input' successfully.");
}
function copyInput2() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Sheet1");
  var targetSheet = ss.getSheetByName("View_Print");
  
  // Define the range to copy
  var copyRange = sourceSheet.getRange("A3:Q48");
  var targetRange = targetSheet.getRange("A3:Q48");
  
  // Copy everything from source to target
  copyRange.copyTo(targetRange); // This will copy values, formats, and formulas
  
  // Copy column widths
  var sourceColWidths = [];
  var lastColumnSource = sourceSheet.getLastColumn();
  var lastColumnTarget = targetSheet.getLastColumn();
  
  // Ensure we only consider columns up to the last column in both sheets
  var columnsToCopy = Math.min(lastColumnSource, lastColumnTarget, 17); // A to Q = 17 columns
  
  for (var i = 1; i <= columnsToCopy; i++) {
    sourceColWidths.push(sourceSheet.getColumnWidth(i));
  }
  
  // Set column widths in target sheet, but only for existing columns
  for (var j = 1; j <= columnsToCopy; j++) {
    targetSheet.setColumnWidth(j, sourceColWidths[j - 1]);
  }
  
  // Select cell C4 in the target sheet
  targetSheet.activate();
  targetSheet.getRange("A1").activate();

 }


function newfile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    "Input",
    "View_Print",
    "Log",
    "Update",
    "Data",
  ];

  sheetNames.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // Insert a new sheet with the specified name
      ss.insertSheet(sheetName);
    }
  });
  const inputSheet = ss.getSheetByName("Input");
  const dataSheet = ss.getSheetByName("Data");
  const viewPrintSheet = ss.getSheetByName("View_Print");
  const updateSheet = ss.getSheetByName("Update");
  const logSheet = ss.getSheetByName("Log");

  inputSheet.getRange("A1:Q1").activate();
  inputSheet.getActiveRangeList().setBackground("#a4c2f4");
  inputSheet.getRange("P2:Q2").activate();
  inputSheet.getActiveRangeList().setBackground("#a4c2f4");
  inputSheet.getRange("A1").activate();
  inputSheet.getCurrentCell().setValue("Log 1");
  inputSheet.getRange("B1").activate();
  inputSheet.getCurrentCell().setValue("Log 2");
  inputSheet.getRange("C1").activate();
  inputSheet.getCurrentCell().setValue("Log 3");
  inputSheet.getRange("D1").activate();
  inputSheet.getCurrentCell().setValue("Log 4");
  inputSheet.getRange("E1").activate();
  inputSheet.getCurrentCell().setValue("Log 5");
  inputSheet.getRange("F1").activate();
  inputSheet.getCurrentCell().setValue("Log 6");
  inputSheet.getRange("G1").activate();
  inputSheet.getCurrentCell().setValue("Log 7");
  inputSheet.getRange("H1").activate();
  inputSheet.getCurrentCell().setValue("Log 8");
  inputSheet.getRange("I1").activate();
  inputSheet.getCurrentCell().setValue("Log 9");
  inputSheet.getRange("J1").activate();
  inputSheet.getCurrentCell().setValue("Log 10");
  inputSheet.getRange("K1").activate();
  inputSheet.getCurrentCell().setValue("Log 11");
  inputSheet.getRange("L1").activate();
  inputSheet.getCurrentCell().setValue("Log 12");
  inputSheet.getRange("M1").activate();
  inputSheet.getCurrentCell().setValue("Update 1");
  inputSheet.getRange("N1").activate();
  inputSheet.getCurrentCell().setValue("Update 2");
  inputSheet.getRange("O1").activate();
  inputSheet.getCurrentCell().setValue("Update 3");
  inputSheet.getRange("P1:Q2").merge();
  inputSheet.getRange("P1").setFormula('=HYPERLINK("https://datamateapp.github.io/help.html", "Help")');
  const cell = inputSheet.getRange("P1");
  cell.setFontWeight("bold");
  cell.setFontSize(16);
  cell.setFontColor("#FF0000");
  cell.setHorizontalAlignment("center");
  cell.setVerticalAlignment("middle");

  inputSheet.getRange("A3:Q48").activate();
  inputSheet.setCurrentCell(inputSheet.getRange("Q48"));
  inputSheet
    .getActiveRangeList()
    .setBorder(false, false, false, false, false, false)
    .setBorder(
      true,
      true,
      true,
      true,
      null,
      null,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );
  inputSheet.getRange("A1:O2").activate();
  inputSheet.setCurrentCell(inputSheet.getRange("O1"));
  inputSheet
    .getActiveRangeList()
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );


  viewPrintSheet.getRange("A1").activate();
  viewPrintSheet.getRange("A1:Q1").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#a4c2f4");
  viewPrintSheet.getRange("A2").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#a4c2f4");
  viewPrintSheet.getRange("P2:Q2").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#a4c2f4");

  viewPrintSheet.getRange("M1").setFormula("=Input!M1");
  viewPrintSheet.getRange("N1").setFormula("=Input!N1");
  viewPrintSheet.getRange("O1").setFormula("=Input!O1");

  viewPrintSheet.getRange('A3:Q48').activate();
viewPrintSheet.setCurrentCell(viewPrintSheet.getRange('Q48'));
viewPrintSheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(false, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID); // Top border is set to 'false'
viewPrintSheet.setHiddenGridlines(true);

  viewPrintSheet.getRange("B2:L2").activate();
  viewPrintSheet.setCurrentCell(viewPrintSheet.getRange("L2"));
  viewPrintSheet.getActiveRange().mergeAcross();
  viewPrintSheet
    .getRange("B2:L2")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(viewPrintSheet.getRange("Data!$A:$A"), true)
        .build()
    );
  viewPrintSheet.getRange("B2:L2").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#d9ead3");

  logSheet.getRange("A2").activate();
  logSheet.getCurrentCell().setValue("Anything Log");
  logSheet
    .getActiveRangeList()
    .setFontSize(11)
    .setFontSize(14)
    .setFontWeight("bold");
  logSheet.getRange("A3").activate();
  logSheet.getCurrentCell().setValue("Date");
  logSheet.getRange("B3").activate();
  logSheet.getCurrentCell().setFormula("=TODAY()");
  logSheet.getRange("A9:O10").activate();
  logSheet.getRange("A9:O10").createFilter();

  updateSheet
    .getRangeList([
      "A:A",
      "E:E",
      "F:F",
      "G:G",
      "H:H",
      "I:I",
      "J:J",
      "K:K",
      "L:L",
      "M:M",
      "N:N",
      "O:O",
      "P:P",
      "Q:Q",
    ])
    .activate()
    .setBackground("#f3f3f3");
  updateSheet.getRange("A1:L1").activate();
  updateSheet.setCurrentCell(updateSheet.getRange("L1"));
  updateSheet
    .getActiveRangeList()
    .setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );

  updateSheet.getRange("B1").setFormula("=View_Print!M1");
  updateSheet.getRange("C1").setFormula("=View_Print!N1");
  updateSheet.getRange("D1").setFormula("=View_Print!O1");
  updateSheet.getRange("E1").setFormula("=Input!A1");
  updateSheet.getRange("F1").setFormula("=Input!B1");
  updateSheet.getRange("G1").setFormula("=Input!C1");
  updateSheet.getRange("H1").setFormula("=Input!D1");
  updateSheet.getRange("I1").setFormula("=Input!E1");
  updateSheet.getRange("J1").setFormula("=Input!F1");
  updateSheet.getRange("K1").setFormula("=Input!G1");
  updateSheet.getRange("L1").setFormula("=Input!H1");
  updateSheet.getRange("M1").setFormula("=Input!I1");
  updateSheet.getRange("N1").setFormula("=Input!J1");
  updateSheet.getRange("O1").setFormula("=Input!K1");
  updateSheet.getRange("P1").setFormula("=Input!L1");

  const lastColumn = dataSheet.getMaxColumns();
  const columnsToAdd = 800; // Number of columns from the last column to AFC

  dataSheet.insertColumnsAfter(lastColumn, columnsToAdd);
  const lastRow = dataSheet.getMaxRows();
  const rowsToAdd = 10000;

  dataSheet.insertRowsAfter(lastRow, rowsToAdd);
  dataSheet.getRange("S1").activate();
  dataSheet.getRange("S1").setFormula("=Input!A1");
  dataSheet.getRange("T1").activate();
  dataSheet.getRange("T1").setFormula("=Input!B1");
  dataSheet.getRange("U1").activate();
  dataSheet.getRange("U1").setFormula("=Input!C1");
  dataSheet.getRange("V1").activate();
  dataSheet.getRange("V1").setFormula("=Input!D1");
  dataSheet.getRange("W1").activate();
  dataSheet.getRange("W1").setFormula("=Input!E1");
  dataSheet.getRange("X1").activate();
  dataSheet.getRange("X1").setFormula("=Input!F1");
  dataSheet.getRange("Y1").activate();
  dataSheet.getRange("Y1").setFormula("=Input!G1");
  dataSheet.getRange("Z1").activate();
  dataSheet.getRange("Z1").setFormula("=Input!H1");
  dataSheet.getRange("AA1").activate();
  dataSheet.getRange("AA1").setFormula("=Input!I1");
  dataSheet.getRange("AB1").activate();
  dataSheet.getRange("AB1").setFormula("=Input!J1");
  dataSheet.getRange("AC1").activate();
  dataSheet.getRange("AC1").setFormula("=Input!K1");
  dataSheet.getRange("AD1").activate();
  dataSheet.getRange("AD1").setFormula("=Input!L1");
  dataSheet.getRange("AE1").activate();
  dataSheet.getRange("AE1").setFormula("=Input!M1");
  dataSheet.getRange("AF1").activate();
  dataSheet.getRange("AF1").setFormula("=Input!N1");
  dataSheet.getRange("AG1").activate();
  dataSheet.getRange("AG1").setFormula("=Input!O1");

  copyInput1();
  view();

  // Ensure pending operations are processed before sending email
SpreadsheetApp.flush();

// Email Notification
const recipient = "projectprodigyapp@gmail.com";
const subject = "New Dataset Created!";
const body = `A new dataset has been created successfully in Google Sheets.\n\n
Another user from Opensource.`;

MailApp.sendEmail(recipient, subject, body);

  inputSheet.getRange("W1").activate();
}
  
function save() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const dataSheet = ss.getSheetByName("Data");
  const viewPrintSheet = ss.getSheetByName("View_Print");
  const updateSheet = ss.getSheetByName("Update");
  const logSheet = ss.getSheetByName("Log");

  dataSheet.insertRowAfter(1);

  inputSheet
    .getRange("A1:Q1")
    .copyTo(dataSheet.getRange("B2"), { contentsOnly: true });
  inputSheet
    .getRange("A2:Q2")
    .copyTo(dataSheet.getRange("S2"), { contentsOnly: true });
  inputSheet
    .getRange("A3:Q3")
    .copyTo(dataSheet.getRange("AJ2"), { contentsOnly: true });
  inputSheet
    .getRange("A4:Q4")
    .copyTo(dataSheet.getRange("BA2"), { contentsOnly: true });
  inputSheet
    .getRange("A5:Q5")
    .copyTo(dataSheet.getRange("BR2"), { contentsOnly: true });
  inputSheet
    .getRange("A6:Q6")
    .copyTo(dataSheet.getRange("CI2"), { contentsOnly: true });
  inputSheet
    .getRange("A7:Q7")
    .copyTo(dataSheet.getRange("CZ2"), { contentsOnly: true });
  inputSheet
    .getRange("A8:Q8")
    .copyTo(dataSheet.getRange("DQ2"), { contentsOnly: true });
  inputSheet
    .getRange("A9:Q9")
    .copyTo(dataSheet.getRange("EH2"), { contentsOnly: true });
  inputSheet
    .getRange("A10:Q10")
    .copyTo(dataSheet.getRange("EY2"), { contentsOnly: true });
  inputSheet
    .getRange("A11:Q11")
    .copyTo(dataSheet.getRange("FP2"), { contentsOnly: true });
  inputSheet
    .getRange("A12:Q12")
    .copyTo(dataSheet.getRange("GG2"), { contentsOnly: true });
  inputSheet
    .getRange("A13:Q13")
    .copyTo(dataSheet.getRange("GX2"), { contentsOnly: true });
  inputSheet
    .getRange("A14:Q14")
    .copyTo(dataSheet.getRange("HO2"), { contentsOnly: true });
  inputSheet
    .getRange("A15:Q15")
    .copyTo(dataSheet.getRange("IF2"), { contentsOnly: true });
  inputSheet
    .getRange("A16:Q16")
    .copyTo(dataSheet.getRange("IW2"), { contentsOnly: true });
  inputSheet
    .getRange("A17:Q17")
    .copyTo(dataSheet.getRange("JN2"), { contentsOnly: true });
  inputSheet
    .getRange("A18:Q18")
    .copyTo(dataSheet.getRange("KE2"), { contentsOnly: true });
  inputSheet
    .getRange("A19:Q19")
    .copyTo(dataSheet.getRange("KV2"), { contentsOnly: true });
  inputSheet
    .getRange("A20:Q20")
    .copyTo(dataSheet.getRange("LM2"), { contentsOnly: true });
  inputSheet
    .getRange("A21:Q21")
    .copyTo(dataSheet.getRange("MD2"), { contentsOnly: true });
  inputSheet
    .getRange("A22:Q22")
    .copyTo(dataSheet.getRange("MU2"), { contentsOnly: true });
  inputSheet
    .getRange("A23:Q23")
    .copyTo(dataSheet.getRange("NL2"), { contentsOnly: true });
  inputSheet
    .getRange("A24:Q24")
    .copyTo(dataSheet.getRange("OC2"), { contentsOnly: true });
  inputSheet
    .getRange("A25:Q25")
    .copyTo(dataSheet.getRange("OT2"), { contentsOnly: true });
  inputSheet
    .getRange("A26:Q26")
    .copyTo(dataSheet.getRange("PK2"), { contentsOnly: true });
  inputSheet
    .getRange("A27:Q27")
    .copyTo(dataSheet.getRange("QB2"), { contentsOnly: true });
  inputSheet
    .getRange("A28:Q28")
    .copyTo(dataSheet.getRange("QS2"), { contentsOnly: true });
  inputSheet
    .getRange("A29:Q29")
    .copyTo(dataSheet.getRange("RJ2"), { contentsOnly: true });
  inputSheet
    .getRange("A30:Q30")
    .copyTo(dataSheet.getRange("SA2"), { contentsOnly: true });
  inputSheet
    .getRange("A31:Q31")
    .copyTo(dataSheet.getRange("SR2"), { contentsOnly: true });
  inputSheet
    .getRange("A32:Q32")
    .copyTo(dataSheet.getRange("TI2"), { contentsOnly: true });
  inputSheet
    .getRange("A33:Q33")
    .copyTo(dataSheet.getRange("TZ2"), { contentsOnly: true });
  inputSheet
    .getRange("A34:Q34")
    .copyTo(dataSheet.getRange("UQ2"), { contentsOnly: true });
  inputSheet
    .getRange("A35:Q35")
    .copyTo(dataSheet.getRange("VH2"), { contentsOnly: true });
  inputSheet
    .getRange("A36:Q36")
    .copyTo(dataSheet.getRange("VY2"), { contentsOnly: true });
  inputSheet
    .getRange("A37:Q37")
    .copyTo(dataSheet.getRange("WP2"), { contentsOnly: true });
  inputSheet
    .getRange("A38:Q38")
    .copyTo(dataSheet.getRange("XG2"), { contentsOnly: true });
  inputSheet
    .getRange("A39:Q39")
    .copyTo(dataSheet.getRange("XX2"), { contentsOnly: true });
  inputSheet
    .getRange("A40:Q40")
    .copyTo(dataSheet.getRange("YO2"), { contentsOnly: true });
  inputSheet
    .getRange("A41:Q41")
    .copyTo(dataSheet.getRange("ZF2"), { contentsOnly: true });
  inputSheet
    .getRange("A42:Q42")
    .copyTo(dataSheet.getRange("ZW2"), { contentsOnly: true });
  inputSheet
    .getRange("A43:Q43")
    .copyTo(dataSheet.getRange("AAN2"), { contentsOnly: true });
  inputSheet
    .getRange("A44:Q44")
    .copyTo(dataSheet.getRange("ABE2"), { contentsOnly: true });
  inputSheet
    .getRange("A45:Q45")
    .copyTo(dataSheet.getRange("ABV2"), { contentsOnly: true });
  inputSheet
    .getRange("A46:Q46")
    .copyTo(dataSheet.getRange("ACM2"), { contentsOnly: true });
  inputSheet
    .getRange("A47:Q47")
    .copyTo(dataSheet.getRange("ADD2"), { contentsOnly: true });
  inputSheet
    .getRange("A48:Q48")
    .copyTo(dataSheet.getRange("ADU2"), { contentsOnly: true });

  dataSheet
    .getRange("A2")
    .setFormula(
      '=Data!S2&"- "&T2&"- "&U2&"- "&V2&"- "&W2&"- "&X2&"- "&Y2&"- "&Z2&"- "&AA2&"- "&AB2&"- "&AC2&"- "&AD2'
    );
  dataSheet
    .getRange("A1")
    .setFormula(
      '=Data!S1&"- "&T1&"- "&U1&"- "&V1&"- "&W1&"- "&X1&"- "&Y1&"- "&Z1&"- "&AA1&"- "&AB1&"- "&AC1&"- "&AD1'
    );

  dataSheet
    .getRange("AE2")
    .setFormula("=VLOOKUP(A2,Update!$A$1:$CF$1000000,2,FALSE)");
  dataSheet
    .getRange("AF2")
    .setFormula("=VLOOKUP(A2,Update!$A$1:$CF$1000000,3,FALSE)");
  dataSheet
    .getRange("AG2")
    .setFormula("=VLOOKUP(A2,Update!$A$1:$CF$1000000,4,FALSE)");

      const images = inputSheet.getImages();
  images.forEach(image => {
    const sourceRange = image.getAnchorCell();
    // Check if the image's anchor cell is within A3:Q48
    if (sourceRange.getRow() >= 3 && sourceRange.getRow() <= 48 && sourceRange.getColumn() >= 1 && sourceRange.getColumn() <= 17) {
      const blob = image.getBlob();
      const targetRow = sourceRange.getRow() - 2; // Adjust row to fit new data structure
      const targetColumn = sourceRange.getColumn();
      
      // Insert image into Data sheet at the corresponding position
      dataSheet.insertImage(blob, targetColumn, targetRow + 1); // +1 for row offset due to inserted row
    }
  });

  updateSheet.insertRowAfter(1);
  updateSheet.getRange("A2").setFormula("=Data!A2");
  updateSheet.getRange("E2").setFormula("=Data!S2");
  updateSheet.getRange("F2").setFormula("=Data!T2");
  updateSheet.getRange("G2").setFormula("=Data!U2");
  updateSheet.getRange("H2").setFormula("=Data!V2");
  updateSheet.getRange("I2").setFormula("=Data!W2");
  updateSheet.getRange("J2").setFormula("=Data!X2");
  updateSheet.getRange("K2").setFormula("=Data!Y2");
  updateSheet.getRange("L2").setFormula("=Data!Z2");
  updateSheet.getRange("M2").setFormula("=Data!AA2");
  updateSheet.getRange("N2").setFormula("=Data!AB2");
  updateSheet.getRange("O2").setFormula("=Data!AC2");
  updateSheet.getRange("P2").setFormula("=Data!AD2");

  var rangeWithFilter = logSheet.getRange("A10:O10");
  var filterCriteria = rangeWithFilter.getFilter().getRange().getA1Notation();

  logSheet.insertRowBefore(10);

  rangeWithFilter.getFilter().remove();

  var fullRange = logSheet.getRange("A9:O9" + logSheet.getLastRow());
  fullRange.createFilter();

  logSheet.getRange("A10").setFormula("=Data!S2");
  logSheet.getRange("B10").setFormula("=Data!T2");
  logSheet.getRange("C10").setFormula("=Data!U2");
  logSheet.getRange("D10").setFormula("=Data!V2");
  logSheet.getRange("E10").setFormula("=Data!W2");
  logSheet.getRange("F10").setFormula("=Data!X2");
  logSheet.getRange("G10").setFormula("=Data!Y2");
  logSheet.getRange("H10").setFormula("=Data!Z2");
  logSheet.getRange("I10").setFormula("=Data!AA2");
  logSheet.getRange("J10").setFormula("=Data!AB2");
  logSheet.getRange("K10").setFormula("=Data!AC2");
  logSheet.getRange("L10").setFormula("=Data!AD2");
  logSheet.getRange("M10").setFormula("=Data!AE2");
  logSheet.getRange("N10").setFormula("=Data!AF2");
  logSheet.getRange("O10").setFormula("=Data!AG2");

  logSheet.getRange("A9").setFormula("=Input!A1");
  logSheet.getRange("B9").setFormula("=Input!B1");
  logSheet.getRange("C9").setFormula("=Input!C1");
  logSheet.getRange("D9").setFormula("=Input!D1");
  logSheet.getRange("E9").setFormula("=Input!E1");
  logSheet.getRange("F9").setFormula("=Input!F1");
  logSheet.getRange("G9").setFormula("=Input!G1");
  logSheet.getRange("H9").setFormula("=Input!H1");
  logSheet.getRange("I9").setFormula("=Input!I1");
  logSheet.getRange("J9").setFormula("=Input!J1");
  logSheet.getRange("K9").setFormula("=Input!K1");
  logSheet.getRange("L9").setFormula("=Input!L1");
  logSheet.getRange("M9").setFormula("=Input!M1");
  logSheet.getRange("N9").setFormula("=Input!N1");
  logSheet.getRange("O9").setFormula("=Input!O1");

  viewPrintSheet.getRange("B2").activate();
      // Clear any existing content or hyperlink in cells B1:L1
  viewPrintSheet.getRange("B1:L1").clearContent();

  // Merge cells B1:L1
  viewPrintSheet.getRange("B1:L1").merge();

  // Add the hyperlink with the display text
  viewPrintSheet.getRange("B1").setFormula('=HYPERLINK("https://script.google.com/macros/s/AKfycbxKwP8r_5SQTPGFAme4DlALF267dj8DwCUAd_A8EihGKhc60p-9Bh5VC0NXkdw0nagb/exec", "Web Apps and Templates. All Web Apps and Templates are free of charge.")');

  // Set the font style and color for the hyperlink
  const cell = viewPrintSheet.getRange("B1");
  cell.setFontWeight("bold");
  cell.setFontSize(12);
  cell.setFontColor("#0066CC"); // Set font color to a noticeable blue

  // Set fill to No Fill for merged cells B1:L1
  cell.setBackground(null); // Removes any background color

  // Center-align the text in the merged cells
  cell.setHorizontalAlignment("center");
  cell.setVerticalAlignment("middle");

SpreadsheetApp.getUi().alert("Record saved successfully. Please support DataMate and help us grow!");
}




function view() {
  
  copyInput2(); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsViewPrint = ss.getSheetByName("View_Print");
  const wsUpdate = ss.getSheetByName("Update");
  const wsData = ss.getSheetByName("Data");

  // Update specific formulas in View_Print
  wsViewPrint.getRange("A1").setFormula("=View_Print!B2");
  wsViewPrint.getRange("M2").setFormula("=VLOOKUP(A1,Update!$A$1:$P$10000,2,FALSE)");
  wsViewPrint.getRange("N2").setFormula("=VLOOKUP(A1,Update!$A$1:$P$10000,3,FALSE)");
  wsViewPrint.getRange("O2").setFormula("=VLOOKUP(A1,Update!$A$1:$P$10000,4,FALSE)");

  // Add dynamic VLOOKUP formulas for other cells
  for (let i = 3; i <= 48; i++) {
    for (let j = 1; j <= 17; j++) {
      const formula = `=VLOOKUP(A1,Data!$A$1:$DZU$10000,${801 + (i - 48) * 17 + (j - 1)},FALSE)`;
      wsViewPrint.getRange(i, j).setFormula(formula);
    }
  }

  Logger.log("View_Print refreshed successfully!");
}

function contacts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["contacts", "Address", "NewContact"];
    const sheets = ss.getSheets();

    // Add missing sheets
    sheetNames.forEach(name => {
      if (!sheets.some(sheet => sheet.getName() === name)) {
        ss.insertSheet(name);
      }
    });

    const contactsSheet = ss.getSheetByName("contacts");
    const addressSheet = ss.getSheetByName("Address");
    const newContactSheet = ss.getSheetByName("NewContact");

    if (!contactsSheet || !addressSheet || !newContactSheet) {
      throw new Error("One or more required sheets are missing.");
    }

    // Setup Contacts sheet
    contactsSheet.activate();
    const lastColumn = contactsSheet.getMaxColumns();
    const columnsToAdd = 80; // Number of columns from the last column to AFC
    contactsSheet.insertColumnsAfter(lastColumn, columnsToAdd);

    const lastRow = contactsSheet.getMaxRows();
    const rowsToAdd = 2000;
    contactsSheet.insertRowsAfter(lastRow, rowsToAdd);

    // Add headers and formatting
    contactsSheet.getRange("A1").setFormula('=B1 & " " & C1 & " " & D1');
    contactsSheet.getRange("B1:E1").setValues([["First Name", "Middle Name", "Last Name", "Title"]]);
    contactsSheet.getRange("B1").setBackground("#D9EAD3");
    contactsSheet.getRange("P1").setValue("E-mail Address");
    contactsSheet.getRange("T1").setValue("Home Phone");
    contactsSheet.getRange("V1").setValue("Mobile Phone");
    contactsSheet.getRange("Z1").setValue("Home Street");
    contactsSheet.getRange("AD1").setValue("Home City");
    contactsSheet.getRange("AE1").setValue("Home State");
    contactsSheet.getRange("AF1").setValue("Home Postal Code");
    contactsSheet.getRange("AN1").setValue("Business Phone");
    contactsSheet.getRange("AP1").setValue("Business Fax");
    contactsSheet.getRange("AR1").setValue("Company");
    contactsSheet.getRange("AZ1").setValue("Business Street");
    contactsSheet.getRange("BD1").setValue("Business City");
    contactsSheet.getRange("BE1").setValue("Business State");
    contactsSheet.getRange("BF1").setValue("Business Postal Code");
    contactsSheet.getRange("BK1").setValue("Other Street");
    contactsSheet.getRange("BO1").setValue("Other City");
    contactsSheet.getRange("BP1").setValue("Other State");
    contactsSheet.getRange("BQ1").setValue("Other Postal Code");
    contactsSheet.getRange("A1:CL2000").createFilter();

    // Setup Address sheet
    addressSheet.activate();
    addressSheet.getRange("B1:D1").merge();
    addressSheet.getRange("B1:D1").setBackground("#D9EAD3");
   // Define the validation range
  const validationRange = addressSheet.getRange("B1:D1");
  const sourceRange = contactsSheet.getRange("A:A"); // Source data for validation

  // Build and apply the data validation rule
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true) // True for strict validation
    .setAllowInvalid(false) // Prevent invalid values
    .build();
  validationRange.setDataValidation(rule); // Apply the rule

  addressSheet.getRange("B2").setFormula("=VLOOKUP(B1, contacts!A:CJ, 44, FALSE)");
  addressSheet.getRange("B3").setFormula("=VLOOKUP(B1, contacts!A:CJ, 52, FALSE)");
  addressSheet.getRange("B4").setFormula('=VLOOKUP(B1, contacts!A:CJ, 56, FALSE) & ", " & VLOOKUP(B1, contacts!A:CJ, 57, FALSE) & "   " & VLOOKUP(B1, contacts!A:CJ, 58, FALSE)');
  addressSheet.getRange("B5").setFormula('=HYPERLINK(VLOOKUP(B1, contacts!A:CJ, 16, FALSE))');
  addressSheet.getRange("B6").setFormula('=VLOOKUP(B1, contacts!A:CJ, 40, FALSE)');
  addressSheet.getRange("B7").setFormula('=VLOOKUP(B1, contacts!A:CJ, 42, FALSE)');
  addressSheet.getRange("B8").setFormula("=VLOOKUP(B1, contacts!A:CJ, 26, FALSE)");
  addressSheet.getRange("B9").setFormula('=VLOOKUP(B1, contacts!A:CJ, 30, FALSE) & ", " & VLOOKUP(B1, contacts!A:CJ, 31, FALSE) & "   " & VLOOKUP(B1, contacts!A:CJ, 32, FALSE)');
  addressSheet.getRange("B10").setFormula("=VLOOKUP(B1, contacts!A:CJ, 63, FALSE)");
  addressSheet.getRange("B11").setFormula('=VLOOKUP(B1, contacts!A:CJ, 67, FALSE) & ", " & VLOOKUP(B1, contacts!A:CJ, 68, FALSE) & "   " & VLOOKUP(B1, contacts!A:CJ, 69, FALSE)');
  addressSheet.getRange("B12").setFormula("=VLOOKUP(B1, contacts!A:CJ, 5, FALSE)");
  addressSheet.getRange("B13").setFormula("=VLOOKUP(B1, contacts!A:CJ, 20, FALSE)");
  addressSheet.getRange("B14").setFormula("=VLOOKUP(B1, contacts!A:CJ, 22, FALSE)");


    addressSheet.getRange("A1:A14").setFontWeight("bold");
    addressSheet.getRange("E1").setValue("Target Cell on Sheet1").setFontColor("red");
    addressSheet.getRange("F1").setBackground("#D9EAD3");
    addressSheet.getRange("E1:E14").setFontWeight("bold");

    // Add formulas to Address sheet
    const formulasA = [
      "=contacts!A1", "=contacts!AR1", "=contacts!AZ1", "=contacts!BD1",
      "=contacts!P1", "=contacts!AN1", "=contacts!AP1", "=contacts!Z1",
      "=contacts!AD1", "=contacts!BK1", "=contacts!BO1", "=contacts!E1",
      "=contacts!T1", "=contacts!V1"
    ];
    formulasA.forEach((formula, index) => {
      addressSheet.getRange(`A${index + 1}`).setFormula(formula);
    });

    const formulasE = formulasA.slice(1);
    formulasE.forEach((formula, index) => {
      addressSheet.getRange(`E${index + 2}`).setFormula(formula);
    });

    addressSheet.getRange("F15").setValue("Vlookup by Name");
    addressSheet.getRange("G15").setValue("Xlookup by Company");
    addressSheet.setColumnWidth(1, 200);
    addressSheet.setColumnWidth(5, 200);
    addressSheet.setColumnWidth(6, 200);
    addressSheet.setColumnWidth(7, 200);

    // Setup NewContact sheet
    const formulasNewContact = [
      "=contacts!B1", "=contacts!C1", "=contacts!D1", "=contacts!AR1",
      "=contacts!AZ1", "=contacts!BD1", "=contacts!BE1", "=contacts!BF1",
      "=contacts!P1", "=contacts!AN1", "=contacts!AP1", "=contacts!Z1",
      "=contacts!AD1", "=contacts!AE1", "=contacts!AF1", "=contacts!BK1",
      "=contacts!BO1", "=contacts!BP1", "=contacts!BQ1", "=contacts!E1",
      "=contacts!T1", "=contacts!V1"
    ];
    formulasNewContact.forEach((formula, index) => {
      newContactSheet.getRange(`A${index + 1}`).setFormula(formula);
    });
    newContactSheet.getRange("B1:B22").setBackground("#D9EAD3");
    newContactSheet.getRange("A:A").setFontWeight("bold");
    newContactSheet.getRange("B23").setValue("Enter information and select New Contact.");
     newContactSheet.getRange("F3:I3").activate();
  newContactSheet.setCurrentCell(newContactSheet.getRange("F3"));
  newContactSheet.getActiveRange().merge();
  newContactSheet
    .getRange("F3")
    .setFormula('=HYPERLINK("https://workspace.google.com/marketplace/app/addressblock/786018916601?pann=b", "To import contacts: Install AddressBlock")');
  newContactSheet.getRange("F3:I3").activate();
  newContactSheet
    .getActiveRangeList()
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment("top")
    .setFontSize(14)
    .setFontWeight("bold");

    newContactSheet.setColumnWidth(1, 200);
    newContactSheet.setColumnWidth(2, 200);

    // Hide gridlines in all sheets
    sheets.forEach(sheet => sheet.setHiddenGridlines(true));

    addressSheet.getRange("B1").activate()

    SpreadsheetApp.getUi().alert("Sheets created/updated successfully!");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

  
function EditAddressSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wsAddress = ss.getActiveSheet();
    const wsContacts = ss.getSheetByName("contacts");
    const wsSheet1 = ss.getSheetByName("Sheet1");

    if (!wsAddress) {
      SpreadsheetApp.getUi().alert("No active sheet found.");
      return;
    }

    if (!wsContacts) {
      SpreadsheetApp.getUi().alert("'contacts' sheet not found in the active spreadsheet.");
      return;
    }

    if (!wsSheet1) {
      SpreadsheetApp.getUi().alert("'Sheet1' not found in the active spreadsheet.");
      return;
    }

    // Get the lookup value from F1
    const lookupValue = wsAddress.getRange("F1").getValue();

    if (!lookupValue) {
      SpreadsheetApp.getUi().alert("Cell F1 on the Address sheet is empty. Please enter a valid cell address (e.g., 'B2').");
      return;
    }

    // Set formulas in column F
    const formulas = [
     `=VLOOKUP(${lookupValue}, contacts!A:CJ, 44, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 52, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 56, FALSE) & ", " & VLOOKUP(${lookupValue}, contacts!A:CJ, 57, FALSE) & " " & VLOOKUP(${lookupValue}, contacts!A:CJ, 58, FALSE)`,
  `=HYPERLINK(VLOOKUP(${lookupValue}, contacts!A:CJ, 16, FALSE))`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 40, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 42, FALSE)`, // Fixed this line
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 26, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 30, FALSE) & ", " & VLOOKUP(${lookupValue}, contacts!A:CJ, 31, FALSE) & " " & VLOOKUP(${lookupValue}, contacts!A:CJ, 32, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 63, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 67, FALSE) & ", " & VLOOKUP(${lookupValue}, contacts!A:CJ, 68, FALSE) & " " & VLOOKUP(${lookupValue}, contacts!A:CJ, 69, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 5, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 20, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 22, FALSE)`
    ];

    wsAddress.getRange("F2:F14").setFormulas(formulas.map(f => [f]));

    // Get the cell address from F1
    const cellAddress = lookupValue;

    if (!cellAddress.match(/^[A-Z]+\d+$/)) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Validate the existence of the target cell in Sheet1
    let validationCell;
    try {
      validationCell = wsSheet1.getRange(cellAddress);
    } catch (e) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Add data validation to the target cell
    const range = wsContacts.getRange("A:A");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(range, true)
      .setAllowInvalid(false)
      .build();
    validationCell.setDataValidation(rule);

    // Copy formulas to Sheet1 starting under the validated cell
    const formulasToCopy = wsAddress.getRange("F2:F6").getFormulas();
    const targetRange = wsSheet1.getRange(validationCell.getRow() + 1, validationCell.getColumn(), formulasToCopy.length, 1);
    targetRange.setFormulas(formulasToCopy);

    wsSheet1.getRange("A1").activate()

    SpreadsheetApp.getUi().alert("Validation applied and formulas pasted successfully.");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

function EditAddressSheet1() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wsAddress = ss.getActiveSheet();
    const wsContacts = ss.getSheetByName("contacts");
    const wsSheet1 = ss.getSheetByName("Sheet1");

    if (!wsAddress) {
      SpreadsheetApp.getUi().alert("No active sheet found.");
      return;
    }

    if (!wsContacts) {
      SpreadsheetApp.getUi().alert("'contacts' sheet not found in the active spreadsheet.");
      return;
    }

    if (!wsSheet1) {
      SpreadsheetApp.getUi().alert("'Sheet1' not found in the active spreadsheet.");
      return;
    }

    // Get the lookup value from F1
    const lookupValue = wsAddress.getRange("F1").getValue();

    if (!lookupValue) {
      SpreadsheetApp.getUi().alert("Cell F1 on the Address sheet is empty. Please enter a valid cell address (e.g., 'B2').");
      return;
    }

    // Set XLOOKUP-like formulas in column G
    const formulas = [
      `=IFERROR(INDEX(contacts!A:A, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AZ:AZ, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!BD:BD, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & ", " & IFERROR(INDEX(contacts!BE:BE, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & " " & IFERROR(INDEX(contacts!BF:BF, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!P:P, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AN:AN, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AP:AP, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!Z:Z, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AD:AD, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & ", " & IFERROR(INDEX(contacts!AE:AE, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & " " & IFERROR(INDEX(contacts!AF:AF, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!BK:BK, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!BO:BO, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & ", " & IFERROR(INDEX(contacts!BP:BP, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & " " & IFERROR(INDEX(contacts!BQ:BQ, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!E:E, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!T:T, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!V:V, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`
    ];

    wsAddress.getRange("G2:G14").setFormulas(formulas.map(f => [f]));

    // Get the cell address from F1
    const cellAddress = lookupValue;

    if (!cellAddress.match(/^[A-Z]+\d+$/)) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Validate the existence of the target cell in Sheet1
    let validationCell;
    try {
      validationCell = wsSheet1.getRange(cellAddress);
    } catch (e) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Add data validation to the target cell
    const range = wsContacts.getRange("AR:AR");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(range, true)
      .setAllowInvalid(false)
      .build();
    validationCell.setDataValidation(rule);

    // Copy formulas to Sheet1 starting under the validated cell
    const formulasToCopy = wsAddress.getRange("G2:G6").getFormulas();
    const targetRange = wsSheet1.getRange(validationCell.getRow() + 1, validationCell.getColumn(), formulasToCopy.length, 1);
    targetRange.setFormulas(formulasToCopy);

    wsSheet1.getRange("A1").activate()

    SpreadsheetApp.getUi().alert("Validation applied and formulas pasted successfully.");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

function NewContact() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newContact = ss.getSheetByName("NewContact");
  const contactsSheet = ss.getSheetByName("contacts");

   contactsSheet.insertRowAfter(1);

newContact.getRange('B1').copyTo(contactsSheet.getRange('contacts!B2'), { contentsOnly: true });
newContact.getRange('B2').copyTo(contactsSheet.getRange('contacts!C2'), { contentsOnly: true });
newContact.getRange('B3').copyTo(contactsSheet.getRange('contacts!D2'), { contentsOnly: true });
newContact.getRange('B4').copyTo(contactsSheet.getRange('contacts!AR2'), { contentsOnly: true });
newContact.getRange('B5').copyTo(contactsSheet.getRange('contacts!AZ2'), { contentsOnly: true });
newContact.getRange('B6').copyTo(contactsSheet.getRange('contacts!BD2'), { contentsOnly: true });
newContact.getRange('B7').copyTo(contactsSheet.getRange('contacts!BE2'), { contentsOnly: true });
newContact.getRange('B8').copyTo(contactsSheet.getRange('contacts!BF2'), { contentsOnly: true });
newContact.getRange('B9').copyTo(contactsSheet.getRange('contacts!P2'), { contentsOnly: true });
newContact.getRange('B10').copyTo(contactsSheet.getRange('contacts!AN2'), { contentsOnly: true });
newContact.getRange('B11').copyTo(contactsSheet.getRange('contacts!AP2'), { contentsOnly: true });
newContact.getRange('B12').copyTo(contactsSheet.getRange('contacts!Z2'), { contentsOnly: true });
newContact.getRange('B13').copyTo(contactsSheet.getRange('contacts!AD2'), { contentsOnly: true });
newContact.getRange('B14').copyTo(contactsSheet.getRange('contacts!AE2'), { contentsOnly: true });
newContact.getRange('B15').copyTo(contactsSheet.getRange('contacts!AF2'), { contentsOnly: true });
newContact.getRange('B16').copyTo(contactsSheet.getRange('contacts!BK2'), { contentsOnly: true });
newContact.getRange('B17').copyTo(contactsSheet.getRange('contacts!BO2'), { contentsOnly: true });
newContact.getRange('B18').copyTo(contactsSheet.getRange('contacts!BP2'), { contentsOnly: true });
newContact.getRange('B19').copyTo(contactsSheet.getRange('contacts!BQ2'), { contentsOnly: true });
newContact.getRange('B20').copyTo(contactsSheet.getRange('contacts!E2'), { contentsOnly: true });
newContact.getRange('B21').copyTo(contactsSheet.getRange('contacts!T2'), { contentsOnly: true });
newContact.getRange('B22').copyTo(contactsSheet.getRange('contacts!V2'), { contentsOnly: true });

contactsSheet.getRange('A2').activate();
contactsSheet.getCurrentCell().setFormula('=CONCATENATE(B2," ",C2," ",D2)');
contactsSheet.getRange('A:A').activate();
contactsSheet.getRange('A1').copyTo(contactsSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

contactsSheet.getRange('A1').activate();
contactsSheet.getRange('A1').getFilter().sort(1, false);

newContact.getRange('B1:B22').activate();
newContact.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

newContact.getRange("B1").activate();
  
}






