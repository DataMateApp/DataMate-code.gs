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
    .addSeparator() // Extra separator to space out the label
    .addItem("➡ Or start with a template ⬅", "doNothing") // Dummy function
    .addSeparator()
    .addItem("Inventory Template", "setup")
    .addSeparator()
    .addItem("Update Inventory", "updateInventory")
    .addSeparator()
    .addItem("Purchase Order Template", "setupPO")
    .addSeparator()
    .addItem("Weekly Timesheets Template", "setupTS")
    .addSeparator()
    .addItem("Update Cost Code Totals", "copyToCodeTotals")
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



function setupTS() {
  createTimeSheet();
  newfile();
  cleanupTS();
  
}

function cleanupTS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const logSheet = ss.getSheetByName("Log");

  const mappings = [
    ["A1", "=J5"], ["A2", "=J6"], ["B1", "=B3"], ["B2", "=B4"],
    ["C1", "=A5"], ["C2", "=A6"], ["D1", "=A41"], ["D2", "=P43"],
    ["E1", "=P41"], ["E2", "=P42"], ["F1", "=Q41"], ["F2", "=Q42"],
    ["G1", "=A45"], ["G2", "=B45"], ["H1", "=E45"], ["H2", "=F45"],
    ["I1", "=I45"],["I2", "=J45"], ["J1", "=M45"], ["J2", "=O45"], ["K1", "Log 11"], ["L1", "Log 12"],
    ["M1", "Update 1"], ["N1", "Update 2"], ["O1", "Update 3"]
  ];

  mappings.forEach(([cell, value]) => {
    inputSheet.getRange(cell).setValue(value);
  });

  logSheet.getRange("A2").setValue("Time/Cost Log");
  logSheet.getRange("C7").setValue("TOTALS");
  logSheet.getRange("D7").setFormula("=SUM(D9:D)");
  logSheet.getRange("E7").setFormula("=SUM(E9:E)");
  logSheet.getRange("F7").setFormula("=SUM(F9:F)");
  logSheet.getRange("G7").setFormula("=SUM(G9:G)");
  logSheet.getRange("H7").setFormula("=SUM(H9:H)");
  logSheet.getRange("I7").setFormula("=SUM(I9:I)");
  logSheet.getRange("J7").setFormula("=SUM(J9:J)");
}

function copyToCodeTotals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Input");
  var totalsSheet = ss.getSheetByName("Code Totals");

  if (!inputSheet || !totalsSheet) {
    Logger.log("One or both sheets are missing.");
    return;
  }

  // Get data from the Input sheet
  var costCodes = inputSheet.getRange("A9:A40").getValues(); // Cost codes (double rows)
  var otHours = inputSheet.getRange("P9:P40").getValues();   // OT hours
  var dtHours = inputSheet.getRange("Q9:Q40").getValues();   // DT hours
  var regHoursP = inputSheet.getRange("P10:P40").getValues(); // Regular hours from P
  var regHoursQ = inputSheet.getRange("Q10:Q40").getValues(); // Regular hours from Q

  // Get existing data from Code Totals
  var totalsData = totalsSheet.getRange("A2:D" + totalsSheet.getLastRow()).getValues();
  var totalsMap = {}; // Store existing codes and their row index

  // Map existing cost codes to their row index
  totalsData.forEach((row, index) => {
    if (row[0]) totalsMap[row[0]] = index + 2; // Row index (considering header row)
  });

  for (var i = 0; i < costCodes.length; i += 2) { // Process in pairs (double rows)
    var code = costCodes[i][0]; // Cost code in A9, A11, etc.
    if (!code) continue; // Skip empty rows

    var ot = otHours[i][0] || 0;
    var dt = dtHours[i][0] || 0;
    var reg = (regHoursP[i][0] || 0) + (regHoursQ[i][0] || 0); // Sum regular hours correctly

    if (code in totalsMap) {
      // Get the existing row index
      var rowIndex = totalsMap[code];

      // Fetch current values from "Code Totals" before updating
      var currentReg = totalsSheet.getRange(rowIndex, 2).getValue() || 0;
      var currentOT = totalsSheet.getRange(rowIndex, 3).getValue() || 0;
      var currentDT = totalsSheet.getRange(rowIndex, 4).getValue() || 0;

      // Update the values by adding to the existing ones
      totalsSheet.getRange(rowIndex, 2).setValue(currentReg + reg); // Regular Hours
      totalsSheet.getRange(rowIndex, 3).setValue(currentOT + ot);   // OT Hours
      totalsSheet.getRange(rowIndex, 4).setValue(currentDT + dt);   // DT Hours
    } else {
      // Append new row if cost code doesn't exist
      var nextRow = totalsSheet.getLastRow() + 1;
      totalsSheet.getRange(nextRow, 1, 1, 4).setValues([[code, reg, ot, dt]]);
    }
  }
}



function createTimeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let listsSheet = ss.getSheetByName("Lists");
  
  // Check if Lists sheet exists, create it if it doesn't
  if (!listsSheet) {
    // Create a new Lists sheet
    listsSheet = ss.insertSheet("Lists");

    // Define headers
    const headers = ["Name", "Emp. No", "Rate", "", "Crafts", "Cost Codes"];
    listsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Define sample data
    const data = ["Moe", 3000, "$43.68", "", "Carpenter", "13100   Project Superintendent"];
    listsSheet.getRange(2, 1, 1, data.length).setValues([data]);

    listsSheet.getRange("H1")
      .setFormula('=HYPERLINK("https://datamateapp.github.io/About%20Timesheet.html", "About Timesheet")');

    // Auto-size columns for better readability
    listsSheet.autoResizeColumns(1, headers.length);
  }

  // Check if Code Totals sheet exists, create it if it doesn't
  let codeSheet = ss.getSheetByName("Code Totals");
  if (!codeSheet) {
    // Create a new Code Totals sheet
    codeSheet = ss.insertSheet("Code Totals");

    // Define headers
    const headers = ["Cost Code", "Regular Hours", "OT Hours", "DT Hours"];
    codeSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Auto-size columns for better readability
    codeSheet.autoResizeColumns(1, headers.length);
  }



  
  const existingSheet = ss.getSheetByName("Sheet1");
  if (existingSheet) ss.deleteSheet(existingSheet);

  const sheet = ss.insertSheet("Sheet1");
  sheet.getRange("A3:Q48").clear();

  // Set column widths
  const columnWidths = [300, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30];
  columnWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  // Merge and set headers
  sheet.getRange("A3").setFontWeight("bold").setValue("EMPLOYEE NO.");
  sheet.getRange("A4").setHorizontalAlignment("center");



  sheet.getRange("B3:G3").merge().setFontWeight("bold").setValue("EMPLOYEE NAME");
  sheet.getRange("B4:G4").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("B6:G6").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("H3:I3").merge().setFontWeight("bold").setValue("Note:");
  sheet.getRange("J3:Q3").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("J4:M4").merge().setFontWeight("bold").setValue("PREPAID CHECK #");
  sheet.getRange("O4:Q4").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("J5:M5").merge().setFontWeight("bold").setValue("BEGINNING DATE");
  sheet.getRange("J6:M6").merge().setValue("=TODAY()").setBorder(true, true, true, true, true, true);

  sheet.getRange("A5").setFontWeight("bold").setValue("RATE");
  sheet.getRange("A6").setHorizontalAlignment("center");
  sheet.getRange("B6:G6").merge();
  sheet.getRange("B5:G5").merge().setFontWeight("bold").setValue("CRAFT");
  sheet.getRange("A7").setValue("COST CODE");

  // Generate dates and days
  for (let i = 0; i < 7; i++) {
    let col = 2 + (i * 2); // Start from column B (2) and increment by 2
    let dateCell = sheet.getRange(7, col);
    dateCell.setFormula(`=IF(J6="", "", J6+${i})`);
    dateCell.offset(1, 0).setFormula(`=TEXT(${dateCell.getA1Notation()}, "ddd")`);
  }

  sheet.getRange("P7").setFontWeight("bold").setValue("TOTAL");

  // Merge formatting for table header
  const mergeRanges = [
    "A7:A8", "B7:C7", "D7:E7", "F7:G7", "H7:I7", "J7:K7", "L7:M7", "N7:O7",
    "B8:C8", "D8:E8", "F8:G8", "H8:I8", "J8:K8", "L8:M8", "N8:O8", "P7:Q8"
  ];
  mergeRanges.forEach(range => sheet.getRange(range).merge());

  // Set formulas for totals
  sheet.getRange("P9").setFormula('=SUM(B9+D9+F9+H9+J9+L9+N9)');
  sheet.getRange("Q9").setFormula('=SUM(C9+E9+G9+I9+K9+M9+O9)');
  sheet.getRange("P10:Q10").setFormula('=SUM(B10+D10+F10+H10+J10+L10+N10)');


// Get last row with data in column F
let lastRow = listsSheet.getLastRow();
let costCodeRange = listsSheet.getRange("F1:F2000" + lastRow);

// Create data validation rule (Dropdown from range)
const costCodeValidation = SpreadsheetApp.newDataValidation()
  .requireValueInRange(costCodeRange, true) // `true` ensures it's a dynamic range
  .setAllowInvalid(false)
  .build();

// Apply to target range
sheet.getRange("A9:A10").merge().setDataValidation(costCodeValidation);



  const cellMerges = ["B10:C10", "D10:E10", "F10:G10", "H10:I10", "J10:K10", "L10:M10", "N10:O10", "P10:Q10"];
  cellMerges.forEach(range => sheet.getRange(range).merge());

  // Copy and paste rows 9 & 10 to target row pairs
  const targetRows = [11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39];

  targetRows.forEach(row => {
    let sourceRange = sheet.getRange("A9:Q10");
    let destinationRange = sheet.getRange(row, 1, 2, 17); // Two-row range at target row
    sourceRange.copyTo(destinationRange);
  });

  // Apply borders to table
  sheet.getRange("A7:Q41").setBorder(true, true, true, true, true, true);

  sheet.getRange("A41:A43").merge().setValue("TOTAL HOURS").setHorizontalAlignment("center").setFontWeight("bold").setBorder(true, true, true, true, true, true);
  sheet.getRange("P41").setFontWeight("bold").setValue("OT")
  sheet.getRange("Q41").setFontWeight("bold").setValue("DT")
  sheet.getRange("B42:C43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("D42:E43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("F42:G43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("H42:I43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("J42:K43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("L42:M43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("N42:O43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("P43:Q43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("P42").setBorder(true, true, true, true, true, true).setFormula('=SUM(P9+P11+P13+P15+P17+P19+P21+P23+P25+P27+P29+P31+P33+P35+P37+P39)');
  sheet.getRange("Q42").setBorder(true, true, true, true, true, true).setFormula('=SUM(Q9+Q11+Q13+Q15+Q17+Q19+Q21+Q23+Q25+Q27+Q29+Q31+Q33+Q35+Q37+Q39)');
  sheet.getRange("P43:Q43").setFormula('=SUM(P10+P12+P14+P16+P18+P20+P22+P24+P26+P28+P30+P32+P34+P36+P38+P40)');
  
    // Get last row with data in column A
let nameRange = listsSheet.getRange("A1:A2000" + lastRow); // Correctly references the range

// Create data validation rule (Dropdown from range)
const nameValidation = SpreadsheetApp.newDataValidation()
  .requireValueInRange(nameRange, true) // `true` ensures it's a dynamic range
  .setAllowInvalid(false)
  .build();

// Apply validation to B4:G4
sheet.getRange("B4:G4").setDataValidation(nameValidation);



   
  // Get last row with data in column E
let craftRange = listsSheet.getRange("E1:E2000" + lastRow); // Correctly references the range

// Create data validation rule (Dropdown from range)
const craftValidation = SpreadsheetApp.newDataValidation()
  .requireValueInRange(craftRange, true) // `true` ensures it's a dynamic range
  .setAllowInvalid(false)
  .build();

// Apply validation to B6:G6
sheet.getRange("B6:G6").setDataValidation(craftValidation);


  sheet.getRange("A4").setFormula('=VLOOKUP(B4, Lists!A:B, 2, FALSE)');
  sheet.getRange("A6").setFormula('=VLOOKUP(B4, Lists!A:C, 3, FALSE)');

  
  sheet.getRange("A45").setHorizontalAlignment("right").setFontWeight("bold").setValue("Regular")
  sheet.getRange("B45:D45").merge().setHorizontalAlignment("center").setFormula('=SUM(P43*A6)')
  sheet.getRange("E45").setFontWeight("bold").setValue("OT")
  sheet.getRange("F45:H45").merge().setHorizontalAlignment("center").setFormula('=SUM(P42*A6*1.5)')
  sheet.getRange("I45").setFontWeight("bold").setValue("DT")
  sheet.getRange("J45:L45").merge().setHorizontalAlignment("center").setFormula('=SUM(Q42*A6*2)')
  sheet.getRange("M45:N45").merge().setFontWeight("bold").setValue("GROSS")
  sheet.getRange("O45:Q45").merge().setHorizontalAlignment("center").setFormula('=SUM(B45+F45+J45)')

   // Email Notification
const recipient = "projectprodigyapp@gmail.com";
const subject = "Timesheet Template Created!";
const body = `A new Timesheet Template has been created successfully in Google Sheets.\n\n
Another user from Opensource.`;

MailApp.sendEmail(recipient, subject, body);

}

function setupPO() {
  contacts();
  createPurchaseOrderTemplate();
  newfile();
  cleanup();
  
}

function cleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const logSheet = ss.getSheetByName("Log");

  const mappings = [
    ["A1", "=A6"], ["A2", "=C6"], ["B1", "=F6"], ["B2", "=H6"],
    ["C1", "=A8"], ["C2", "=B8"], ["D1", "=A20"], ["D2", "=B20"],
    ["E1", "=F47"], ["E2", "=G47"], ["F1", "=H37"], ["F2", "=J37"],
    ["G1", "=A23"], ["G2", "=B23"], ["H1", "=F39"], ["H2", "=G39"],
    ["I1", "Log 9"], ["J1", "Log 10"], ["K1", "Log 11"], ["L1", "Log 12"],
    ["M1", "=A44"], ["N1", "=A47"], ["O1", "Update 3"]
  ];

  mappings.forEach(([cell, value]) => {
    inputSheet.getRange(cell).setValue(value);
  });

  logSheet.getRange("A2").setValue("Purchase Order Log");
}


function createPurchaseOrderTemplate() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.insertSheet("Sheet1");
    sheet.clear(); // Clear previous content

    // Set column widths to match the layout
    sheet.setColumnWidth(1, 130); // A
    sheet.setColumnWidth(6, 130); // F


    var contactsSheet = ss.getSheetByName("contacts");
    var contactsListRange = contactsSheet.getRange("AR:AR");

    // Apply dropdown validation for "To:" and "Ship To:"
    var validation = SpreadsheetApp.newDataValidation().requireValueInRange(contactsListRange).build();
    sheet.getRange("B6:D6").merge().setFontWeight("bold").setDataValidation(validation); // "To:" dropdown
    sheet.getRange("B6").setValue("Company");
    sheet.getRange("G6:I6").merge().setFontWeight("bold").setDataValidation(validation); // "Ship To:" dropdown
    sheet.getRange("G6").setValue("Company");
    sheet.getRange("G15:I15").merge().setFontWeight("bold").setDataValidation(validation); // "Invoice To:" dropdown
    sheet.getRange("G15").setValue("Company");

    // Set headers and structure
    sheet.getRange("A1:J1").merge()
        .setValue("Your Company Name")
        .setFontWeight("bold")
        .setFontSize(20)
        .setHorizontalAlignment("center");
    sheet.getRange("A2:E2").merge()
        .setValue("Address")
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("center");
    sheet.getRange("F2:J2").merge()
        .setValue("City, State Zip")
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("center");

    sheet.getRange("A1:J46").setBorder(true, true, true, true, false, false);
    sheet.getRange("A3:J3").merge().setBackground('#cccccc').setBorder(true, true, true, true, false, false);
    sheet.getRange("F4:J17").setBorder(true, true, true, true, false, false);


    sheet.getRange("A4:B4").merge().setFontWeight("bold")
        .setValue("PURCHASE ORDER NUMBER:");
    sheet.getRange("F4:G4").merge().setFontWeight("bold").setValue("JOB/PHASE NUMBER:");

    sheet.getRange("A6").setValue("To:");
    sheet.getRange("B6").setFontWeight("bold"); // Company name (dropdown)

    sheet.getRange("F6").setValue("Ship To:");
    sheet.getRange("H6").setFontWeight("bold"); // Ship To company (dropdown)

    // Company Details
    sheet.getRange("B7:D7").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AZ:AZ, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("B8:D8").merge().setFontWeight("bold").setFormula(
  '=IFERROR(INDEX(contacts!BD:BD, MATCH(B6, contacts!AR:AR, 0)), "") & ", " & ' + 
  'IFERROR(INDEX(contacts!BE:BE, MATCH(B6, contacts!AR:AR, 0)), "") & " " & ' +
  'IFERROR(INDEX(contacts!BF:BF, MATCH(B6, contacts!AR:AR, 0)), "Not Found")'
);

    sheet.getRange("A9").setValue("Attn:");
    sheet.getRange("B9:D9").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!A:A, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");

    sheet.getRange("G7:I7").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AZ:AZ, MATCH(G6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("G8:I8").merge().setFontWeight("bold").setFormula(
  '=IFERROR(INDEX(contacts!BD:BD, MATCH(G6, contacts!AR:AR, 0)), "") & ", " & ' + 
  'IFERROR(INDEX(contacts!BE:BE, MATCH(G6, contacts!AR:AR, 0)), "") & " " & ' +
  'IFERROR(INDEX(contacts!BF:BF, MATCH(G6, contacts!AR:AR, 0)), "Not Found")'
);

    sheet.getRange("F10").setValue("Delivery-Site Phone:");
    sheet.getRange("G10:I10").merge().setFontWeight("bold").setValue("555-5555");

    sheet.getRange("F11").setValue("Site Contact:");
    sheet.getRange("G11:I11").merge().setFontWeight("bold").setValue("Joe Blow");

    sheet.getRange("A12").setValue("Phone:");
    sheet.getRange("B12:D12").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AN:AN, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("A13").setValue("Fax:");
    sheet.getRange("B13:D13").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AP:AP, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("A14").setValue("Email:");
    sheet.getRange("B14:D14").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!P:P, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");

    sheet.getRange("A15").setValue("Ship VIA:");
    sheet.getRange("B15").setValue("Best Route");

    sheet.getRange("F15").setValue("Send Invoices To:");
    sheet.getRange("G16:I16").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AZ:AZ, MATCH(G15, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("G17:I17").merge().setFontWeight("bold").setFormula(
  '=IFERROR(INDEX(contacts!BD:BD, MATCH(G15, contacts!AR:AR, 0)), "") & ", " & ' + 
  'IFERROR(INDEX(contacts!BE:BE, MATCH(G15, contacts!AR:AR, 0)), "") & " " & ' +
  'IFERROR(INDEX(contacts!BF:BF, MATCH(G15, contacts!AR:AR, 0)), "Not Found")'
);

    sheet.getRange("A18").setValue("Delivery Required By:");
    sheet.getRange("B18").setValue("=TODAY()");
    sheet.getRange("E18").setValue("F.O.B.:");
    sheet.getRange("F18:G18").merge().setBorder(true, true, true, true, false, false).setValue("Delivery Site");
    sheet.getRange("H18:I18").merge().setBorder(true, true, true, true, false, false).setValue("SALES TAX EXEMPT:");
    sheet.getRange("J18").setBorder(true, true, true, true, false, false).setValue("YES");
    sheet.getRange("A19:J19").merge().setBackground('#cccccc').setBorder(true, true, true, true, false, false);

    // Description of Materials Section
    sheet.getRange("A20:G20").merge().setBorder(true, true, true, true, false, false).setValue("DESCRIPTION OF MATERIALS").setFontWeight("bold");
    sheet.getRange("H20").setBorder(true, true, true, true, false, false).setValue("Unit Price").setFontWeight("bold");
    sheet.getRange("I20").setBorder(true, true, true, true, false, false).setValue("Quantity").setFontWeight("bold");
    sheet.getRange("J20").setBorder(true, true, true, true, false, false).setValue("Amount").setFontWeight("bold");

    sheet.getRange("A21").setBorder(true, true, true, true, false, false).setValue("PROJECT:");
    sheet.getRange("B21:G21").merge().setBorder(true, true, true, true, false, false).setValue("Project Name").setFontStyle("italic");

    sheet.getRange("A22:G22").merge().setBorder(true, true, true, true, false, false).setValue("Fabricate and furnish the following materials per the plans and specifications prepared by");
    sheet.getRange("A23:G23").merge().setBorder(true, true, true, true, false, false).setValue("A/E Name").setFontWeight("bold");
    sheet.getRange("A24:G24").merge().setBorder(true, true, true, true, false, false).setValue("dated --/--/---- and all applicable addenda and correspondence").setFontStyle("italic");
    sheet.getRange("A25:G25").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A26:G26").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A27:G27").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A28:G28").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A29:G29").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A30:G30").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A31:G31").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");

    // Move "Unit Price, Quantity, Amount" section below
    sheet.getRange("H25").setBorder(true, true, true, true, false, false)
    sheet.getRange("I25").setBorder(true, true, true, true, false, false)
    sheet.getRange("J25").setFormula("=H25*I25");

    sheet.getRange("H26").setBorder(true, true, true, true, false, false)
    sheet.getRange("I26").setBorder(true, true, true, true, false, false)
    sheet.getRange("J26").setFormula("=H26*I26");

    sheet.getRange("H27").setBorder(true, true, true, true, false, false)
    sheet.getRange("I27").setBorder(true, true, true, true, false, false)
    sheet.getRange("J27").setFormula("=H27*I27");

    sheet.getRange("H28").setBorder(true, true, true, true, false, false)
    sheet.getRange("I28").setBorder(true, true, true, true, false, false)
    sheet.getRange("J28").setFormula("=H28*I28");

    sheet.getRange("H29").setBorder(true, true, true, true, false, false)
    sheet.getRange("I29").setBorder(true, true, true, true, false, false)
    sheet.getRange("J29").setFormula("=H29*I29");

    sheet.getRange("H30").setBorder(true, true, true, true, false, false)
    sheet.getRange("I30").setBorder(true, true, true, true, false, false)
    sheet.getRange("J30").setFormula("=H30*I30");

    sheet.getRange("H31").setBorder(true, true, true, true, false, false)
    sheet.getRange("I31").setBorder(true, true, true, true, false, false)
    sheet.getRange("J31").setFormula("=H31*I31");
    sheet.getRange("J21:J34").setBorder(true, true, true, true, false, false)
    sheet.getRange("H32:H34").setBorder(true, true, true, true, false, false)
    sheet.getRange("I32:I34").setBorder(true, true, true, true, false, false)

    sheet.getRange("G32").setValue("Subtotal").setFontWeight("bold");
    sheet.getRange("J32").setFormula("=SUM(J25:J31)").setFontWeight("bold");

    sheet.getRange("A33").setValue("EXCLUSIONS:");
    sheet.getRange("B34").setValue("none");
    sheet.getRange("B34:F35").merge()
		
    sheet.getRange("h35").setValue("GRAND TOTAL:").setFontWeight("bold");
    sheet.getRange("J35").setFormula("=SUM(J32)").setFontWeight("bold");

    sheet.getRange("A36:J36").merge().setBorder(true, true, true, true, false, false).setValue("Purchase Order Number must appear on all invoices, shipments, and correspondence");
    sheet.getRange("F37").setValue("Attachment Link:");
    sheet.getRange("A37:E37").merge().setValue("See attached sheet for additional terms and conditions of this offer")
    sheet.getRange("G37:J37").merge()

    sheet.getRange("A38:E38").merge().setValue("Payment Terms: Reference contract between the owner and");
    sheet.getRange("F38:J38").merge().setValue("Your Company Name");

    sheet.getRange("A39:J39").merge().setBorder(true, true, true, true, false, false)

    sheet.getRange("A40:J40").merge().setValue("Vendor is to supply a minimum of 24 hours advance notice of shipment of material");

    sheet.getRange("A42:B42").merge().setValue("Acknowledged By:");
    sheet.getRange("F42").setValue("Originated By:");

    sheet.getRange("C42:E42").merge()
    sheet.getRange("G42:J42").merge()
  
    sheet.getRange("A43").setValue("Vendor:");
    sheet.getRange("B43").setValue("Company");
    sheet.getRange("B43:E43").merge()

    sheet.getRange("A45").setValue("Date Signed:");
    sheet.getRange("F45").setValue("Date of Order");

    sheet.getRange("A46:J46").merge().setBorder(true, true, true, true, false, false).setValue("IMPORTANT: THIS OFFER DOES NOT BECOME AN ORDER UNTIL ALL COPIES ARE SIGNED AND BOTH COPIES ARE RETURNED TO THIS OFFICE.");

    // Insert two rows at the top
  sheet.insertRowsBefore(1, 2);


  // Email Notification
const recipient = "projectprodigyapp@gmail.com";
const subject = "Purchase Order Template Created!";
const body = `A new Purchase Order Template has been created successfully in Google Sheets.\n\n
Another user from Opensource.`;

MailApp.sendEmail(recipient, subject, body);


    Logger.log("Purchase Order Template Created Successfully");
}

function setup() {
  newfileit();
  contactsit();
  createInventoryTemplate();
  createInvoiceTemplate();
  createReceiptTemplate();
  createPackingSlipTemplate();
  
}

function createInvoiceTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Define ss
  var sheet = ss.getSheetByName('Sheet1') || ss.insertSheet('Sheet1');
  sheet.clear();

  // Header Section
  sheet.getRange('A1:E1').merge().setValue('Your Company Name').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A3:E3').merge().setValue('Business Street').setFontSize(10);
  sheet.getRange('A4:E4').merge().setValue('Business City, Business State Business Postal Code').setFontSize(10);
  sheet.getRange('A5:E5').merge().setValue('E-mail Address').setFontSize(10);
  sheet.getRange('A6:E6').merge().setValue('Business Phone').setFontSize(10);
  sheet.getRange('A8:E8').merge().setValue('INVOICE').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A9').setValue('Number').setFontSize(10).setFontWeight('bold');
  sheet.getRange('A10').setValue('Bill to:').setFontSize(10).setFontWeight('bold');

  // Client Information
  var contactsSheet = ss.getSheetByName("contacts"); 
  if (contactsSheet) { // Ensure the sheet exists before using it
    sheet.getRange("A11").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(contactsSheet.getRange("A:A"), true) // Removed "$" for Apps Script compatibility
        .build()
    );
  } else {
    Logger.log("Error: 'contacts' sheet not found.");
  } 
  

  sheet.getRange('A12:B12').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 44, FALSE)");
  sheet.getRange('A13:B13').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 26, FALSE)");
  sheet.getRange('A14:B14').merge().setFormula(
    '=VLOOKUP(A11, contacts!A:CJ, 30, FALSE) & ", " & ' +
    'VLOOKUP(A11, contacts!A:CJ, 31, FALSE) & "   " & ' +
    'VLOOKUP(A11, contacts!A:CJ, 32, FALSE)'
  );
  sheet.getRange('A15:B15').merge().setFormula("=HYPERLINK(VLOOKUP(A11, contacts!A:CJ, 16, FALSE))");
  sheet.getRange('A16:B16').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 20, FALSE)");

  sheet.getRange('D10:E10').merge().setValue('Date:').setFontWeight('bold');
  sheet.getRange('D11:E11').merge().setFormula("=TODAY()");

  // Invoice Details Table Header
  sheet.getRange('A18').setValue('Description').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('B18').setValue('Quantity').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('C18').setValue('Unit Price').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('D18').setValue('Total').setFontWeight('bold').setBackground('#cccccc');

  var inventorySheet = ss.getSheetByName("Inventory"); 
  if (inventorySheet) { // Ensure the sheet exists before using it
    sheet.getRange("A19:A28").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(inventorySheet.getRange("A:A"), true) // Removed "$" for Apps Script compatibility
        .build()
    );
  } else {
    
    Logger.log("Error: 'Inventory' sheet not found.");
  }

  // Populate formulas for items in the invoice
  for (var row = 19; row <= 28; row++) {
    sheet.getRange('C' + row).setFormula(`=IFERROR(VLOOKUP(A${row},Inventory!$A$2:$CL$9341,3,FALSE), 0)`).setFontSize(10);
    sheet.getRange('D' + row).setFormula(`=IFERROR($B${row}*$C${row},0)`).setFontSize(10);
  }

  // Summary Section
  sheet.getRange('C30').setValue('Subtotal:').setFontWeight('bold');
  sheet.getRange('D30').setFormula('=SUM(D19:D28)');
  
  sheet.getRange('B31').setValue('Tax:').setFontWeight('bold');
  sheet.getRange('C31').setValue('.07');
  sheet.getRange('D31').setFormula('=D30*C31');
  
  sheet.getRange('C32').setValue('Total:').setFontWeight('bold');
  sheet.getRange('D32').setFormula('=D30+D31');
  
  // Payment Instructions
  sheet.getRange('A34:E34').merge().setValue('Payment Instructions:').setFontWeight('bold');
  sheet.getRange('A35:E35').merge().setValue('[Your Payment Instructions]');

  // Formatting the sheet
  sheet.getRange('A1:E32').setHorizontalAlignment('center');
  sheet.getRange('A1:E6').setHorizontalAlignment('left');
  sheet.getRange('A9:B17').setHorizontalAlignment('left');
  sheet.getRange('A34:E35').setHorizontalAlignment('left');
  
  sheet.setColumnWidth(1, 350); // Set column A to 350
  sheet.setColumnWidths(2, 3, 100); // Set columns B, C, D to 100
  
  // Setting borders for the table
  sheet.getRange('A18:D28').setBorder(true, true, true, true, true, true);
  
  // Setting number formats
  sheet.getRange('C19:C28').setNumberFormat('$#,##0.00');
  sheet.getRange('D19:D28').setNumberFormat('$#,##0.00'); // Fixed range typo
  sheet.getRange('D30:D32').setNumberFormat('$#,##0.00');

  // Insert two rows at the top
  sheet.insertRowsBefore(1, 2);
}


function createReceiptTemplate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Receipt') || spreadsheet.insertSheet('Receipt');
  sheet.clear();

    // Header Section
  sheet.getRange('A1:E1').merge().setFormula("=View_Print!A3").setFontSize(16).setFontWeight('bold');
  sheet.getRange('A3:E3').merge().setFormula("=View_Print!A4").setFontSize(10);
  sheet.getRange('A4:E4').merge().setFormula("=View_Print!A5").setFontSize(10);
  sheet.getRange('A5:E5').merge().setFormula("=View_Print!A6").setFontSize(10);
  sheet.getRange('A6:E6').merge().setFormula("=View_Print!A7").setFontSize(10);
  sheet.getRange('A8:E8').merge().setValue('RECEIPT').setFontSize(14).setFontWeight('bold');

  
  // Client Information
  sheet.getRange('A10:B10').merge().setValue('Bill To:').setFontWeight('bold');
  sheet.getRange('A11:B11').merge().setFormula("=View_Print!A13");
  sheet.getRange('A12:B12').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 44, FALSE)");
  sheet.getRange('A13:B13').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 26, FALSE)");
  sheet.getRange('A14:B14').merge().setFormula(
  '=VLOOKUP(A11, contacts!A:CJ, 30, FALSE) & ", " & ' +
  'VLOOKUP(A11, contacts!A:CJ, 31, FALSE) & "   " & ' +
  'VLOOKUP(A11, contacts!A:CJ, 32, FALSE)'
);
  sheet.getRange('A15:B15').merge().setFormula("=HYPERLINK(VLOOKUP(A11, contacts!A:CJ, 16, FALSE))");
  sheet.getRange('A16:B16').merge().setFormula("=VLOOKUP(A13, contacts!A:CJ, 20, FALSE)");

   
  sheet.getRange('D10:E10').merge().setValue('Date Shipped:').setFontWeight('bold');
  sheet.getRange('D11:E11').merge().setFormula("=View_Print!O2");
  
  // Invoice Details Table Header
  sheet.getRange('A18').setValue('Description').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('B18').setValue('Quantity').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('C18').setValue('Unit Price').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('D18').setValue('Total').setFontWeight('bold').setBackground('#cccccc');
  
  // Invoice Details Table (sample rows)
  sheet.getRange('A19').setFormula("=View_Print!A21").setFontSize(10);
  sheet.getRange('A20').setFormula("=View_Print!A22").setFontSize(10);
  sheet.getRange('A21').setFormula("=View_Print!A23").setFontSize(10);
  sheet.getRange('A22').setFormula("=View_Print!A24").setFontSize(10);
  sheet.getRange('A23').setFormula("=View_Print!A25").setFontSize(10);
  sheet.getRange('A24').setFormula("=View_Print!A26").setFontSize(10);
  sheet.getRange('A25').setFormula("=View_Print!A27").setFontSize(10);
  sheet.getRange('A26').setFormula("=View_Print!A28").setFontSize(10);
  sheet.getRange('A27').setFormula("=View_Print!A29").setFontSize(10);
  sheet.getRange('A28').setFormula("=View_Print!A30").setFontSize(10);

  sheet.getRange('B19').setFormula("=View_Print!B21").setFontSize(10);
  sheet.getRange('B20').setFormula("=View_Print!B22").setFontSize(10);
  sheet.getRange('B21').setFormula("=View_Print!B23").setFontSize(10);
  sheet.getRange('B22').setFormula("=View_Print!B24").setFontSize(10);
  sheet.getRange('B23').setFormula("=View_Print!B25").setFontSize(10);
  sheet.getRange('B24').setFormula("=View_Print!B26").setFontSize(10);
  sheet.getRange('B25').setFormula("=View_Print!B27").setFontSize(10);
  sheet.getRange('B26').setFormula("=View_Print!B28").setFontSize(10);
  sheet.getRange('B27').setFormula("=View_Print!B29").setFontSize(10);
  sheet.getRange('B28').setFormula("=View_Print!B30").setFontSize(10);

  sheet.getRange('C19').setFormula("=View_Print!C21").setFontSize(10);
  sheet.getRange('C20').setFormula("=View_Print!C22").setFontSize(10);
  sheet.getRange('C21').setFormula("=View_Print!C23").setFontSize(10);
  sheet.getRange('C22').setFormula("=View_Print!C24").setFontSize(10);
  sheet.getRange('C23').setFormula("=View_Print!C25").setFontSize(10);
  sheet.getRange('C24').setFormula("=View_Print!C26").setFontSize(10);
  sheet.getRange('C25').setFormula("=View_Print!C27").setFontSize(10);
  sheet.getRange('C26').setFormula("=View_Print!C28").setFontSize(10);
  sheet.getRange('C27').setFormula("=View_Print!C29").setFontSize(10);
  sheet.getRange('C28').setFormula("=View_Print!C30").setFontSize(10);

  sheet.getRange('D19').setFormula("=View_Print!D21").setFontSize(10);
  sheet.getRange('D20').setFormula("=View_Print!D22").setFontSize(10);
  sheet.getRange('D21').setFormula("=View_Print!D23").setFontSize(10);
  sheet.getRange('D22').setFormula("=View_Print!D24").setFontSize(10);
  sheet.getRange('D23').setFormula("=View_Print!D25").setFontSize(10);
  sheet.getRange('D24').setFormula("=View_Print!D26").setFontSize(10);
  sheet.getRange('D25').setFormula("=View_Print!D27").setFontSize(10);
  sheet.getRange('D26').setFormula("=View_Print!D28").setFontSize(10);
  sheet.getRange('D27').setFormula("=View_Print!D29").setFontSize(10);
  sheet.getRange('D28').setFormula("=View_Print!D30").setFontSize(10);
  
  
  // Summary Section
  sheet.getRange('C30').setValue('Subtotal:').setFontWeight('bold');
  sheet.getRange('D30').setValue('=View_Print!D32');
  
  sheet.getRange('B31').setValue('Tax:').setFontWeight('bold');
  sheet.getRange('C31').setValue('=View_Print!C33');
  sheet.getRange('D31').setValue('=View_Print!D33');
  
  sheet.getRange('C32').setValue('Total:').setFontWeight('bold');
  sheet.getRange('D32').setValue('=View_Print!D34');

  // Receipt Note
  sheet.getRange('A29:E29').merge().setValue('Thank you for your business!').setFontWeight('bold');
  
   // Formatting the sheet
  sheet.getRange('A1:E32').setHorizontalAlignment('center');
  sheet.getRange('A1:E6').setHorizontalAlignment('left');
  sheet.getRange('A9:B17').setHorizontalAlignment('left');
  sheet.getRange('A34:E35').setHorizontalAlignment('left');
  sheet.setColumnWidth(1, 350); // Set column A to 350
  sheet.setColumnWidths(2, 3, 100); // Set columns B, C, D to 100
  
  // Setting borders for the table
  sheet.getRange('A18:D28').setBorder(true, true, true, true, true, true);
  
  // Setting number formats
  sheet.getRange('C19:C28').setNumberFormat('$#,##0.00');
  sheet.getRange('D193:D28').setNumberFormat('$#,##0.00');
  sheet.getRange('D30:D32').setNumberFormat('$#,##0.00');
}

function createInventoryTemplate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Inventory');
  
  // If the sheet already exists, delete it
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
  
  // Create a new Inventory sheet
  sheet = spreadsheet.insertSheet('Inventory');

  // Set up the header row
  sheet.getRange('A1').setValue('Item Description').setFontWeight('bold');
  sheet.getRange('B1').setValue('Quantity in Stock').setFontWeight('bold');
  sheet.getRange('C1').setValue('Unit Price').setFontWeight('bold');
  sheet.getRange('D1').setValue('Category').setFontWeight('bold');
  sheet.getRange('E1').setValue('Supplier').setFontWeight('bold');
  sheet.getRange('F1').setValue('Image').setFontWeight('bold');

  // Set column widths for better readability
  sheet.setColumnWidths(1, 5, 150);

  // Sample data
  var sampleData = [
    ['Item 1', 100, 10.00, 'Category 1', 'Supplier A', 'https://drive.google.com/uc?export=view&id=165kqv1atBk1WBbSkIbj6pnoikR9JOpLj'],
    ['Item 2', 200, 15.00, 'Category 2', 'Supplier B', 'https://drive.google.com/uc?export=view&id=165kqv1atBk1WBbSkIbj6pnoikR9JOpLj'],
    ['Item 3', 150, 20.00, 'Category 3', 'Supplier C', 'https://drive.google.com/uc?export=view&id=165kqv1atBk1WBbSkIbj6pnoikR9JOpLj']
  ];

  // Populate the sheet with sample data
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
}



function createPackingSlipTemplate() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Packing Slip') || spreadsheet.insertSheet('Packing Slip');
    sheet.clear();

    // Header Section
    sheet.getRange('A1:E1').merge().setFormula("=View_Print!A3").setFontSize(16).setFontWeight('bold');
    sheet.getRange('A3:E3').merge().setFormula("=View_Print!A4").setFontSize(10);
    sheet.getRange('A4:E4').merge().setFormula("=View_Print!A5").setFontSize(10);
    sheet.getRange('A5:E5').merge().setFormula("=View_Print!A6").setFontSize(10);
    sheet.getRange('A6:E6').merge().setFormula("=View_Print!A7").setFontSize(10);

    sheet.getRange('A8:E8').merge().setValue('PACKING SLIP').setFontSize(14).setFontWeight('bold');

    // Client Information
    sheet.getRange('A10:B10').merge().setValue('Ship To:').setFontWeight('bold');
    sheet.getRange('A11:B11').merge().setFormula("=View_Print!A13");
    sheet.getRange('A12:B12').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 44, FALSE)");
    sheet.getRange('A13:B13').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 63, FALSE)");
    sheet.getRange('A14:B14').merge().setFormula(
      '=VLOOKUP(A11, contacts!A:CJ, 67, FALSE) & ", " & ' +
      'VLOOKUP(A11, contacts!A:CJ, 68, FALSE) & "   " & ' +
      'VLOOKUP(A11, contacts!A:CJ, 69, FALSE)'
    );
    sheet.getRange('A15:B15').merge().setFormula("=HYPERLINK(VLOOKUP(A11, contacts!A:CJ, 16, FALSE))");
    sheet.getRange('A16:B16').merge().setFormula("=VLOOKUP(A13, contacts!A:CJ, 20, FALSE)");

    sheet.getRange('D10:E10').merge().setValue('Date Shipped:').setFontWeight('bold');
    sheet.getRange('D11:E11').merge().setFormula("=View_Print!O2");

    // Invoice Details Table Header
    sheet.getRange('A18').setValue('Description').setFontWeight('bold').setBackground('#cccccc');
    sheet.getRange('B18').setValue('Quantity').setFontWeight('bold').setBackground('#cccccc');
    sheet.getRange('C18').setValue('Unit Price').setFontWeight('bold').setBackground('#cccccc');
    sheet.getRange('D18').setValue('Total').setFontWeight('bold').setBackground('#cccccc');

    // Invoice Details Table (sample rows)
    for (var i = 19; i <= 28; i++) {
      sheet.getRange('A' + i).setFormula(`=View_Print!A${i + 2}`).setFontSize(10);
      sheet.getRange('B' + i).setFormula(`=View_Print!B${i + 2}`).setFontSize(10);
      sheet.getRange('C' + i).setFormula(`=View_Print!C${i + 2}`).setFontSize(10);
      sheet.getRange('D' + i).setFormula(`=View_Print!D${i + 2}`).setFontSize(10);
    }

    // Summary Section
    sheet.getRange('C30').setValue('Subtotal:').setFontWeight('bold');
    sheet.getRange('D30').setFormula('=View_Print!D32');

    sheet.getRange('B31').setValue('Tax:').setFontWeight('bold');
    sheet.getRange('C31').setFormula('=View_Print!C33');
    sheet.getRange('D31').setFormula('=View_Print!D33');

    sheet.getRange('C32').setValue('Total:').setFontWeight('bold');
    sheet.getRange('D32').setFormula('=View_Print!D34');

    // Note Section
    sheet.getRange('A29:E29').merge().setValue('Thank you for your business!').setFontWeight('bold');

    // Formatting the sheet
    sheet.getRange('A1:E32').setHorizontalAlignment('center');
    sheet.getRange('A1:E6').setHorizontalAlignment('left');
    sheet.getRange('A9:B17').setHorizontalAlignment('left');
    sheet.getRange('A34:E35').setHorizontalAlignment('left');
    sheet.setColumnWidth(1, 350); // Set column A to 350
    sheet.setColumnWidths(2, 3, 100); // Set columns B, C, D to 100

    // Setting borders for the table
    sheet.getRange('A18:D28').setBorder(true, true, true, true, true, true);

    // Setting number formats
    sheet.getRange('C19:C28').setNumberFormat('$#,##0.00');
    sheet.getRange('D19:D28').setNumberFormat('$#,##0.00');
    sheet.getRange('D30:D32').setNumberFormat('$#,##0.00');
    

try {
  var sheetsToMove = ['Packing Slip', 'Receipt', 'Inventory']; // Reverse order
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Loop through the sheets and move them to the front
  sheetsToMove.forEach(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      spreadsheet.setActiveSheet(sheet);
      spreadsheet.moveActiveSheet(0); // Move to the front (index 0)
    }
  });

  // Activate 'Sheet1' at the end
  var sheet1 = spreadsheet.getSheetByName('Sheet1');
  if (sheet1) {
    sheet1.activate();
  } else {
    SpreadsheetApp.getUi().alert("Sheet1 not found.");
  }

  // Email Notification
const recipient = "projectprodigyapp@gmail.com";
const subject = "Inventory Template Created!";
const body = `A new Inventory Template has been created successfully in Google Sheets.\n\n
Another user from Opensource.`;

MailApp.sendEmail(recipient, subject, body);

  SpreadsheetApp.getUi().alert("Inventory Template created successfully. Please support DataMateApps and help us grow!");

} catch (e) {
  SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
}


  copyInput1it()
  viewit()

  SpreadsheetApp.getUi().alert("Inventory Template created successfully. Please support DataMateApps and help us grow!");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}


function newfileit() {
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
  inputSheet.getCurrentCell().setValue("=A11");
  inputSheet.getRange("A2").activate();
  inputSheet.getCurrentCell().setValue("=B11");
  inputSheet.getRange("B1").activate();
  inputSheet.getCurrentCell().setValue("=D12");
  inputSheet.getRange("B2").activate();
  inputSheet.getCurrentCell().setValue("=D13");
  inputSheet.getRange("C1").activate();
  inputSheet.getCurrentCell().setValue("=A12");
  inputSheet.getRange("C2").activate();
  inputSheet.getCurrentCell().setValue("=A13");
  inputSheet.getRange("D1").activate();
  inputSheet.getCurrentCell().setValue("=C34");
  inputSheet.getRange("D2").activate();
  inputSheet.getCurrentCell().setValue("=D34");
  inputSheet.getRange("E1").activate();
  inputSheet.getCurrentCell().setValue("=C32");
  inputSheet.getRange("E2").activate();
  inputSheet.getCurrentCell().setValue("=D32");
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
  inputSheet.getCurrentCell().setValue("Date Paid");
  inputSheet.getRange("N1").activate();
  inputSheet.getCurrentCell().setValue("Amount");
  inputSheet.getRange("O1").activate();
  inputSheet.getCurrentCell().setValue("Date Shipped");
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
  logSheet.getCurrentCell().setValue("Orders Log");
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

}

function copyInput1it() {
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
}

function copyInput2it() {
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
 }

function viewit() {
  
  copyInput2it(); 
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


}

function contactsit() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["contacts", "Address", "NewContact"];
    const sheets = ss.getSheets();
    const wsSheet1 = ss.getSheetByName("Sheet1");

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
  } catch (error) {
    Logger.log("Error in contactsit: " + error.message);
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
  }
}

function updateInventory() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var inventorySheet = spreadsheet.getSheetByName('Inventory');
  var invoiceSheet = spreadsheet.getSheetByName('Input');

  if (!inventorySheet || !invoiceSheet) {
    Logger.log('Missing required sheets: Inventory or Input');
    return;
  }

  // Get invoice data
  var invoiceData = invoiceSheet.getRange('A21:D30').getValues();

  // Loop through invoice data to process each item
  for (var i = 0; i < invoiceData.length; i++) {
    var itemDescription = invoiceData[i][0];
    var quantitySold = invoiceData[i][1];

    if (itemDescription && quantitySold) {
      // Get inventory data
      var inventoryData = inventorySheet.getRange('A2:B' + inventorySheet.getLastRow()).getValues();

      for (var j = 0; j < inventoryData.length; j++) {
        if (inventoryData[j][0] == itemDescription) {
          var currentStock = inventoryData[j][1];

          if (typeof currentStock === 'number' && currentStock >= quantitySold) {
            // Update inventory stock
            inventorySheet.getRange('B' + (j + 2)).setValue(currentStock - quantitySold);
          } else {
            Logger.log('Insufficient stock for item: ' + itemDescription);
          }
          break; // Exit inner loop once match is found
        }
      }
    }
  }
}
