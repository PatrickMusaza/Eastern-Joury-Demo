/**
 * Google Apps Script for Client Bill Management
 * Handles CRUD operations and integrates with Google Sheets
 */

// CONSTANTS
const SPREADSHEET_ID = "1hoiskygCvco34k1pLNllvBXR9MBdf9aEjs07YGHnJ2I"; // Replace with your Google Sheets ID
const CLIENT_BILL_SHEET = "ClientBill"; // Name of the sheet for Client Bill data
const CLIENT_BILL_RANGE = "ClientBill!A2:S"; // Range for Client Bill data (adjust columns as needed)
const ABBREVIATIONS_SHEET = "Abbreviations"; // Name of the sheet for Abbreviations
const ABBREVIATIONS_RANGE = "Abbreviations!A2:B"; // Range for Abbreviations data
const ITEM_DATA_SHEET = "Item";
const EXPENSES_SHEET='Expenses';

// Display HTML page
function doGet(request) {
  let html = HtmlService.createTemplateFromFile("Index").evaluate();
  let htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
  return htmlOutput;
}

// PROCESS CLIENT BILL FORM SUBMISSION
function processClientBill(formObject) {
  if (formObject.recId && checkClientBillId(formObject.recId)) {
    // Update existing record
    const values = [
      [
        formObject.recId,
        formObject.billNo,
        formObject.shipperName,
        formObject.telephoneNo,
        formObject.receiverName1,
        formObject.phoneNo1,
        formObject.receiverName2,
        formObject.phoneNo2,
        formObject.containerNo,
        formObject.totalPieces,
        formObject.actualWeight,
        formObject.discountWeight,
        formObject.chargeableWeight,
        formObject.ratePerKg,
        formObject.billCharge,
        formObject.discountCharge,
        formObject.totalCharges,
        formObject.paidAmount,
        formObject.outstandingBalance,
        new Date().toLocaleString(),
      ],
    ];
    const updateRange = getClientBillRangeById(formObject.recId);
    updateClientBillRecord(values, updateRange);
  } else {
    // Create new record
    const values = [
      [
        generateUniqueId(),
        formObject.billNo,
        formObject.shipperName,
        formObject.telephoneNo,
        formObject.receiverName1,
        formObject.phoneNo1,
        formObject.receiverName2,
        formObject.phoneNo2,
        formObject.containerNo,
        formObject.totalPieces,
        formObject.actualWeight,
        formObject.discountWeight,
        formObject.chargeableWeight,
        formObject.ratePerKg,
        formObject.billCharge,
        formObject.discountCharge,
        formObject.totalCharges,
        formObject.paidAmount,
        formObject.outstandingBalance,
        new Date().toLocaleString(),
        new Date().toLocaleString(),
      ],
    ];
    createClientBillRecord(values);
  }
  return getClientBillData();
}

// GET ITEMS FOR A SPECIFIC BILL NO
function getItemsForBill(billNo) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Item");
  const data = sheet.getRange("A2:D").getValues(); // Fetch columns A to D
  return data.filter((row) => row[3] === billNo); // Filter rows with matching Bill No
}

// GET NEXT BILL NO
function getNextBillNo() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const lastRow = sheet.getLastRow();
  const lastBillNo = sheet.getRange(lastRow, 2).getValue(); // Assuming Bill No is in column B
  const nextBillNo = incrementBillNo(lastBillNo); // Increment the last Bill No
  return nextBillNo;
}

// HELPER FUNCTION TO INCREMENT BILL NO
function incrementBillNo(billNo) {
  const prefix = billNo.split("-")[0]; // Extract prefix (e.g., "BILL")
  const number = parseInt(billNo.split("-")[1]); // Extract number
  return `${prefix}-${number + 1}`; // Increment and return
}

// GET LAST BILL NO
function getLastBillNo() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null; // No data
  const lastBillNo = sheet.getRange(lastRow, 2).getValue(); // Fetch from Column B
  return lastBillNo;
}

// CREATE CLIENT BILL RECORD
function createClientBillRecord(values) {
  try {
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
    sheet.appendRow(values[0]);
  } catch (err) {
    console.log("Failed with error: " + err.message);
  }
}

// SAVE ITEM TO ITEM SHEET
function saveItem(item) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ITEM_DATA_SHEET);
  sheet.appendRow([item.number, item.type, item.weight, item.bill]);
}

// CLEAR ITEMS FOR CURRENT BILL NO
function clearItemsForBill(billNo) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Item");
  const data = sheet.getRange("A2:D").getValues(); // Assuming columns: Number, Type, Weight, Bill
  const rowsToDelete = data
    .map((row, index) => (row[3] === billNo ? index + 2 : null))
    .filter((index) => index !== null); // Find rows with matching Bill No
  rowsToDelete.reverse().forEach((rowIndex) => sheet.deleteRow(rowIndex)); // Delete rows (reverse to avoid index issues)
}

// READ CLIENT BILL RECORDS
function readClientBillRecords(range) {
  try {
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
    return sheet.getRange(range).getValues();
  } catch (err) {
    console.log("Failed with error: " + err.message);
  }
}

// UPDATE CLIENT BILL RECORD
function updateClientBillRecord(values, range) {
  try {
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
    sheet.getRange(range).setValues(values);
  } catch (err) {
    console.log("Failed with error: " + err.message);
  }
}

// DELETE CLIENT BILL RECORD
function deleteClientBillRecord(id) {
  try {
    const rowIndex = getClientBillRowIndexById(id);
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
    sheet.deleteRow(rowIndex);
  } catch (err) {
    console.log("Failed with error: " + err.message);
  }
  return getClientBillData();
}

// GET ALL CLIENT BILL RECORDS
function getAllClientBillRecords() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  return data.filter((row) => row.some((cell) => cell !== "")); // Filter out completely empty rows
}

// GET ALL CLIENT BILL DATA
function getClientBillData() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues();
  return data.filter((row) => row.some((cell) => cell !== "")); // Filter out completely empty rows
}

// UPDATE CLIENT BILL RECORD
function updateClientBill(formObject) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  const rowIndex = data.findIndex((row) => row[0] === formObject.recId); // Find the row with the matching recId
  if (rowIndex !== -1) {
    const values = [
      formObject.recId,
      formObject.billNo,
      formObject.shipperName,
      formObject.telephoneNo,
      formObject.receiverName1,
      formObject.phoneNo1,
      formObject.receiverName2,
      formObject.phoneNo2,
      formObject.containerNo,
      formObject.totalPieces,
      formObject.actualWeight,
      formObject.discountWeight,
      formObject.chargeableWeight,
      formObject.ratePerKg,
      formObject.billCharge,
      formObject.discountCharge,
      formObject.totalCharges,
      formObject.paidAmount,
      formObject.outstandingBalance,
      new Date().toLocaleString(),
    ];
    sheet.getRange(rowIndex + 2, 1, 1, values.length).setValues([values]); // Update the row
  }
}

// GET CLIENT BILL RECORD BY ID
function getClientBillRecordById(id) {
  const range = getClientBillRangeById(id);
  if (!range) return null;
  return readClientBillRecords(range);
}

// GET ROW INDEX BY CLIENT BILL ID
function getClientBillRowIndexById(id) {
  const idList = readClientBillRecords("ClientBill!A2:A");
  for (let i = 0; i < idList.length; i++) {
    if (id === idList[i][0]) {
      return i + 2; // +2 to account for header row and 0-based index
    }
  }
  return -1;
}

// VALIDATE CLIENT BILL ID
function checkClientBillId(id) {
  const idList = readClientBillRecords("ClientBill!A2:A").flat();
  return idList.includes(id);
}

// GET RANGE IN A1 NOTATION FOR CLIENT BILL ID
function getClientBillRangeById(id) {
  const rowIndex = getClientBillRowIndexById(id);
  if (rowIndex === -1) return null;
  return `ClientBill!A${rowIndex}:O${rowIndex}`; // Adjust columns as needed
}

// GENERATE UNIQUE ID
function generateUniqueId() {
  return Utilities.getUuid();
}

// SEARCH CLIENT BILL RECORDS
function searchClientBill(searchText) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  return data.filter(
    (row) =>
      row.some((cell) => cell.toString().toLowerCase().includes(searchText)) // Filter rows that match the search text
  );
}

// GET RECORD BY ID
function getRecordByIds(recId) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  const record = data.find((row) => row[0] === recId); // Find the record with the matching recId
  return record ? [record] : null; // Return the record as an array (to match the expected format)
}

// DELETE ITEM FROM ITEM SHEET
function deleteItem(billNo, itemNumber) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Item");
  const data = sheet.getRange("A2:D").getValues(); // Fetch columns A to D
  const rowIndex = data.findIndex(
    (row) => row[3] === billNo && row[0] == itemNumber
  ); // Find the row with matching Bill No and item number
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 2); // +2 to account for header row and 0-based index
  }
}

// GET FIRST 20 CLIENT BILL RECORDS
function getFirstTwentyRecords() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  const filteredData = data.filter((row) => row.some((cell) => cell !== "")); // Filter out empty rows
  return filteredData.slice(0, 20); // Return only the first 20 records
}

// GET RECORD BY ID (WITH ITEMS)
function getRecordById(recId) {
  const clientBillSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const itemSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Item');

  // Fetch client bill details
  const clientBillData = clientBillSheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  const clientBill = clientBillData.find(row => row[0] === recId); // Find the record with the matching recId

  // Fetch associated items
  const itemData = itemSheet.getRange("A2:D").getValues(); // Fetch columns A to D
  const items = itemData.filter(row => row[3] === clientBill[1]); // Filter items with matching Bill No

  return {
    clientBill: clientBill,
    items: items
  };
}

//ABBREVIATIONS

// SAVE ABBREVIATION
function saveAbbreviation(formObject) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);
  const data = sheet.getRange(ABBREVIATIONS_RANGE).getValues(); // Fetch columns A and B
  const rowIndex = data.findIndex(row => row[0] === formObject.name); // Find the row with the matching name

  if (rowIndex !== -1) {
    // Update existing abbreviation
    sheet.getRange(rowIndex + 2, 1, 1, 2).setValues([[formObject.name, formObject.value]]);
  } else {
    // Create new abbreviation
    sheet.appendRow([formObject.name, formObject.value]);
  }
}

// GET ABBREVIATIONS
function getAbbreviations() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);
  const data = sheet.getRange(ABBREVIATIONS_RANGE).getValues(); // Fetch columns A and B
  return data.filter((row) => row[0] && row[1]); // Filter out empty rows
}

// GET ABBREVIATION BY NAME
function getAbbreviationByName(name) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);
  const data = sheet.getRange(ABBREVIATIONS_RANGE).getValues(); // Fetch columns A and B
  const abbreviation = data.find((row) => row[0] === name); // Find the abbreviation with the matching name
  return abbreviation ? [abbreviation] : null; // Return the abbreviation as an array
}

// DELETE ABBREVIATION
function deleteAbbreviation(name) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);
  const data = sheet.getRange(ABBREVIATIONS_RANGE).getValues(); // Fetch columns A and B
  const rowIndex = data.findIndex((row) => row[0] === name); // Find the row with the matching name
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 2); // +2 to account for header row and 0-based index
  }
}

//HOUSE WAY BILL

// GENERATE HOUSE WAY BILL
function generateHouseWayBill(recId) {
  const clientBillSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = clientBillSheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  const record = data.find(row => row[0] === recId); // Find the record with the matching recId

  if (!record) {
    throw new Error("Record not found");
  }

  // Template ID (replace with your Google Docs template ID)
  const templateId = '1YflqixKBH--hayUdudZn4PYagFwRxWrzqhuWMkoJDZ8';

  // Create a copy of the template
  const templateDoc = DriveApp.getFileById(templateId);
  const newDoc = templateDoc.makeCopy(`House_Way_Bill_${record[1]}`, DriveApp.getRootFolder());
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // Replace placeholders with actual values
  body.replaceText('{{BillNo}}', record[1]);
  body.replaceText('{{ShipperName}}', record[2]);
  body.replaceText('{{ShipperTel}}', record[3]);
  body.replaceText('{{ReceiverName1}}', record[4]);
  body.replaceText('{{PhoneNo1}}', record[5]);
  body.replaceText('{{ReceiverName2}}', record[6]);
  body.replaceText('{{PhoneNo2}}', record[7]);
  body.replaceText('{{ContainerNo}}', record[8]);

  // Save and close the document
  doc.saveAndClose();

  // Export as PDF
  const pdfBlob = newDoc.getAs(MimeType.PDF);
  const pdfFile = DriveApp.createFile(pdfBlob);

  // Get the PDF download URL
  const pdfUrl = pdfFile.getUrl();

  // Delete the temporary document
  DriveApp.getFileById(newDocId).setTrashed(true);

  return pdfUrl;
}


//CONTAINER NO

function getAllContainerRecords() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues();
  return data.filter((row) => row.some((cell) => cell !== "")); // Filter out completely empty rows
}

function updateContainer(originalContainer, newContainer) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    if (data[i][8] === originalContainer) { // Assuming container number is at index 8
      sheet.getRange(i + 1, 9).setValue(newContainer); // Update container number
    }
  }
}

//EXPENSES

function getAllData() {
  const expensesSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET); // Replace with your sheet name
  const containersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET); // Replace with your sheet name

  const expensesData = expensesSheet.getDataRange().getValues();
  const containersData = containersSheet.getDataRange().getValues();

  return {
    expenses: expensesData.slice(1), // Skip header row
    containers: containersData.map(row => row[0]) // Assuming container numbers are in the first column
  };
}

function saveExpense(expenseData) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET); // Replace with your sheet name
  sheet.appendRow([
    Utilities.getUuid(), // Generate a unique ID for the expense
    expenseData.containerNo,
    expenseData.officeTransport,
    expenseData.containerTransportOffice,
    expenseData.containerLoadingLabour,
    // Add more expense fields here
    expenseData.date
  ]);
}


//RECEIPT


/*-------------------GENERAL-------------------------*/

// INCLUDE HTML PARTS (JS, CSS, OTHER HTML FILES)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
