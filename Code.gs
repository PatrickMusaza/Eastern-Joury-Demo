/**
 * Google Apps Script for Client Bill Management
 * Handles CRUD operations and integrates with Google Sheets
 */

// CONSTANTS
const SPREADSHEET_ID = "1hoiskygCvco34k1pLNllvBXR9MBdf9aEjs07YGHnJ2I"; // Replace with your Google Sheets ID
const CLIENT_BILL_SHEET = "ClientBill"; // Name of the sheet for Client Bill data
const CLIENT_BILL_RANGE = "ClientBill!A2:I"; // Range for Client Bill data (adjust columns as needed)
const ABBREVIATIONS_SHEET = "Abbreviations"; // Name of the sheet for Abbreviations
const ABBREVIATIONS_RANGE = "Abbreviations!A2:B"; // Range for Abbreviations data
const ITEM_DATA_SHEET = "Item";

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
        new Date().toLocaleString(),
      ],
    ];
    const updateRange = getClientBillRangeById(formObject.recId);
    updateClientBillRecord(values, updateRange);
    clearItemsForBill(formObject.billNo); // Clear items for the current Bill No
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
        new Date().toLocaleString(),
      ],
    ];
    createClientBillRecord(values);
    clearItemsForBill(formObject.billNo); // Clear items for the current Bill No
  }
  return getClientBillData();
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
  const data = sheet.getRange("A2:I").getValues(); // Fetch columns A to I
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
function searchClientBill(formObject) {
  let result = [];
  try {
    if (formObject.searchText) {
      const data = readClientBillRecords(CLIENT_BILL_RANGE);
      const searchText = formObject.searchText.toLowerCase();
      for (let i = 0; i < data.length; i++) {
        for (let j = 0; j < data[i].length; j++) {
          if (data[i][j].toString().toLowerCase().includes(searchText)) {
            result.push(data[i]);
            break;
          }
        }
      }
    }
  } catch (err) {
    console.log("Failed with error: " + err.message);
  }
  return result;
}

// GET RECORD BY ID
function getRecordById(recId) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  const record = data.find((row) => row[0] === recId); // Find the record with the matching recId
  return record ? [record] : null; // Return the record as an array (to match the expected format)
}

// GET ABBREVIATIONS FOR DROPDOWN
function getAbbreviations() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);
  const data = sheet.getRange(ABBREVIATIONS_RANGE).getValues();
  return data.filter((row) => row[0] && row[1]); // Filter out rows where Column A or B is empty
}

// DELETE ITEM FROM ITEM SHEET
function deleteItem(itemNumber) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ITEM_DATA_SHEET);
  const data = sheet.getRange("A2:D").getValues(); // Assuming columns: Number, Type, Weight, Bill
  const rowIndex = data.findIndex((row) => row[0] == itemNumber); // Find the row with the matching item number
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 2); // +2 to account for header row and 0-based index
  }
}

// INCLUDE HTML PARTS (JS, CSS, OTHER HTML FILES)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
