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
const ITEM_DATA_RANGE = "Item!A2:D";
const EXPENSES_SHEET = "Expenses";
const EXPENSES_SHEET_RANGE = "Expenses!A2:F";
const USER_SHEET = "Users";
const USER_SHEET_RANGE = "Users!A2:C";
const LOGS_SHEET = "Logs";
const REPORT_TEMPLATE_ID = "1X_P2IBPF93K-hnzaQjdoXweZuJUDGcKgrIHc5IHaxpY"; // Replace with your Google Docs template ID
const REPORT_FOLDER_ID = "1tF1YwKsgzR9kYI1iWhB-NywThtbnP4sN"; // Replace with your target folder ID
const LOGS_SHEET_UPDATE = "LogFile";

const COLUMN_NAMES = {
  recId: 0,
  billNo: 1,
  shipperName: 2,
  telephoneNo: 3,
  receiverName1: 4,
  phoneNo1: 5,
  receiverName2: 6,
  phoneNo2: 7,
  containerNo: 8,
  totalPieces: 9,
  actualWeight: 10,
  discountWeight: 11,
  chargeableWeight: 12,
  ratePerKg: 13,
  billCharge: 14,
  discountCharge: 15,
  totalCharges: 16,
  paidAmount: 17,
  outstandingBalance: 18,
  dateUpdated: 19,
  dateCreated: 20,
};

// Display HTML page
function doGet(request) {
  let html = HtmlService.createTemplateFromFile("Index").evaluate();
  let htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
  return htmlOutput;
}

//LOGIN
function getUserCredentials() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USER_SHEET);
  const data = sheet.getRange(USER_SHEET_RANGE).getValues();
  const users = data.map((row) => ({
    username: row[0],
    password: row[1],
    role: row[2],
  }));
  return users;
}

//LOGS LOGIN
function logEvent(username, event, details) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGS_SHEET);
  const timestamp = new Date();
  sheet.appendRow([timestamp, username, event, details]);
}

//LOGS UPDATE
function logEventUpdate(username, event, details) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGS_SHEET_UPDATE);
  const timestamp = new Date().toString();
  sheet.appendRow([timestamp, username, event, details]);
}

// PROCESS CLIENT BILL FORM SUBMISSION
function processClientBill(formObject, username) {
  if (formObject.recId && checkClientBillId(formObject.recId)) {
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
    console.log("updated");
    updateClientBillRecord(values, updateRange, username); // Pass the username
  } else {
    console.log("Creating new record"); // Debugging line
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
function updateClientBillRecord(values, range, username) {
  try {
    const sheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
    const oldValues = sheet.getRange(range).getValues()[0]; // Get the old values before updating
    sheet.getRange(range).setValues(values); // Update the record

    // Log the changes
    const newValues = values[0]; // Get the new values
    const changes = [];

    // Compare old and new values to identify changes
    for (let i = 0; i < oldValues.length; i++) {
      if (oldValues[i] !== newValues[i]) {
        changes.push(`Column ${i + 1}: ${oldValues[i]} -> ${newValues[i]}`);
      }
    }

    if (changes.length > 0) {
      const details = `Updated record ${newValues[0]} (Bill No: ${
        newValues[1]
      }). Changes: ${changes.join(", ")}`;
      console.log("logs insert");
      logEventUpdate(username, "Update", details);
    }
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
function updateClientBill(formObject, username) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch all rows
  const rowIndex = data.findIndex(
    (row) => row[COLUMN_NAMES.recId] === formObject.recId
  ); // Find the row with the matching recId

  if (rowIndex !== -1) {
    // Fetch the old values before updating
    const oldValues = data[rowIndex];

    // Prepare the new values
    const newValues = [
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
      new Date().toLocaleString(), // Update the "Date Updated" field
      oldValues[COLUMN_NAMES.dateCreated], // Keep the original "Date Created" value
    ];

    // Update the row in the sheet
    sheet.getRange(rowIndex + 2, 1, 1, newValues.length).setValues([newValues]);

    // Compare old and new values to identify changes
    const changes = [];
    for (const [columnName, columnIndex] of Object.entries(COLUMN_NAMES)) {
      if (oldValues[columnIndex] !== newValues[columnIndex]) {
        changes.push(
          `${columnName}: ${oldValues[columnIndex]} -> ${newValues[columnIndex]}`
        );
      }
    }

    // Log the changes
    if (changes.length > 0) {
      const details = `Updated client bill: (Bill No: ${
        formObject.billNo
      }). Changes: ${changes.join(", ")}`;
      logEventUpdate(username, "Update Client Bill", details);
    }
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
  const clientBillSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const itemSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Item");

  // Fetch client bill details
  const clientBillData = clientBillSheet
    .getRange(CLIENT_BILL_RANGE)
    .getValues(); // Fetch columns A to I
  const clientBill = clientBillData.find((row) => row[0] === recId); // Find the record with the matching recId

  // Fetch associated items
  const itemData = itemSheet.getRange("A2:D").getValues(); // Fetch columns A to D
  const items = itemData.filter((row) => row[3] === clientBill[1]); // Filter items with matching Bill No

  return {
    clientBill: clientBill,
    items: items,
  };
}

//ABBREVIATIONS

// SAVE ABBREVIATION
function saveAbbreviation(formObject) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);
  const data = sheet.getRange(ABBREVIATIONS_RANGE).getValues(); // Fetch columns A and B
  const rowIndex = data.findIndex((row) => row[0] === formObject.name); // Find the row with the matching name

  if (rowIndex !== -1) {
    // Update existing abbreviation
    sheet
      .getRange(rowIndex + 2, 1, 1, 2)
      .setValues([[formObject.name, formObject.value]]);
  } else {
    // Create new abbreviation
    sheet.appendRow([formObject.name, formObject.value]);
  }
}

// GET ABBREVIATIONS
function getAbbreviations() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);
  const data = sheet.getRange(ABBREVIATIONS_RANGE).getValues();
  const filteredData = data.filter((row) => row[0] && row[1]);

  const uniqueData = Array.from(new Set(filteredData.map(JSON.stringify))).map(
    JSON.parse
  );

  uniqueData.sort((a, b) => a[0].localeCompare(b[0]));

  Logger.log(uniqueData);
  return uniqueData;
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

//CONTAINER NO

function getAllContainerRecords() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues();
  return data.filter((row) => row.some((cell) => cell !== "")); // Filter out completely empty rows
}

function updateContainer(originalContainer, newContainer) {
  const clientBillSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const expenseSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);

  const clientBillData = clientBillSheet.getDataRange().getValues();
  const expenseData = expenseSheet.getDataRange().getValues();

  // Update Client Bill Sheet
  for (let i = 1; i < clientBillData.length; i++) {
    if (clientBillData[i][8] === originalContainer) {
      // Assuming container number is at index 8
      clientBillSheet.getRange(i + 1, 9).setValue(newContainer); // Update container number (column index 9)
    }
  }

  // Update Expense Sheet
  for (let i = 1; i < expenseData.length; i++) {
    if (expenseData[i][1] === originalContainer) {
      // Assuming container number is at index 1
      expenseSheet.getRange(i + 1, 2).setValue(newContainer); // Update container number (column index 2)
    }
  }
}

//EXPENSES

// Fetch unique container numbers from ClientBill sheet
function getContainers() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange("I2:I").getValues().flat().filter(String);
  return [...new Set(data)]; // Remove duplicates
}

// Save new expense entry
function saveExpense(data) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);
  sheet.appendRow([
    generateUniqueId(),
    data.container,
    data.type,
    data.description,
    data.amount,
    new Date().toString(),
  ]);
  return "Expense added successfully!";
}

function updateExpense(data, username) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);
  const dataRange = sheet.getDataRange().getValues();

  for (let i = 1; i < dataRange.length; i++) {
    if (dataRange[i][0] === data.id) {
      // Fetch the old values before updating
      const oldValues = dataRange[i];

      // Update the expense
      sheet
        .getRange(i + 1, 2, 1, 4)
        .setValues([
          [data.container, data.type, data.description, data.amount],
        ]);

      // Compare old and new values to identify changes
      const changes = [];
      if (oldValues[1] !== data.container) {
        changes.push(`Container: ${oldValues[1]} -> ${data.container}`);
      }
      if (oldValues[2] !== data.type) {
        changes.push(`Type: ${oldValues[2]} -> ${data.type}`);
      }
      if (oldValues[3] !== data.description) {
        changes.push(`Description: ${oldValues[3]} -> ${data.description}`);
      }
      if (oldValues[4] !== data.amount) {
        changes.push(`Amount: ${oldValues[4]} -> ${data.amount}`);
      }

      // Log the changes
      if (changes.length > 0) {
        const details = `Updated expense ${
          data.container
        }. Changes: ${changes.join(", ")}`;
        logEventUpdate(username, "Update Expense", details);
      }

      return "Expense updated successfully!";
    }
  }
  return "Expense not found.";
}

// Fetch expenses with optional container filter
function getExpenses() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);
  const data = sheet.getRange(EXPENSES_SHEET_RANGE).getValues(); // Fetch columns A to F
  return data.filter((row) => row.some((cell) => cell !== "")); // Filter out completely empty rows
}

// Edit an existing expense
function editExpense(index, updatedData) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);
  const row = index + 2; // Adjust for header row
  sheet
    .getRange(row, 2, 1, 3)
    .setValues([
      [
        updatedData.container,
        updatedData.type,
        updatedData.description,
        updatedData.amount,
      ],
    ]);
  return "Expense updated!";
}

//RECEIPT

// GENERATE RECEIPT
function generateReceipt(recId) {
  const clientBillSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = clientBillSheet.getRange(CLIENT_BILL_RANGE).getValues(); // Fetch columns A to I
  const record = data.find((row) => row[0] === recId); // Find the record with the matching recId

  if (!record) {
    throw new Error("Record not found");
  }

  // Template ID (replace with your Google Docs template ID)
  const templateId = "1XI-qUgaCUWdP5J4RSSgbvBK4Ndycu-LuRbW8GjPnlWQ";

  // Create a copy of the template
  const templateDoc = DriveApp.getFileById(templateId);
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const newDoc = templateDoc.makeCopy(
    `Receipt_${record[1]}_${timestamp}`,
    DriveApp.getFolderById("1RDX1N7o6RPFx6pr_bVwiduQSMYaD-F2_")
  );
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // Replace placeholders with actual values
  body.replaceText("{{BillNo}}", record[1]);
  body.replaceText("{{ShipperName}}", record[2]);
  body.replaceText("{{ShipperTel}}", record[3]);
  body.replaceText("{{ReceiverName1}}", record[4]);
  body.replaceText("{{PhoneNo1}}", record[5]);
  body.replaceText("{{ReceiverName2}}", record[6]);
  body.replaceText("{{PhoneNo2}}", record[7]);
  body.replaceText("{{ContainerNo}}", record[8]);
  body.replaceText("{{TotalPieces}}", record[9]);
  body.replaceText("{{ActualWeight}}", record[10]);
  body.replaceText("{{DiscountWeight}}", record[11]);
  body.replaceText("{{ChargeableWeight}}", record[12]);
  body.replaceText("{{RatePerKg}}", record[13]);
  body.replaceText("{{BillCharge}}", record[14]);
  body.replaceText("{{DiscountCharge}}", record[15]);
  body.replaceText("{{TotalCharges}}", record[16]);
  body.replaceText("{{PaidAmount}}", record[17]);
  body.replaceText("{{OutstandingBalance}}", record[18]);

  const currentDate = new Date();
  const date = currentDate.toLocaleDateString(); // Get the date part
  const time = currentDate.toLocaleTimeString(); // Get the time part

  // Replace the {{Date}} placeholder with the formatted string
  body.replaceText("{{Date}}", date + "\n" + time);
  body.replaceText("{{AMOUNTWORDS}}", numberToWords(record[16]));

  // Save and close the document
  doc.saveAndClose();

  // Export as PDF
  const pdfBlob = newDoc.getAs(MimeType.PDF);
  const folder = DriveApp.getFolderById("1RDX1N7o6RPFx6pr_bVwiduQSMYaD-F2_"); // Desired folder
  const pdfFile = folder.createFile(pdfBlob); // Save the PDF in the specific folder

  // Set sharing permissions for the PDF
  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  // Generate preview and download links for the PDF
  const pdfPreviewUrl = `https://drive.google.com/file/d/${pdfFile.getId()}/preview`;
  const pdfDownloadUrl = `https://drive.google.com/uc?export=download&id=${pdfFile.getId()}`;

  // Delete the temporary document
  DriveApp.getFileById(newDocId).setTrashed(true);

  // Return both links as an object
  return {
    previewUrl: pdfPreviewUrl,
    downloadUrl: pdfDownloadUrl,
  };
}

//HOUSE WAY BILL

// GENERATE HOUSE WAY
function generateHouseWaybill(recId) {
  const clientBillSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = clientBillSheet.getRange(CLIENT_BILL_RANGE).getValues();

  // Find the record matching recId
  const record = data.find((row) => row[0] === recId); // recId is in column A

  if (!record) {
    throw new Error("Record not found");
  }

  const billNo = record[1]; // BillNo is in column B

  // Fetch Items Data
  const itemsSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ITEM_DATA_SHEET);
  const abbreviationsSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABBREVIATIONS_SHEET);

  const itemsData = itemsSheet.getDataRange().getValues();
  const abbreviationsData = abbreviationsSheet.getDataRange().getValues();

  // Create abbreviation dictionary to map short form to full form
  let abbreviationMap = {};
  for (let i = 1; i < abbreviationsData.length; i++) {
    let shortForm = abbreviationsData[i][1]; // Short form in Column B of Abbreviations
    let fullForm = abbreviationsData[i][0]; // Full form in Column A of Abbreviations
    abbreviationMap[shortForm] = fullForm; // { "Short Form": "Full Form" }
  }

  // Count occurrences of each item type
  let itemCountMap = {};
  for (let i = 1; i < itemsData.length; i++) {
    if (itemsData[i][3] === billNo) {
      // Assuming Bill No is in Column D (Index 3)
      let itemType = itemsData[i][1]; // Item Type (short form in Column B of Items sheet)
      let fullItemType = abbreviationMap[itemType] || itemType; // Convert short form to full form if found
      itemCountMap[fullItemType] = (itemCountMap[fullItemType] || 0) + 1; // Count item type
    }
  }

  // Format Items List based on the count
  let itemsList = [];
  for (let itemType in itemCountMap) {
    let count = itemCountMap[itemType];
    itemsList.push(`${count} ${itemType}`); // Display number of items and their full form
  }

  if (itemsList.length === 0) {
    throw new Error("No items found for the selected Bill No.");
  }

  // Join items list into a single formatted string
  let itemsFormatted = itemsList.join("\n");

  // House Waybill Template
  const templateId = "1pVQdnDmbE0OHMd5ElzOPx2lshECyIyz13cevk3iCOd4"; // Template Doc ID
  const folderId = "1Co0m5ScDdtoRNn2KysMC46O7YlZyoyF9"; // Drive Folder ID

  // Copy Template
  const templateDoc = DriveApp.getFileById(templateId);
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const newDoc = templateDoc.makeCopy(
    `HouseWaybill_${billNo}_${timestamp}`,
    DriveApp.getFolderById(folderId)
  );
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // Replace placeholders
  body.replaceText("{{BillNo}}", billNo);
  body.replaceText("{{ItemsList}}", itemsFormatted);
  body.replaceText("{{Date}}", new Date().toLocaleString());

  // Other fields from Client Bill
  body.replaceText("{{ShipperName}}", record[2]);
  body.replaceText("{{ShipperTel}}", record[3]);
  body.replaceText("{{ReceiverName1}}", record[4]);
  body.replaceText("{{PhoneNo1}}", record[5]);
  body.replaceText("{{ReceiverName2}}", record[6]);
  body.replaceText("{{PhoneNo2}}", record[7]);
  body.replaceText("{{ContainerNo}}", record[8]);
  body.replaceText("{{TotalPieces}}", record[9]);
  body.replaceText("{{ActualWeight}}", record[10]);
  body.replaceText("{{DiscountWeight}}", record[11]);
  body.replaceText("{{ChargeableWeight}}", record[12]);
  body.replaceText("{{RatePerKg}}", record[13]);
  body.replaceText("{{BillCharge}}", record[14]);
  body.replaceText("{{DiscountCharge}}", record[15]);
  body.replaceText("{{TotalCharges}}", record[16]);
  body.replaceText("{{PaidAmount}}", record[17]);
  body.replaceText("{{OutstandingBalance}}", record[18]);

  // Save and Close Doc
  doc.saveAndClose();

  // Export as PDF
  const pdfBlob = newDoc.getAs(MimeType.PDF);
  const folder = DriveApp.getFolderById("1Co0m5ScDdtoRNn2KysMC46O7YlZyoyF9"); // Desired folder
  const pdfFile = folder.createFile(pdfBlob); // Save the PDF in the specific folder

  // Delete the temporary document
  DriveApp.getFileById(newDocId).setTrashed(true);

  // Set sharing permissions for the PDF
  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  // Return preview & download links
  return {
    previewUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/preview`,
    downloadUrl: `https://drive.google.com/uc?export=download&id=${pdfFile.getId()}`,
  };
}

//MANIFEST

//funstion to generate the manifest list
function generateManifestList(recId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const clientBillSheet = ss.getSheetByName(CLIENT_BILL_SHEET);
  const itemsSheet = ss.getSheetByName(ITEM_DATA_SHEET);
  const abbreviationsSheet = ss.getSheetByName(ABBREVIATIONS_SHEET);
  const manifestTemplateSheet = ss.getSheetByName("Manifest");

  const clientBillData = clientBillSheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();
  const abbreviationsData = abbreviationsSheet.getDataRange().getValues();

  const abbreviationMap = {};
  for (let i = 1; i < abbreviationsData.length; i++) {
    abbreviationMap[abbreviationsData[i][1]] = abbreviationsData[i][0];
  }

  let containerNo = null;
  for (let i = 1; i < clientBillData.length; i++) {
    if (clientBillData[i][0] === recId) {
      containerNo = clientBillData[i][8];
      break;
    }
  }
  if (!containerNo) throw new Error("No Container No found for the provided RecId.");

  const bills = clientBillData.filter(row => row[8] === containerNo);
  if (bills.length === 0) throw new Error("No bills found for the container.");

  let manifestList = [];
  let totalPieces = 0;
  let totalWeight = 0;

  bills.forEach((bill, index) => {
    const billNo = bill[1];
    const shipperName = bill[2];
    const shipperTel = bill[3];
    const receiverName = bill[4];
    const receiverTel = bill[5];
    const receiverName2 = bill[6];
    const receiverTel2 = bill[7];
    const totalPiecesForBill = bill[9];
    const totalWeightForBill = bill[10];

    let itemCountMap = {};
    for (let i = 1; i < itemsData.length; i++) {
      if (itemsData[i][3] === billNo) {
        const itemType = abbreviationMap[itemsData[i][1]] || itemsData[i][1];
        itemCountMap[itemType] = (itemCountMap[itemType] || 0) + 1;
      }
    }

    const itemsList = Object.entries(itemCountMap)
      .map(([type, count]) => `${count} ${type}`)
      .join(", ");

    manifestList.push([
      index + 1,
      billNo,
      `${shipperName}\n${receiverName}${receiverName2 ? '\n' + receiverName2 : ''}`.replace(/\//g, ''),
      `${shipperTel}\n${receiverTel}${receiverTel2 ? '\n' + receiverTel2 : ''}`.replace(/\//g, ''),
      itemsList,
      totalPiecesForBill,
      totalWeightForBill
    ]);

    totalPieces += totalPiecesForBill;
    totalWeight += totalWeightForBill;
  });

  const folder = DriveApp.getFolderById("1PDPiUFkyO0vM0yM_0Gir4FU41PKZkEI1");
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const newSS = ss.copy(`Manifest_${containerNo}_${timestamp}`);
  const newSSId = newSS.getId();

  newSS.getSheets().forEach(sheet => {
    if (sheet.getName() !== "Manifest") newSS.deleteSheet(sheet);
  });

  const manifestSheet = newSS.getSheetByName("Manifest");
  manifestSheet.getRange("E5").setValue(containerNo);
  manifestSheet.getDataRange().createTextFinder("{{ContainerNo}}").replaceAllWith(containerNo);

  const startRow = 9;
  const rowCount = manifestList.length;
  const fontSize = 12;

  if (rowCount > 0) {
    manifestSheet.getRange(startRow, 1, rowCount, 7).setValues(manifestList);

    for (let col = 1; col <= 7; col++) {
      manifestSheet.getRange(startRow, col, rowCount, 1)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true)
        .setFontSize(fontSize);
    }

    const totalRow = startRow + rowCount;
    manifestSheet.getRange(totalRow, 1, 1, 5).merge();
    manifestSheet.getRange(totalRow, 1).setValue("TOTAL")
      .setHorizontalAlignment("center")
      .setFontWeight("bold")
      .setFontSize(fontSize + 2);

    manifestSheet.getRange(totalRow, 6).setValue(totalPieces)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    manifestSheet.getRange(totalRow, 7).setValue(totalWeight)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    manifestSheet.autoResizeRows(startRow, rowCount + 1);
  }

  const lastUsedRow = manifestSheet.getLastRow();
  const maxRows = manifestSheet.getMaxRows();
  if (lastUsedRow < maxRows) {
    manifestSheet.deleteRows(lastUsedRow + 1, maxRows - lastUsedRow);
  }

  // Build custom PDF URL
  const exportUrl = `https://docs.google.com/spreadsheets/d/${newSSId}/export?format=pdf&size=A4&portrait=true&fitw=true` +
                    `&top_margin=0.2&bottom_margin=0.4&left_margin=0.2&right_margin=0.2` +
                    `&sheetnames=false&printtitle=false&gridlines=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  const pdfBlob = response.getBlob().setName(`Manifest_${containerNo}_${timestamp}.pdf`);
  const pdfFile = folder.createFile(pdfBlob);

  // Trash the spreadsheet copy
  DriveApp.getFileById(newSSId).setTrashed(true);

  // Make it viewable
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    previewUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/preview`,
    downloadUrl: `https://drive.google.com/uc?id=${pdfFile.getId()}&export=download`
  };
}

// LOADING LIST
function generateLoadingList(recId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const clientBillSheet = ss.getSheetByName(CLIENT_BILL_SHEET);
  const itemsSheet = ss.getSheetByName(ITEM_DATA_SHEET);
  const templateSheet = ss.getSheetByName("Loading List");

  const clientBillData = clientBillSheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();

  // Find container number
  let containerNo = null;
  for (let i = 1; i < clientBillData.length; i++) {
    if (clientBillData[i][0] === recId) {
      containerNo = clientBillData[i][8];
      break;
    }
  }
  if (!containerNo) throw new Error("No Container No found for the provided RecId.");

  const bills = clientBillData.filter(row => row[8] === containerNo);
  if (bills.length === 0) throw new Error("No bills found for the container.");

  // Prepare loading list data
  let loadingList = [];
  let totalPieces = 0;
  let totalWeight = 0;
  let rowIndex = 1;

  bills.forEach((bill) => {
    const billNo = bill[1];
    const totalPiecesForBill = bill[9];
    const totalWeightForBill = bill[10];

    // Get items for this bill
    const billItems = itemsData
      .filter(item => item[3] === billNo)
      .map(item => item[1]);

    const itemCount = billItems.length;
    let position = 1;

    // Split into chunks of 20 items
    for (let i = 0; i < itemCount; i += 20) {
      const chunk = billItems.slice(i, i + 20);
      let row = [];

      if (i === 0) {
        row = [rowIndex++, billNo, totalPiecesForBill, totalWeightForBill];
        totalPieces += totalPiecesForBill;
        totalWeight += totalWeightForBill;
      } else {
        row = ["", "", "", ""];
      }

      // Add items to row
      chunk.forEach(item => {
        row.push(`${position} ${item}`);
        position++;
      });
      
      // Fill remaining columns if needed
      while (row.length < 24) row.push("");
      
      loadingList.push(row);
    }
  });

  // Create new spreadsheet
  const folder = DriveApp.getFolderById("1TmbOGmOxcnfo5aiP6g7O50FmdznAEwhH");
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const newSS = ss.copy(`LoadingList_${containerNo}_${timestamp}`);
  const newSSId = newSS.getId();

  // Remove other sheets
  newSS.getSheets().forEach(sheet => {
    if (sheet.getName() !== "Loading List") newSS.deleteSheet(sheet);
  });

  const sheet = newSS.getSheetByName("Loading List");

  // Replace placeholders
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  sheet.getDataRange().createTextFinder("{{ContainerNo}}").replaceAllWith(containerNo);
  sheet.getDataRange().createTextFinder("{{DATE}}").replaceAllWith(dateStr);

  // Write data with formatting
  const startRow = 9;
  const numCols = 24;
  
  if (loadingList.length > 0) {
    // Write data
    sheet.getRange(startRow, 1, loadingList.length, numCols)
      .setValues(loadingList)
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    // Set optimized column widths (in pixels)
    const columnWidths = {
      1: 30,    // SR NO (Column A)
      2: 100,   // Bill No (Column B)
      3: 70,    // Pieces (Column C)
      4: 70,    // Weight (Column D)
      5: 60,    // Column E (reduced width)
      14: 60,   // Column N (reduced width)
      19: 60    // Column S (reduced width)
    };

    // Apply specific column widths
    Object.entries(columnWidths).forEach(([col, width]) => {
      sheet.setColumnWidth(Number(col), width);
    });

    // Auto-resize other columns (excluding the ones we manually set)
    const excludedCols = new Set(Object.keys(columnWidths).map(Number));
    for (let col = 1; col <= numCols; col++) {
      if (!excludedCols.has(col)) {
        sheet.autoResizeColumn(col);
      }
    }

    // Add TOTAL row
    const totalRow = startRow + loadingList.length;
    sheet.getRange(totalRow, 1, 1, 2).merge();
    sheet.getRange(totalRow, 1)
      .setValue("TOTAL")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center");
      
    sheet.getRange(totalRow, 3)
      .setValue(totalPieces)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
      
    sheet.getRange(totalRow, 4)
      .setValue(totalWeight)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    // Auto-resize rows
    sheet.autoResizeRows(startRow, loadingList.length + 2);
  }

  // Remove excess rows
  const lastUsedRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (lastUsedRow < maxRows) {
    sheet.deleteRows(lastUsedRow + 1, maxRows - lastUsedRow);
  }

  // Generate PDF (Landscape)
  const pdfUrl = `https://docs.google.com/spreadsheets/d/${newSSId}/export?` +
    `format=pdf&` +
    `size=A4&` +
    `portrait=false&` +
    `fitw=true&` +
    `top_margin=0.25&` +
    `bottom_margin=0.25&` +
    `left_margin=0.25&` +
    `right_margin=0.25&` +
    `sheetnames=false&` +
    `printtitle=false&` +
    `gridlines=false&` +
    `range=A1:${sheet.getLastColumn()}${sheet.getLastRow()}`;

  const token = ScriptApp.getOAuthToken();
  const pdfResponse = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const pdfBlob = pdfResponse.getBlob().setName(`LoadingList_${containerNo}_${timestamp}.pdf`);
  const pdfFile = folder.createFile(pdfBlob);

  // Generate XLS file
  const xlsUrl = `https://docs.google.com/spreadsheets/d/${newSSId}/export?format=xlsx`;
  const xlsResponse = UrlFetchApp.fetch(xlsUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const xlsBlob = xlsResponse.getBlob().setName(`LoadingList_${containerNo}_${timestamp}.xlsx`);
  const xlsFile = folder.createFile(xlsBlob);

  // Clean up
  DriveApp.getFileById(newSSId).setTrashed(true);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  xlsFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    previewUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/preview`,
    downloadUrl: `https://drive.google.com/uc?id=${pdfFile.getId()}&export=download`,
    xlsDownloadUrl: `https://drive.google.com/uc?id=${xlsFile.getId()}&export=download`
  };
}

//ITEM LIST COUNT

function generateItemList(recId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const clientBillSheet = ss.getSheetByName(CLIENT_BILL_SHEET);
  const itemsSheet = ss.getSheetByName(ITEM_DATA_SHEET);
  const templateSheet = ss.getSheetByName("Item List");

  const clientBillData = clientBillSheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();

  let containerNo = null;
  for (let i = 1; i < clientBillData.length; i++) {
    if (clientBillData[i][0] === recId) {
      containerNo = clientBillData[i][8];
      break;
    }
  }
  if (!containerNo) throw new Error("No Container No found for the provided RecId.");

  const bills = clientBillData.filter(row => row[8] === containerNo);
  if (bills.length === 0) throw new Error("No bills found for the container.");

  let itemCounts = {};
  let totalPieces = 0;
  let totalItemCounts = {};

  bills.forEach((bill) => {
    const billNo = bill[1];
    const totalPiecesForBill = bill[9];
    totalPieces += totalPiecesForBill;

    if (!itemCounts[billNo]) itemCounts[billNo] = { totalPieces: totalPiecesForBill };

    for (let i = 1; i < itemsData.length; i++) {
      if (itemsData[i][3] === billNo) {
        const itemType = itemsData[i][1];
        itemCounts[billNo][itemType] = (itemCounts[billNo][itemType] || 0) + 1;
        totalItemCounts[itemType] = (totalItemCounts[itemType] || 0) + 1;
      }
    }
  });

  const folderId = "1l1n66G2MPgEmlFB5iHN7ZUbijgbx7k8y";
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const newSS = ss.copy(`ItemList_${containerNo}_${timestamp}`);
  const newSSId = newSS.getId();
  const sheet = newSS.getSheetByName("Item List");

  // Remove all other sheets
  newSS.getSheets().forEach(s => {
    if (s.getName() !== "Item List") newSS.deleteSheet(s);
  });

  // Replace placeholders
  sheet.createTextFinder("{{containerNo}}").matchCase(false).replaceAllWith(containerNo);
  sheet.createTextFinder("{{date}}").matchCase(false).replaceAllWith(new Date().toLocaleDateString());

  // Table generation
  const itemTypes = Object.keys(totalItemCounts);
  const headers = ["HWB NO", "TTL PCS", ...itemTypes];
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]).setFontSize(14).setFontWeight("bold");

  const startRow = 9;
  const data = [];
  Object.entries(itemCounts).forEach(([billNo, counts]) => {
    const row = [billNo, counts.totalPieces];
    itemTypes.forEach(type => row.push(counts[type] || ""));
    data.push(row);
  });

  if (data.length > 0) {
    sheet.getRange(startRow, 1, data.length, headers.length).setValues(data).setFontSize(13);

    const totalRowIdx = startRow + data.length;
    const totalRow = ["TOTAL", totalPieces];
    itemTypes.forEach(type => totalRow.push(totalItemCounts[type] || ""));
    sheet.getRange(totalRowIdx, 1, 1, headers.length).setValues([totalRow]).setFontSize(13).setFontWeight("bold");

    sheet.getRange(startRow, 1, data.length + 1, headers.length)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true);
  }

  // Delete extra rows to avoid white space
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (lastRow < maxRows) sheet.deleteRows(lastRow + 1, maxRows - lastRow);

  // Export the sheet to PDF with custom margins (top=0.2in, left=0.2in, right=0.2in)
  const exportUrl = `https://docs.google.com/spreadsheets/d/${newSSId}/export?` +
    `format=pdf&portrait=true&sheetnames=false&printtitle=false&pagenumbers=false&` +
    `gridlines=false&fzr=false&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2&` +
    `gid=${sheet.getSheetId()}`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  const blob = response.getBlob().setName(`ItemList_${containerNo}_${timestamp}.pdf`);
  const pdfFile = DriveApp.getFolderById(folderId).createFile(blob);
  DriveApp.getFileById(newSSId).setTrashed(true);

  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    previewUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/preview`,
    downloadUrl: `https://drive.google.com/uc?id=${pdfFile.getId()}&export=download`,
  };
}

// Generate report
function generateReport({ reportType, container, fromDate, toDate }) {
  const clientBillSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const expensesSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);

  // Fetch client bill data
  const clientBillData = clientBillSheet.getDataRange().getValues();
  const expensesData = expensesSheet.getDataRange().getValues();

  // Filter data based on container or date range
  let filteredClientBills = [];
  let filteredExpenses = [];

  if (container) {
    // Filter by container
    filteredClientBills = clientBillData.filter((row) => row[8] === container); // Container No is in column I
    filteredExpenses = expensesData.filter((row) => row[1] === container); // Container No is in column B
  } else {
    // Filter by date range
    const from = new Date(fromDate);
    const to = new Date(toDate);

    filteredClientBills = clientBillData.filter((row) => {
      const date = new Date(row[19]); // Date is in column T
      return date >= from && date <= to;
    });

    filteredExpenses = expensesData.filter((row) => {
      const date = new Date(row[5]); // Date is in column F
      return date >= from && date <= to;
    });
  }

  // Calculate totals
  const totalIncome = filteredClientBills.reduce(
    (sum, row) => sum + parseFloat(row[16] || 0),
    0
  ); // Total Charges is in column Q
  const totalExpenses = filteredExpenses.reduce(
    (sum, row) => sum + parseFloat(row[4] || 0),
    0
  ); // Amount is in column E
  const profitLoss = totalIncome - totalExpenses;

  // Generate the report document
  const templateId = REPORT_TEMPLATE_ID;
  const folderId = REPORT_FOLDER_ID;
  const templateDoc = DriveApp.getFileById(templateId);
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const newDoc = templateDoc.makeCopy(
    `Report_${container || "All"}_${timestamp}`,
    DriveApp.getFolderById(folderId)
  );
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // Clear the template body (remove placeholders)
  body.clear();

  // Add the title (centered and blue)
  const title = body.appendParagraph("Eastern Joury Est - Report");
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setForegroundColor("#1a73e8"); // Blue color

  // Add container info or date range (right-aligned)
  const containerInfo = container
    ? `Container: ${container}`
    : `Date Range: ${fromDate} to ${toDate}`;
  const infoParagraph = body.appendParagraph(containerInfo);
  infoParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  // Add totals and other details (left-aligned)
  const totalsParagraph = body.appendParagraph(
    `Total Income: ${totalIncome.toFixed(2)}\n` +
      `Total Expenses: ${totalExpenses.toFixed(2)}`
  );
  totalsParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

  // Add profit or loss (green for profit, red for loss)
  const profitLossParagraph = body.appendParagraph(
    profitLoss >= 0
      ? `Profit: ${profitLoss.toFixed(2)}`
      : `Loss: ${Math.abs(profitLoss).toFixed(2)}`
  );
  profitLossParagraph.setForegroundColor(
    profitLoss >= 0 ? "#0f9d58" : "#db4437"
  ); // Green for profit, red for loss
  profitLossParagraph.setBold(true);

  // Add a section for income (client bills) if report type is "All" or "Income"
  if (reportType === "all" || reportType === "income") {
    if (filteredClientBills.length > 0) {
      body
        .appendParagraph("Income (Client Bills)")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);

      // Create a table for client bills
      const clientBillTable = body.appendTable();
      const headerRow = clientBillTable.appendTableRow();
      headerRow
        .appendTableCell("Bill No")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");
      headerRow
        .appendTableCell("Shipper Name")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");
      headerRow
        .appendTableCell("Tel No")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");
      headerRow
        .appendTableCell("Total Weight")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");
      headerRow
        .appendTableCell("Total Amount")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");

      filteredClientBills.forEach((row) => {
        const dataRow = clientBillTable.appendTableRow();
        dataRow.appendTableCell(row[1]).setForegroundColor("#000000"); // Bill No (black)
        dataRow.appendTableCell(row[2]).setForegroundColor("#000000"); // Shipper Name (black)
        dataRow.appendTableCell(row[3]).setForegroundColor("#000000"); // Tel No (black)
        dataRow.appendTableCell(row[10]).setForegroundColor("#000000"); // Total Weight (black)
        dataRow.appendTableCell(row[16]).setForegroundColor("#000000"); // Total Charges (black)
      });

      // Add total income
      body
        .appendParagraph(`Total Income: ${totalIncome.toFixed(2)}`)
        .setBold(true);
    } else {
      body
        .appendParagraph("No income data found for the selected criteria.")
        .setItalic(true);
    }
  }

  // Add a section for expenses if report type is "All" or "Expense"
  if (reportType === "all" || reportType === "expense") {
    if (filteredExpenses.length > 0) {
      body
        .appendParagraph("Expenses")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);

      // Create a table for expenses
      const expenseTable = body.appendTable();
      const headerRow = expenseTable.appendTableRow();
      headerRow
        .appendTableCell("Container")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");
      headerRow
        .appendTableCell("Type")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");
      headerRow
        .appendTableCell("Description")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");
      headerRow
        .appendTableCell("Amount")
        .setBackgroundColor("#1a73e8")
        .setForegroundColor("#ffffff");

      filteredExpenses.forEach((row) => {
        const dataRow = expenseTable.appendTableRow();
        dataRow.appendTableCell(row[1]).setForegroundColor("#000000"); // Container (black)
        dataRow.appendTableCell(row[2]).setForegroundColor("#000000"); // Type (black)
        dataRow.appendTableCell(row[3]).setForegroundColor("#000000"); // Description (black)
        dataRow.appendTableCell(row[4]).setForegroundColor("#000000"); // Amount (black)
      });

      // Add total expenses
      body
        .appendParagraph(`Total Expenses: ${totalExpenses.toFixed(2)}`)
        .setBold(true);
    } else {
      body
        .appendParagraph("No expense data found for the selected criteria.")
        .setItalic(true);
    }
  }

  // Save and close the document
  doc.saveAndClose();

  // Export as PDF
  const pdfBlob = newDoc.getAs(MimeType.PDF);
  const folder = DriveApp.getFolderById(folderId);
  const pdfFile = folder.createFile(pdfBlob);
  DriveApp.getFileById(newDocId).setTrashed(true);
  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  // Return preview and download URLs
  return {
    previewUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/preview`,
    downloadUrl: `https://drive.google.com/uc?export=download&id=${pdfFile.getId()}`,
  };
}


function generateContainerSummary(params) {
  const { container, fromDate, toDate } = params;
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const clientBillSheet = ss.getSheetByName(CLIENT_BILL_SHEET);
  const summaryTemplateSheet = ss.getSheetByName("Container Summary");
  
  if (!summaryTemplateSheet) {
    throw new Error("Container Summary template sheet not found");
  }

  // Get all data including headers
  const clientBillData = clientBillSheet.getDataRange().getValues();
  const headers = clientBillData[0];
  console.log("Headers:", headers);
  
  // Filter bills (skip header row)
  let bills = [];
  let headerText = "";
  
  if (container) {
    bills = clientBillData.slice(1).filter(row => row[8] && row[8].toString().trim() === container.toString().trim());
    headerText = `CONTAINER NO: ${container}`;
    console.log(`Found ${bills.length} bills for container ${container}`);
  } else {
    const from = new Date(fromDate);
    const to = new Date(toDate);
    
    bills = clientBillData.slice(1).filter(row => {
      try {
        const billDate = row[19] ? new Date(row[19]) : (row[20] ? new Date(row[20]) : null);
        return billDate && billDate >= from && billDate <= to;
      } catch (e) {
        console.warn("Date error:", e);
        return false;
      }
    });
    headerText = `DATE RANGE: ${formatDateForDisplay(fromDate)} TO ${formatDateForDisplay(toDate)}`;
    console.log(`Found ${bills.length} bills for date range`);
  }
  
  if (bills.length === 0) {
    throw new Error(`No matching bills found`);
  }

  // CREATE NEW SPREADSHEET
  const folder = DriveApp.getFolderById("1Mevmd81eoXttjgPMHKthOrI29hue0ng6");
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const newSSName = `Container_Summary_${container || "DATERANGE"}_${timestamp}`;
  const newSS = SpreadsheetApp.create(newSSName);
  const newSSId = newSS.getId();
  
  // COPY TEMPLATE
  const copiedSheet = summaryTemplateSheet.copyTo(newSS);
  copiedSheet.setName("Container Summary");
  newSS.deleteSheet(newSS.getSheets()[0]); // Remove default sheet
  
  // GET COPIED SHEET
  const summarySheet = newSS.getSheetByName("Container Summary");
  
  // REPLACE PLACEHOLDER
  summarySheet.getDataRange().createTextFinder("{{Placeholder}}").replaceAllWith(headerText);
  
  // PREPARE DATA
  const summaryData = [];
  let totals = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // For pieces, kgs, charges, etc.
  
  bills.forEach((bill, index) => {
    const rowData = [
      index + 1,                           // SR NO
      bill[8],                             // CONTAINER NO
      bill[1],                             // HWB NO
      bill[9],                             // TOTAL PCS
      bill[10],                            // TOTAL KG'S
      bill[16],                            // TOTAL CHARGES
      bill[14],                            // BILL CHARGE
      bill[21],                            // BAREL CHG (adjust if needed)
      bill[16],                            // TOTAL CHARGE (duplicate)
      bill[15],                            // DISCOUNT AMOUNT
      bill[16],                            // TOTAL AMOUNT
      bill[17],                            // PAID AMOUNT
      bill[18],                            // DUE AMOUNT
      bill[2]                              // CLIENT NAME
    ];
    
    summaryData.push(rowData);
    
    // Update totals
    totals[0] += bill[9] || 0;    // PCS
    totals[1] += bill[10] || 0;   // KGS
    totals[2] += bill[16] || 0;   // CHARGES
    totals[3] += bill[14] || 0;   // BILL CHARGE
    totals[4] += 0;               // BAREL CHG
    totals[5] += bill[16] || 0;   // TOTAL
    totals[6] += bill[15] || 0;   // DISCOUNT
    totals[7] += bill[16] || 0;   // TOTAL
    totals[8] += bill[17] || 0;   // PAID
    totals[9] += bill[18] || 0;   // DUE
  });
  
  // WRITE DATA TO SHEET (STARTING AT ROW 9)
  if (summaryData.length > 0) {
    const startRow = 9; // Data starts at row 9
    const numCols = summaryData[0].length;
    
    // Clear any existing data below row 8
    if (summarySheet.getLastRow() >= startRow) {
      summarySheet.getRange(startRow, 1, summarySheet.getLastRow() - startRow + 1, numCols).clearContent();
    }
    
    // Write new data starting at row 9
    summarySheet.getRange(startRow, 1, summaryData.length, numCols).setValues(summaryData);
    
    // Add totals row
    const totalsRow = startRow + summaryData.length;
    summarySheet.getRange(totalsRow, 1).setValue("TOTAL")
      .setFontWeight("bold")
      .setFontSize(12); // Increased font size
    
    // Set totals values with formatting
    const totalsColumns = [4,5,6,7,8,9,10,11,12,13]; // Columns to format
    totalsColumns.forEach((col, idx) => {
      if (idx < totals.length) {
        summarySheet.getRange(totalsRow, col)
          .setValue(totals[idx])
          .setFontWeight("bold")
          .setFontSize(12) // Increased font size
          .setHorizontalAlignment("center");
      }
    });
    
    // Format totals row
    summarySheet.getRange(totalsRow, 1, 1, numCols)
      .setFontWeight("bold")
      .setBackground("#f2f2f2");
    
    // Remove all empty rows below data
    const maxRows = summarySheet.getMaxRows();
    const lastUsedRow = summarySheet.getLastRow();
    if (lastUsedRow < maxRows) {
      summarySheet.deleteRows(lastUsedRow + 1, maxRows - lastUsedRow);
    }
    
    // Auto-resize columns
    for (let i = 1; i <= numCols; i++) {
      //summarySheet.autoResizeColumn(i);
    }
  }
  
  // GENERATE PDF WITH PROPER SETTINGS
  const lastRow = summarySheet.getLastRow();
  const exportUrl = `https://docs.google.com/spreadsheets/d/${newSSId}/export?` +
    `format=pdf&` +
    `size=A4&` +
    `portrait=true&` +
    `fitw=true&` +
    `top_margin=0.2&` +
    `bottom_margin=0.4&` +
    `left_margin=0.2&` +
    `right_margin=0.2&` +
    `sheetnames=false&` +
    `printtitle=false&` +
    `gridlines=false&` +
    `range=A1:M${lastRow}`; // Explicitly include all rows
  
  let pdfBlob;
  try {
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`PDF generation failed: ${response.getContentText()}`);
    }
    
    pdfBlob = response.getBlob().setName(`${newSSName}.pdf`);
  } catch (e) {
    console.error("PDF generation error:", e);
    throw new Error("Failed to generate PDF. Please try again.");
  }
  
  const pdfFile = folder.createFile(pdfBlob);
  DriveApp.getFileById(newSSId).setTrashed(true);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return {
    previewUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/preview`,
    downloadUrl: `https://drive.google.com/uc?id=${pdfFile.getId()}&export=download`
  };
}

function formatDateForDisplay(dateString) {
  return Utilities.formatDate(new Date(dateString), Session.getScriptTimeZone(), "dd-MMM-yyyy");
}



//LOGS

function getLogs() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGS_SHEET_UPDATE);
  const data = sheet.getDataRange().getValues(); // Fetch all rows
  return data.slice(1); // Skip the header row
}

/*-------------------GENERAL-------------------------*/

// INCLUDE HTML PARTS (JS, CSS, OTHER HTML FILES)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function numberToWords(num) {
  if (num === 0) return "ZERO";
  
  // Split into whole and decimal parts
  const parts = num.toString().split('.');
  let whole = parseInt(parts[0]);
  let decimal = parts.length > 1 ? parseInt(parts[1].padEnd(2, '0').substring(0, 2)) : 0;

  const belowTwenty = [
    "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN",
    "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN",
    "EIGHTEEN", "NINETEEN"
  ];
  const tens = [
    "", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"
  ];
  const thousands = ["", "THOUSAND", "MILLION", "BILLION"];

  function helper(n) {
    if (n === 0) return "";
    if (n < 20) return belowTwenty[n - 1];
    if (n < 100) {
      return tens[Math.floor(n / 10)] + 
             (n % 10 === 0 ? "" : "-" + belowTwenty[(n % 10) - 1]);
    }
    if (n < 1000) {
      return belowTwenty[Math.floor(n / 100) - 1] + 
             " HUNDRED" + 
             (n % 100 === 0 ? "" : " AND " + helper(n % 100));
    }
    for (let i = 0; i < thousands.length; i++) {
      const unit = 1000 ** (i + 1);
      if (n < unit) {
        return helper(Math.floor(n / (1000 ** i))) + 
               " " + thousands[i] + 
               (n % (1000 ** i) === 0 ? "" : " " + helper(n % (1000 ** i)));
      }
    }
    return "";
  }

  let wholeText = helper(whole) || "ZERO";
  let decimalText = helper(decimal);
  
  if (decimal > 0) {
    return wholeText + " POINT " + decimalText;
  }
  
  return wholeText;
}