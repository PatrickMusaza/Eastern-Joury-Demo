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
  const folder = DriveApp.getFolderById("1RDX1N7o6RPFx6pr_bVwiduQSMYaD-F2_"); // Desired folder
  const pdfFile = folder.createFile(pdfBlob); // Save the PDF in the specific folder

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

//CONTAINER NO

function getAllContainerRecords() {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getRange(CLIENT_BILL_RANGE).getValues();
  return data.filter((row) => row.some((cell) => cell !== "")); // Filter out completely empty rows
}

function updateContainer(originalContainer, newContainer) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // Start from 1 to skip header row
    if (data[i][8] === originalContainer) {
      // Assuming container number is at index 8
      sheet.getRange(i + 1, 9).setValue(newContainer); // Update container number
    }
  }
}

//EXPENSES

function getAllData() {
  const expensesSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET); // Replace with your sheet name
  const containersSheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLIENT_BILL_SHEET); // Replace with your sheet name

  const expensesData = expensesSheet.getDataRange().getValues();
  const containersData = containersSheet.getDataRange().getValues();

  return {
    expenses: expensesData.slice(1), // Skip header row
    containers: containersData.map((row) => row[0]), // Assuming container numbers are in the first column
  };
}

function saveExpense(expenseData) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET); // Replace with your sheet name
  sheet.appendRow([
    Utilities.getUuid(), // Generate a unique ID for the expense
    expenseData.containerNo,
    expenseData.officeTransport,
    expenseData.containerTransportOffice,
    expenseData.containerLoadingLabour,
    // Add more expense fields here
    expenseData.date,
  ]);
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
  body.replaceText("{{Date}}", new Date().toLocaleString());
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

/*-------------------GENERAL-------------------------*/

// INCLUDE HTML PARTS (JS, CSS, OTHER HTML FILES)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function numberToWords(num) {
  if (num === 0) return "zero";

  const belowTwenty = [
    "one",
    "two",
    "three",
    "four",
    "five",
    "six",
    "seven",
    "eight",
    "nine",
    "ten",
    "eleven",
    "twelve",
    "thirteen",
    "fourteen",
    "fifteen",
    "sixteen",
    "seventeen",
    "eighteen",
    "nineteen",
  ];
  const tens = [
    "",
    "",
    "twenty",
    "thirty",
    "forty",
    "fifty",
    "sixty",
    "seventy",
    "eighty",
    "ninety",
  ];
  const thousands = ["", "thousand", "million", "billion"];

  function helper(n) {
    if (n < 20) return belowTwenty[n - 1];
    if (n < 100)
      return (
        tens[Math.floor(n / 10)] +
        (n % 10 === 0 ? "" : "-" + belowTwenty[(n % 10) - 1])
      );
    if (n < 1000)
      return (
        belowTwenty[Math.floor(n / 100) - 1] +
        " hundred" +
        (n % 100 === 0 ? "" : " and " + helper(n % 100))
      );
    for (let i = 0; i < thousands.length; i++) {
      const unit = 1000 ** (i + 1);
      if (n < unit)
        return (
          helper(Math.floor(n / 1000 ** i)) +
          " " +
          thousands[i] +
          (n % 1000 ** i === 0 ? "" : " " + helper(n % 1000 ** i))
        );
    }
  }

  return helper(num);
}
