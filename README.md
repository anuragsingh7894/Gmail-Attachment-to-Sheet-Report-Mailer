# Gmail-Attachment-to-Sheet-Report-Mailer

Code:

function importAttachmentToSheet() {
// Define search criteria for the email
const SEARCH_QUERY = 'has:attachment from:noreplysf@hdfcbank.com subject:"Report results (Lucknow Daily Lead Data)"'; // Adjust as needed
const SHEET_NAME = 'Dump'; // Name of the sheet to import data into
const TARGET_FOLDER_ID = ''; // Optional: ID of a Google Drive folder to save attachments

const threads = GmailApp.search(SEARCH_QUERY);

if (threads.length === 0) {
Logger.log('No emails found matching the search query.');
return;
}

// Process the most recent email found
const latestMessage = threads[0].getMessages().pop(); // Gets the last message in the thread
const attachments = latestMessage.getAttachments();

if (attachments.length === 0) {
Logger.log('No attachments found in the latest matching email.');
return;
}

for (const attachment of attachments) {
// Filter for specific attachment types if needed (e.g., CSV, XLSX)
if (attachment.getContentType() === MimeType.CSV || attachment.getName().endsWith('.csv')) {
const fileBlob = attachment.copyBlob();

// Optional: Save the attachment to Google Drive
if (TARGET_FOLDER_ID) {
const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
folder.createFile(fileBlob);
Logger.log(`Attachment "${attachment.getName()}" saved to Google Drive.`);
}

// Read CSV data and import to Google Sheet
const csvData = Utilities.parseCsv(fileBlob.getDataAsString());
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let sheet = spreadsheet.getSheetByName(SHEET_NAME);

if (!sheet) {
sheet = spreadsheet.insertSheet(SHEET_NAME);
Logger.log(`Created new sheet: ${SHEET_NAME}`);
} else {
sheet.clearContents(); // Clear existing data before importing
}

// Write data to the sheet
if (csvData.length > 0) {
sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
Logger.log(`Data from "${attachment.getName()}" imported to sheet "${SHEET_NAME}".`);
} else {
Logger.log(`CSV file "${attachment.getName()}" is empty.`);
}
}
}
}


//** below function sends mail with attachment triggred after uploading the data */


function sendReportEmail() {
// Get active spreadsheet and sheets
const ss = SpreadsheetApp.getActiveSpreadsheet();
const reportSheet = ss.getSheetByName("Reports");
const mailSheet = ss.getSheetByName("Mail IDs");
var today = new Date();
var timeZone = Session.getScriptTimeZone();
var formattedDate = Utilities.formatDate(today, timeZone, "dd MMMM yyyy")

const yesterday = new Date(today); // Create a new Date object based on today
yesterday.setDate(today.getDate() - 1);
var yest = Utilities.formatDate(yesterday, timeZone, "dd MMMM yyyy")

// Get all email IDs from column 1
const emails = mailSheet.getRange("A2:A").getValues()
.flat()
.filter(e => e && e.includes("@"));

const dd = mailSheet.getRange("B2:B").getValues()
.flat()
.filter(f => f && f.includes("@"));

// Convert report sheet to HTML table
const range = reportSheet.getDataRange();
const values = range.getDisplayValues();
const htmlTable = createHtmlTable(values);


function exportSpreadsheetAsExcel(spreadsheet) {
const spreadsheetId = spreadsheet.getId();
const url = `https://docs.google.com/spreadsheets/export?id=${spreadsheetId}&exportFormat=xlsx`;

const params = {
method: "get",
headers: {
Authorization: "Bearer " + ScriptApp.getOAuthToken(),
},
muteHttpExceptions: true,
};

const blob = UrlFetchApp.fetch(url, params).getBlob();
return blob.setName(spreadsheet.getName() + ".xlsx");
}

// --- Export Spreadsheet as Excel (.xlsx) ---
const excelBlob = exportSpreadsheetAsExcel(ss);

// Email subject and body
const subject = "Lucknow HSPL - LEAD COUNT REPORT - " + yest;
const body = `
<p>Dear All,</p>
<p>Please find the Lead Count Report on LMS calling Yes vs NO and BSA and NON BSA leads wise till yesterday.:</p>;
${htmlTable}
<p>Kindly consider the attached file for detailed report and data file.<br> <br>Regards~<br>Anurag<br>LMS Coordinator<br>HDFC SALES</p>
`;

// Send email
MailApp.sendEmail({
to: emails.join(","),
cc: dd.join(","),
subject: subject,
htmlBody: body,
attachments: [{
fileName: "LMS Data Lucknow " + yest + ".xlsx",
content: excelBlob.getBytes(),
mimeType: MimeType.MICROSOFT_EXCEL
}],
});

Logger.log("Email sent to: " + emails.join(", "));
Logger.log("Email sent to: " + dd.join(", "));
}

// Helper function: convert array to HTML table
function createHtmlTable(data) {
let html = '<table border="1" style="border-collapse:collapse;">';
data.forEach((row, i) => {
html += '<tr>';
row.forEach(cell => {
html += i === 0
? `<th style="background:#f0f0f0;padding:6px;">${cell}</th>`
: `<td style="padding:6px;">${cell}</td>`;
});
html += '</tr>';
});
html += '</table>';
return html;
}
