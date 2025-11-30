Automated Daily Lead Reporting System
Overview
This Google Apps Script automates the entire workflow of daily reporting. It fetches raw CSV data from a specific email attachment, processes it into a Google Sheet, converts the final report into an Excel file, and emails it to a distribution list.
Key Features:
* Auto-Fetch: Scans Gmail for specific subject lines and attachments.
* Auto-Import: Parses CSV attachments and updates a "Dump" sheet.
* Auto-Export: Converts the Google Sheet to .xlsx format.
* Auto-Email: Sends the report with a dynamic HTML table preview to stakeholders.
Prerequisites
* A Google account (Workspace or Gmail).
* A Google Sheet with the following tabs:
   1. Dump: Where raw data will be pasted.
   2. Reports: The formatted report (Pivot table or formulas) based on the "Dump" data.
   3. Mail IDs: Column A for "To" addresses, Column B for "CC" addresses.
Installation
1. Open your Google Sheet.
2. Go to Extensions > Apps Script.
3. Delete any existing code in the Code.gs file.
4. Copy and paste the code from this repository into the editor.
5. Update the Configuration Section at the top of the code (see below).
6. Save the project.
Configuration
Edit the constants at the top of the script to match your specific needs:
const CONFIG = {
 SEARCH_QUERY: 'has:attachment from:reports@example.com subject:"Daily Lead Data"',
 SHEET_NAME_DUMP: 'Dump',
 SHEET_NAME_REPORT: 'Reports',
 SHEET_NAME_MAILS: 'Mail IDs',
 EMAIL_SUBJECT_PREFIX: 'Daily Lead Count Report - ',
 SENDER_NAME: 'Your Name',
 SENDER_TITLE: 'LMS Coordinator'
};

How to Automate (Triggers)
To make this run automatically every day:
1. In the Apps Script editor, click on the Triggers icon (alarm clock) on the left sidebar.
2. Click + Add Trigger.
3. Trigger 1 (Import Data):
   * Function: importAttachmentToSheet
   * Event Source: Time-driven
   * Type: Day timer (e.g., 8am to 9am)
4. Trigger 2 (Send Report):
   * Function: sendReportEmail
   * Event Source: Time-driven
   * Type: Day timer (e.g., 9am to 10am)
   * Note: Ensure this runs AFTER the import trigger.
Permissions
When running for the first time, Google will ask for permission to access your:
* Gmail (to search and send)
* Drive (to save files)
* Spreadsheets (to write data)