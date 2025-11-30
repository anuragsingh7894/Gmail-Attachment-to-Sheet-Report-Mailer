Step 1 – Prepare your Google Sheet
Create a new Google Sheet (or use your existing one).​

Add these sheets (tabs) with exact names:

Dump – this will be filled with CSV data from email.

Reports – place the final summary/report here; this is what goes in the email body.

Mail IDs – put recipient emails in:

Column A starting from A2: main recipients (To).

Column B starting from B2: CC recipients.

Make sure the company report email you receive actually has a CSV attachment and the subject matches:
Report results (Lucknow Daily Lead Data)

If the subject/sender are slightly different, update SEARCH_QUERY in the code to match your real Gmail search. You can test the query directly in Gmail first (search bar), then copy that same text into SEARCH_QUERY.​

Step 2 – Add the script to the Sheet
In the Sheet, go to Extensions → Apps Script.​

Delete any default myFunction code.

Paste your entire code (both importAttachmentToSheet, sendReportEmail, createHtmlTable, and exportSpreadsheetAsExcel if it is not already inside) into the editor.

At the top, set:

const SHEET_NAME = 'Dump'; (or your chosen data sheet name)

const TARGET_FOLDER_ID = '' or put a Google Drive folder ID if you want to save each CSV file there.​

Click the Save icon and give the project a meaningful name (e.g., “Lucknow Lead Automation”).

Step 3 – Run each function manually once (and authorize)
In the Apps Script editor, open the function dropdown and choose importAttachmentToSheet.​

Click Run.

On first run, Apps Script will ask for permissions (access Gmail, Drive, Sheets, and send email).

Review and allow; this links the script to your Google account.​

After it runs, go back to the Sheet and check the Dump tab:

It should now contain the CSV data from the latest matching email.

Next, from the function dropdown choose sendReportEmail and click Run.

Check your email inbox (for one of the addresses in Mail IDs) to confirm the report arrived with the HTML table and Excel attachment.​

If importAttachmentToSheet logs “No emails found matching the search query”, adjust SEARCH_QUERY until Gmail finds your report email.​

Step 4 – Automate with time-based triggers
To make this run automatically every day (for example):

In Apps Script, click the clock icon on the left (Triggers).​

Click Add Trigger (bottom right).

Choose function: importAttachmentToSheet

Event source: Time-driven

Type: Day timer

Choose an hour window (e.g., 7–8 AM) when the bank email is already received.

Add another trigger for sendReportEmail with a slightly later time (e.g., 8–9 AM), so data import happens first, then the report is sent.​

This way the process becomes:

Every morning, Script 1 pulls the latest CSV from Gmail into Dump.

Shortly after, Script 2 sends your “Lucknow HSPL – LEAD COUNT REPORT” with updated information automatically.

If you like, the next step can be: explain how to build the Reports sheet formulas so it summarizes the Dump data exactly as you need.