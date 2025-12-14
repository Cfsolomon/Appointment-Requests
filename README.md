# Appointment-Requests — Google Apps Script

This repository contains a Google Apps Script that captures Google Form appointments, generates a Google Doc summary for each submission, exports and stores a PDF copy in Google Drive, sends an internal notification (with PDF) to clinic staff, and sends a personalized acknowledgement email to the submitter with useful patient resources. The project is easy to configure (FORM_ID, PDF_FOLDER_ID, EMAIL_ON_CREATE) and intended for small-to-medium clinical intake workflows.

## Files
- apps-script/Code.gs — main Apps Script file to paste into your spreadsheet's Apps Script project.

## Quick overview
- When the form is submitted, a Google Doc is created summarizing the responses.
- A PDF copy of the doc is saved to Drive (PDF_FOLDER_ID).
- Internal team (EMAIL_ON_CREATE) receives an email with the PDF attached.
- The submitter receives a personalized acknowledgement email (no PDF, no doc link).

## Setup & Configuration

1. Open your Google Sheet that stores the Form responses.
2. Extensions → Apps Script.
3. Create a new project (or open the existing one) and replace Code.gs with the file at `apps-script/Code.gs`.
4. Edit the top configuration constants:
   - `FORM_ID` — your Google Form ID.
   - `TEMPLATE_DOC_ID` — optional: a Google Doc ID to copy as template (or leave empty to create blank).
   - `DEST_FOLDER_ID` — optional: Drive folder ID to move created docs to.
   - `PDF_FOLDER_ID` — Drive folder ID where PDFs will be saved.
   - `EMAIL_ON_CREATE` — internal email to receive notifications and PDFs.
   - `SENDER_NAME`, clinic address, phone, fax — customize signature.
   - `DOB_FIELD_NAME` — exact form/sheet column title used for DOB extraction (if needed).
   - Link constants: `CUSTOM_LINK_URL`, `CUSTOM_LINK2_URL` and anchor texts.

5. Save the script.

## Authorization & Trigger

- In the Apps Script editor, run any function (e.g., `onFormSubmit`) once to trigger the OAuth consent and grant required scopes.
- Add an installable trigger:
  - Triggers → + Add Trigger
  - Function: onFormSubmit
  - Event source: From spreadsheet
  - Event type: On form submit
  - Save (authorize as necessary).

## Testing
- Submit the form (Preview) using an email you control.
- Confirm:
  - A Google Doc was created and a PDF saved in the `PDF_FOLDER_ID`.
  - Column X (DOC_URL_COLUMN) in the sheet contains the doc URL.
  - Internal email (EMAIL_ON_CREATE) arrives with the PDF attached.
  - Submitter receives the acknowledgement email with the two links and clinic signature.

Helper functions (paste into the Apps Script editor if you want them)
- `getOrCreatePdfFolder()` — create/find a folder by name and log its ID to use as `PDF_FOLDER_ID`.
- `testPdfFolderWrite()` — creates a small test file in the target folder to confirm permissions.

## Send-as alias (optional)
If you want messages to appear as actually sent from `info@azendosurg.com`:
- Add `info@azendosurg.com` as a verified "Send mail as" alias in the Gmail settings for the Google account running the script (or configure via Workspace admin).
- After verification, you can switch to `GmailApp.sendEmail(..., {from: 'info@azendosurg.com'})` if needed.

## Notes & Permissions
- The account running the Apps Script must have Editor access to the target Drive folder and access to the Form and Sheet.
- Be mindful of Apps Script quotas for daily emails, execution time, and Drive writes.

## License
MIT — see LICENSE
