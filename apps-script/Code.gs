/**
 * Form -> Sheet -> Google Doc helper (safe, ready to paste)
 *
 * Instructions:
 * 1) Open your spreadsheet -> Extensions -> Apps Script.
 * 2) Replace the contents of Code.gs with this file and Save.
 * 3) Authorize and test with a form submission.
 *
 * NOTE: Set DEST_FOLDER_ID / PDF_FOLDER_ID to Drive folder IDs if you want files moved/saved there.
 */

// CONFIG - update these values as needed
const FORM_ID = '12s8Xah3w5RHRNZijGuNBka8Z5PGi3X_x_wk8LD4tbOo'; // your Form ID
const TEMPLATE_DOC_ID = '1ZNS8AGnD6S2AgAZJnnVmW92EPnz_hMWPybJ0IzRKqBY'; // optional template doc ID or '' to create blank docs
const DEST_FOLDER_ID = ''; // Drive folder ID (NOT a local path). Leave empty to keep doc in My Drive.
const EMAIL_ON_CREATE = 'info@azendosurg.com'; // internal notification email (will receive the PDF)
const MATCH_TOLERANCE_MS = 60000; // timestamp match tolerance in ms (increase if needed)

// SENDER + PDF settings
const SENDER_NAME = 'Richard J. Harding, MD, FACS';   // display name shown to recipients
const SAVE_PDF = true;                // save a PDF copy of the created Google Doc to Drive
const ATTACH_PDF_TO_EMAIL = false;    // attachments to submitter? keep false per your request
const PDF_FOLDER_ID = '1HeOOcPMO0HFnH3CUgErhZlS1sAqlF-Vi'; // your PDF folder ID

// CLINIC CONTACT (update to your actual clinic address)
const CLINIC_ADDRESS_PLAIN = '2320 N. Third Street\nPhoenix, AZ 85004-1303';
const CLINIC_ADDRESS_HTML  = '2320 N. Third Street<br>Phoenix, AZ 85004-1303';
const CLINIC_PHONE_PLAIN = 'Call: (602) 340-0201';
const CLINIC_PHONE_HTML  = 'Call: (602) 340-0201';
const CLINIC_FAX_PLAIN = 'FAX: (602) 889-2926';
const CLINIC_FAX_HTML  = 'FAX: (602) 889-2926';

// CUSTOMIZABLE FIELDS
const DOB_FIELD_NAME = 'Birthday (Month/Day/Year)'; // exact header/question title for DOB (Column K)
const DOC_URL_COLUMN = 24; // column number to write doc URL (Column X)

// NEW: Email link/content options (two links provided by you)
// Note: we will NOT include the Google Doc link in the submitter acknowledgement per your request
const INCLUDE_CUSTOM_LINK = true;
const CUSTOM_LINK_URL = 'https://www.thyroidnoduletreatment.center/participant-page/thyroidnodulelearningcenter';
const CUSTOM_LINK_TEXT = 'Thyroid Nodule Learning Center';

const INCLUDE_CUSTOM_LINK_2 = true;
const CUSTOM_LINK2_URL = 'https://www.thyroidnoduletreatment.center/thyroid-library';
const CUSTOM_LINK2_TEXT = 'Thyroid Library';

const INCLUDE_DOC_LINK = false;              // DO NOT include a link to the created Google Doc in the ack email to the submitter
const SHARE_DOC_WITH_RESPONDENT = false;     // keep false (we won't share the doc with respondent)

/* ---------- Main trigger handler ---------- */
function onFormSubmit(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var lastRow = sheet.getLastRow();
    var timestampCell = sheet.getRange(lastRow, 1).getValue();
    var rowTimestamp = timestampCell instanceof Date ? timestampCell : new Date(timestampCell);

    var matchedResponse = null;
    if (FORM_ID && FORM_ID.trim()) {
      try {
        var form = FormApp.openById(FORM_ID);
        var responses = form.getResponses();
        for (var i = responses.length - 1; i >= 0; i--) {
          var r = responses[i];
          if (Math.abs(r.getTimestamp().getTime() - rowTimestamp.getTime()) <= MATCH_TOLERANCE_MS) {
            matchedResponse = r;
            break;
          }
        }
      } catch (err) {
        Logger.log('Could not open Form by ID or insufficient permission: ' + err);
        matchedResponse = null;
      }
    } else {
      Logger.log('FORM_ID is empty — falling back to spreadsheet values only.');
    }

    var newDocFile = createDocFromResponse(matchedResponse, e && e.namedValues ? e.namedValues : {}, rowTimestamp);

    // Save the created document URL back into the spreadsheet (fixed column DOC_URL_COLUMN)
    var docUrl = '';
    try {
      docUrl = 'https://docs.google.com/document/d/' + newDocFile.getId() + '/edit';
      sheet.getRange(lastRow, DOC_URL_COLUMN).setValue(docUrl);
    } catch (err) {
      Logger.log('Could not write doc URL to sheet (column ' + DOC_URL_COLUMN + '): ' + err);
    }

    // Move doc to destination folder if configured (DEST_FOLDER_ID must be a Drive folder ID)
    if (DEST_FOLDER_ID && DEST_FOLDER_ID.trim()) {
      try {
        var dest = DriveApp.getFolderById(DEST_FOLDER_ID);
        dest.addFile(newDocFile);
        var parents = newDocFile.getParents();
        while (parents.hasNext()) {
          var p = parents.next();
          p.removeFile(newDocFile);
        }
      } catch (err) {
        Logger.log('Could not move file to folder (check DEST_FOLDER_ID and permissions): ' + err);
      }
    }

    // Generate PDF blob (try to get from doc). Use this blob for saving and attaching.
    var pdfBlob = null;
    try {
      pdfBlob = DriveApp.getFileById(newDocFile.getId()).getAs('application/pdf');
    } catch (errPdfBlob) {
      Logger.log('Could not generate PDF blob from doc (will try other methods): ' + errPdfBlob);
      pdfBlob = null;
    }

    // Save PDF copy (optional) - this will always save for internal records if SAVE_PDF = true
    var savedPdfFile = null;
    if (SAVE_PDF) {
      try {
        var pdfName = newDocFile.getName() + '.pdf';
        if (pdfBlob) {
          if (PDF_FOLDER_ID && PDF_FOLDER_ID.trim()) {
            var pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
            savedPdfFile = pdfFolder.createFile(pdfBlob).setName(pdfName);
          } else {
            var parentsIter = newDocFile.getParents();
            if (parentsIter.hasNext()) {
              var parent = parentsIter.next();
              savedPdfFile = parent.createFile(pdfBlob).setName(pdfName);
            } else {
              savedPdfFile = DriveApp.createFile(pdfBlob).setName(pdfName);
            }
          }
        } else {
          // If pdfBlob wasn't available, attempt to create by opening the Document and getting content
          try {
            var docFile = DriveApp.getFileById(newDocFile.getId());
            var altBlob = docFile.getAs('application/pdf');
            if (PDF_FOLDER_ID && PDF_FOLDER_ID.trim()) {
              var pdfFolder2 = DriveApp.getFolderById(PDF_FOLDER_ID);
              savedPdfFile = pdfFolder2.createFile(altBlob).setName(pdfName);
            } else {
              var parentsIter2 = newDocFile.getParents();
              if (parentsIter2.hasNext()) {
                var parent2 = parentsIter2.next();
                savedPdfFile = parent2.createFile(altBlob).setName(pdfName);
              } else {
                savedPdfFile = DriveApp.createFile(altBlob).setName(pdfName);
              }
              // also set pdfBlob so we can attach below
              pdfBlob = altBlob;
            }
          } catch (errAlt) {
            Logger.log('Could not save PDF by alternate method: ' + errAlt);
            savedPdfFile = null;
          }
        }
        Logger.log('Saved PDF: ' + (savedPdfFile ? savedPdfFile.getUrl() : 'none'));
      } catch (errPdf) {
        Logger.log('Could not generate/save PDF for doc: ' + errPdf);
        savedPdfFile = null;
      }
    }

    // Optional email notification to internal address (with display name)
    // Attach the PDF only to the internal notification (not to the submitter)
    if (EMAIL_ON_CREATE && EMAIL_ON_CREATE.trim()) {
      try {
        var internalAttachments = [];
        // Prefer the pdfBlob (already available if we generated it); otherwise try savedPdfFile
        if (pdfBlob) {
          internalAttachments.push(pdfBlob);
        } else if (savedPdfFile) {
          try {
            internalAttachments.push(savedPdfFile.getAs('application/pdf'));
          } catch (errConv) {
            Logger.log('Could not get blob from savedPdfFile for attachment: ' + errConv);
          }
        }

        var internalMail = {
          to: EMAIL_ON_CREATE,
          subject: 'New Form Response Document: ' + newDocFile.getName(),
          body: 'A new document was created for a form submission:\n\n' + docUrl,
          name: SENDER_NAME,
          replyTo: EMAIL_ON_CREATE
        };
        if (internalAttachments.length > 0) internalMail.attachments = internalAttachments;

        MailApp.sendEmail(internalMail);
      } catch (err) {
        Logger.log('Could not send internal email notification: ' + err);
      }
    }

    // --- Simple personalised acknowledgement email to the submitter (no appointment details, no PDF attachments) ---
    try {
      var respondent = '';
      // Prefer the helper if present
      if (typeof extractRespondentEmail === 'function') {
        respondent = extractRespondentEmail(matchedResponse, e && e.namedValues ? e.namedValues : {});
      }

      // Fallback: search namedValues for any key containing "email"
      if ((!respondent || !String(respondent).trim()) && e && e.namedValues) {
        for (var k in e.namedValues) {
          if (!e.namedValues.hasOwnProperty(k)) continue;
          try {
            if (k.toLowerCase().indexOf('email') !== -1) {
              var vv = e.namedValues[k];
              respondent = Array.isArray(vv) ? vv[0] : vv;
              if (respondent) break;
            }
          } catch (inner) { /* ignore */ }
        }
      }

      // extract first name for personalization
      var firstName = extractFirstName(matchedResponse, e && e.namedValues ? e.namedValues : {});
      var greeting = firstName ? 'Dear ' + firstName + ',' : 'Hi,';

      if (respondent && String(respondent).trim()) {
        respondent = String(respondent).trim();

        var subject = 'We received your appointment request';

        // Build links (two custom links only)
        var linksPlain = '';
        var linksHtml = '';

        if (INCLUDE_CUSTOM_LINK && CUSTOM_LINK_URL && CUSTOM_LINK_TEXT) {
          linksPlain += CUSTOM_LINK_TEXT + ': ' + CUSTOM_LINK_URL + '\n';
          linksHtml  += '<p><a href="' + CUSTOM_LINK_URL + '">' + CUSTOM_LINK_TEXT + '</a></p>';
        }

        if (INCLUDE_CUSTOM_LINK_2 && CUSTOM_LINK2_URL && CUSTOM_LINK2_TEXT) {
          linksPlain += CUSTOM_LINK2_TEXT + ': ' + CUSTOM_LINK2_URL + '\n';
          linksHtml  += '<p><a href="' + CUSTOM_LINK2_URL + '">' + CUSTOM_LINK2_TEXT + '</a></p>';
        }

        // Final email order requested:
        // 1) Greeting ("Dear FirstName,")
        // 2) Thank you + response timing
        // 3) Links
        // 4) Signature (address + phone + fax)
        var phoneFaxPlain = CLINIC_PHONE_PLAIN + ' | ' + CLINIC_FAX_PLAIN;
        var phoneFaxHtml  = CLINIC_PHONE_HTML + ' | ' + CLINIC_FAX_HTML;

        var plainBody = greeting + '\n\n'
                      + 'Thank you — we have received your appointment request form submission. Our medical staff will review your information and follow up by email and/or phone if further details are needed.\n\n'
                      + 'We typically respond within 2 business days. In the interim, if you need to contact us, please send email to ' + EMAIL_ON_CREATE + '.\n\n'
                      + linksPlain + '\n'
                      + 'Thank you,\n' + SENDER_NAME + '\n' + CLINIC_ADDRESS_PLAIN + '\n' + phoneFaxPlain;

        var htmlBody = '<p>' + (firstName ? ('Dear ' + firstName + ',') : 'Hi,') + '</p>'
                     + '<p>Thank you — we have received your appointment request form submission. Our medical staff will review your information and follow up by email and/or phone if further details are needed.</p>'
                     + '<p>We typically respond within 2 business days. If you need to contact us sooner, email <a href="mailto:' + EMAIL_ON_CREATE + '">' + EMAIL_ON_CREATE + '</a>.</p>'
                     + linksHtml
                     + '<p>Thank you,<br>' + SENDER_NAME + '<br>' + CLINIC_ADDRESS_HTML + '<br>' + phoneFaxHtml + '</p>';

        // Ensure submitter receives NO PDF attachments
        MailApp.sendEmail({
          to: respondent,
          subject: subject,
          body: plainBody,
          htmlBody: htmlBody,
          name: SENDER_NAME,
          replyTo: EMAIL_ON_CREATE
        });

        Logger.log('Acknowledgement email sent to: %s', respondent);
      } else {
        Logger.log('No respondent email found — acknowledgement not sent.');
      }
    } catch (errAck) {
      Logger.log('Error sending acknowledgement email: ' + errAck);
    }

  } catch (err) {
    Logger.log('onFormSubmit top-level error: ' + err);
  }
}

/* ---------- Respondent email extraction helper ---------- */
function extractRespondentEmail(formResponse, namedValues) {
  // 1) Try FormResponse.getRespondentEmail() (works when "collect email addresses" is ON)
  try {
    if (formResponse && typeof formResponse.getRespondentEmail === 'function') {
      var fe = formResponse.getRespondentEmail();
      if (fe && String(fe).trim()) return String(fe).trim();
    }
  } catch (err) {
    // ignore and fall back to namedValues
  }

  // 2) Fallback: prefer the manual field titled "Preferred contact email" if present
  try {
    var preferredKey = 'Preferred contact email'; // adjust if your question title differs
    if (namedValues && preferredKey in namedValues) {
      var pv = namedValues[preferredKey];
      if (Array.isArray(pv)) pv = pv[0];
      if (pv && String(pv).trim()) return String(pv).trim();
    }
  } catch (err) { /* ignore */ }

  // 3) Fallback: search namedValues keys for anything containing "email" (case-insensitive)
  if (namedValues) {
    for (var k in namedValues) {
      if (!namedValues.hasOwnProperty(k)) continue;
      try {
        if (k.toLowerCase().indexOf('email') !== -1) {
          var v = namedValues[k];
          if (Array.isArray(v)) v = v[0];
          if (v && String(v).trim()) return String(v).trim();
        }
      } catch (inner) { /* continue */ }
    }
  }
  return '';
}

/* ---------- Extract first name helper ---------- */
function extractFirstName(formResponse, namedValues) {
  // 1) Try namedValues first
  try {
    if (namedValues && namedValues['First Name']) {
      var fn = Array.isArray(namedValues['First Name']) ? namedValues['First Name'][0] : namedValues['First Name'];
      if (fn && String(fn).trim()) return String(fn).trim();
    }
  } catch (e) { /* ignore */ }

  // 2) Try common title matches in formResponse
  try {
    if (formResponse && typeof formResponse.getItemResponses === 'function') {
      var irs = formResponse.getItemResponses();
      for (var i = 0; i < irs.length; i++) {
        try {
          var ir = irs[i];
          var title = ir.getItem() && ir.getItem().getTitle ? ir.getItem().getTitle() : '';
          var tl = title ? title.toLowerCase() : '';
          if (tl.indexOf('first name') !== -1 || tl === 'first' || tl.indexOf('given name') !== -1) {
            var r = ir.getResponse();
            if (r && String(r).trim()) return String(r).trim();
          }
        } catch (inner) { /* ignore per-item errors */ }
      }
    }
  } catch (e) { /* ignore */ }

  return '';
}

/* ---------- Helper: normalize DOB value to Month/Day/Year (MM/dd/yyyy) when possible ---------- */
function formatDateOfBirth(rawValue, tz) {
  if (!rawValue && rawValue !== 0) return '';
  try {
    if (rawValue instanceof Date) {
      // return in MM/dd/yyyy format
      return Utilities.formatDate(rawValue, tz, 'MM/dd/yyyy');
    }
  } catch (e) {
    // continue
  }
  // If it's a string that parses to a valid date, use that
  try {
    var parsed = new Date(rawValue);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, tz, 'MM/dd/yyyy');
    }
  } catch (e) {
    // continue
  }
  // fallback: return trimmed raw string (shortened)
  var s = String(rawValue).trim();
  if (s.length > 30) s = s.substring(0, 30);
  return s;
}

/* ---------- Document creation ---------- */
function createDocFromResponse(formResponse, namedValues, timestamp) {
  // Build document name using First Name + Last Name + Date of Birth (fallbacks + sanitization)
  var tz = Session.getScriptTimeZone();
  var ts = timestamp instanceof Date ? timestamp : new Date();

  // Try to get first/last from namedValues (spreadsheet column titles)
  var firstName = null;
  var lastName = null;

  if (namedValues) {
    if (namedValues['First Name']) {
      firstName = Array.isArray(namedValues['First Name']) ? namedValues['First Name'][0] : namedValues['First Name'];
    }
    if (namedValues['Last Name']) {
      lastName = Array.isArray(namedValues['Last Name']) ? namedValues['Last Name'][0] : namedValues['Last Name'];
    }
  }

  // If not found in namedValues, try to read from the FormResponse itemResponses
  if ((!firstName || !lastName) && formResponse) {
    try {
      var irs = formResponse.getItemResponses();
      for (var j = 0; j < irs.length; j++) {
        var ir = irs[j];
        try {
          var title = ir.getItem() && ir.getItem().getTitle ? ir.getItem().getTitle() : '';
          if (!firstName && title === 'First Name') {
            firstName = ir.getResponse();
          }
          if (!lastName && title === 'Last Name') {
            lastName = ir.getResponse();
          }
        } catch (inner) {
          // ignore individual item errors
        }
        if (firstName && lastName) break;
      }
    } catch (err) {
      // silent fallback to other sources
    }
  }

  // Normalize/fallback values
  firstName = firstName ? String(firstName).trim() : '';
  lastName = lastName ? String(lastName).trim() : '';

  var baseName = '';
  if (firstName && lastName) {
    baseName = firstName + ' ' + lastName;
  } else if (firstName) {
    baseName = firstName;
  } else if (lastName) {
    baseName = lastName;
  } else {
    baseName = 'Form Response';
  }

  // sanitize baseName: remove characters not allowed in Drive filenames and limit length
  baseName = baseName.replace(/[\/\\:\*\?"<>\|]/g, '').substring(0, 60);

  // --- Attempt to read Date of Birth and use it in the doc title ---
  var dobRaw = null;
  // Preferred namedValue key (exact header)
  if (namedValues && namedValues[DOB_FIELD_NAME]) {
    dobRaw = Array.isArray(namedValues[DOB_FIELD_NAME]) ? namedValues[DOB_FIELD_NAME][0] : namedValues[DOB_FIELD_NAME];
  }

  // If not found in namedValues, try to find in the FormResponse itemResponses by checking common DOB keys
  if ((!dobRaw || String(dobRaw).trim() === '') && formResponse) {
    try {
      var itemResponses = formResponse.getItemResponses();
      var dobFieldLower = DOB_FIELD_NAME ? DOB_FIELD_NAME.toLowerCase() : '';
      for (var k = 0; k < itemResponses.length; k++) {
        try {
          var ir = itemResponses[k];
          var title = ir.getItem() && ir.getItem().getTitle ? ir.getItem().getTitle() : '';
          var tl = title ? title.toLowerCase() : '';
          if (tl === dobFieldLower || tl.indexOf('date of birth') !== -1 || tl.indexOf('dob') !== -1 || tl.indexOf('birth') !== -1) {
            dobRaw = ir.getResponse();
            break;
          }
        } catch (inner) {
          /* ignore */
        }
      }
    } catch (err) {
      /* ignore */
    }
  }

  var dobString = formatDateOfBirth(dobRaw, tz); // returns MM/dd/yyyy when possible

  // final doc name: use DOB if available, otherwise fall back to timestamp
  var docName;
  if (dobString && dobString !== '') {
    // sanitize dobString for filename: replace '/' with '-' (slashes not allowed in Drive filenames)
    var safeDob = dobString.replace(/\//g, '-').replace(/[\\:\*\?"<>\|]/g, '');
    docName = baseName + ' - ' + safeDob;
  } else {
    // fallback to timestamp for uniqueness
    docName = baseName + ' - ' + Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm:ss');
  }

  var newDocFile;
  try {
    if (TEMPLATE_DOC_ID && TEMPLATE_DOC_ID.trim()) {
      newDocFile = DriveApp.getFileById(TEMPLATE_DOC_ID).makeCopy(docName);
    } else {
      var doc = DocumentApp.create(docName);
      newDocFile = DriveApp.getFileById(doc.getId());
    }
  } catch (err) {
    throw new Error('Could not create or copy template doc: ' + err);
  }

  var docId = newDocFile.getId();
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  body.appendParagraph('Form Response').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Submitted: ' + Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm:ss'));
  body.appendHorizontalRule();

  if (formResponse) {
    var itemResponses = formResponse.getItemResponses();
    for (var i = 0; i < itemResponses.length; i++) {
      var ir = itemResponses[i];
      var qTitle = '(question title unavailable)';
      var itemType = null;
      var resp = null;

      try {
        var item = ir.getItem();
        qTitle = item ? (item.getTitle ? item.getTitle() : qTitle) : qTitle;
        try { itemType = item ? (item.getType ? item.getType() : null) : null; } catch (inner) { itemType = null; }
      } catch (err) {
        Logger.log('Could not get Item for an ItemResponse at index ' + i + ': ' + err);
      }

      try { resp = ir.getResponse(); } catch (err) { Logger.log('Could not get response for item "' + qTitle + '": ' + err); resp = null; }

      body.appendParagraph(qTitle).setHeading(DocumentApp.ParagraphHeading.HEADING3);

      if (itemType == FormApp.ItemType.FILE_UPLOAD) {
        var fileIds = Array.isArray(resp) ? resp : (resp ? [resp] : []);
        if (fileIds.length === 0) {
          body.appendParagraph('(no files uploaded)');
        } else {
          fileIds.forEach(function(raw) {
            try {
              var fid = extractDriveFileId(raw);
              var file = DriveApp.getFileById(fid);
              var blob = file.getBlob();
              if (blob && blob.getContentType().indexOf('image/') === 0) {
                body.appendImage(blob).setAltDescription(file.getName());
                body.appendParagraph('File: ' + file.getName() + ' (image)');
              } else {
                body.appendParagraph('File: ' + file.getName() + ' — ' + file.getUrl());
              }
            } catch (err) {
              body.appendParagraph('File upload (could not fetch): ' + String(raw));
              Logger.log('Error retrieving uploaded file for question "' + qTitle + '": ' + err);
            }
          });
        }
        body.appendHorizontalRule();
        continue;
      }

      try {
        if (Array.isArray(resp)) {
          body.appendParagraph(resp.join(', '));
        } else if (resp === null || resp === undefined || resp === '') {
          body.appendParagraph('(no answer)');
        } else {
          body.appendParagraph(String(resp));
        }
      } catch (err) {
        Logger.log('Error writing response for question "' + qTitle + '": ' + err);
        body.appendParagraph('(could not write response)');
      }
      body.appendHorizontalRule();
    }
  } else {
    body.appendParagraph('(FormResponse not found — using spreadsheet values)').setItalic(true);
    for (var key in namedValues) {
      if (!namedValues.hasOwnProperty(key)) continue;
      var val = namedValues[key];
      body.appendParagraph(key).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      if (Array.isArray(val)) {
        body.appendParagraph(val.join(', '));
      } else {
        body.appendParagraph(String(val));
      }
      body.appendHorizontalRule();
    }
  }

  doc.saveAndClose();
  return newDocFile;
}

function extractDriveFileId(value) {
  if (!value) return value;
  if (typeof value !== 'string') return value;
  var idPattern = /^[a-zA-Z0-9\-_]{20,}$/;
  if (idPattern.test(value)) return value;
  var m = value.match(/[-\w]{25,}/);
  if (m) return m[0];
  return value;
}
/* ----- testFormAccess (commented out) -----
function testFormAccess() {
  try {
    var f = FormApp.openById('12s8Xah3w5RHRNZijGuNBka8Z5PGi3X_x_wk8LD4tbOo');
  } catch (err) {
    Logger.log('Form open error: ' + err);
  }
}
----- end testFormAccess (commented out) ----- */
