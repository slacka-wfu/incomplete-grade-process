/**
 * THE GOOGLE APPS SCRIPT PARTNER - AUTOMATED INCOMPLETE REQUESTS
 * * ==========================================
 * CONFIGURATION: UPDATE THESE VARIABLES FIRST
 * ==========================================
 */
const CONFIG = {
  MASTER_DATA_SHEET_ID: '1boHlW-x_Qr0Z11qfU40WUDSgq7aRFjVnpyJS5KZIIm8', // ID of the Master Data Spreadsheet
  FORM_ID: '1SjsqrXuUej4zp7xMSQ5rQDSZtlGCDoVFL4Zekw31MXQ', // ID of the actual Google Form
  RESPONSES_SHEET_ID: '18VE3YpjqO3cII6wHCTV-JBFCRzNjqmDfMW8D9nxbMBg', // ID of the Spreadsheet holding the form responses
  RESPONSES_TAB_NAME: 'Form Responses 1', // Name of the tab holding the responses
  STATUS_COL_INDEX: 20, // Column U is the 21st column (1-indexed for getRange)
  KEY_PERSONNEL_EMAILS: 'spsreg@wfu.edu, spshelp@wfu.edu, stuserv@wfu.edu',
  
  // Update these exactly as they appear in your Google Form
  FORM_QUESTION_PROGRAM: 'Academic Program:',
  FORM_QUESTION_COURSE: 'Course Code and Number (e.g., FTA 714) for which the Incomplete is requested',

  // --- TEST MODE SETTINGS ---
  TEST_MODE: true, // Change to false when ready to go live
  TEST_EMAIL: 'slacka@wfu.edu' // All emails will route here while TEST_MODE is true
};

/**
 * ==========================================
 * SAFE EMAIL ROUTER (TEST MODE HANDLER)
 * ==========================================
 * Intercepts all emails. If TEST_MODE is true, routes them to the TEST_EMAIL.
 */
function sendRoutedEmail(originalTo, originalCc, subject, htmlBody) {
  let finalTo = originalTo;
  let finalCc = originalCc;
  let finalSubject = subject;

  if (CONFIG.TEST_MODE) {
    finalTo = CONFIG.TEST_EMAIL;
    finalCc = ''; // Clear CCs to prevent accidental test spam
    finalSubject = `[TEST MODE] ${subject}`;
    htmlBody = `<div style="background-color: #ffcccc; padding: 10px; margin-bottom: 20px; border: 1px solid red; color: black;">
                  <strong>TEST MODE ACTIVE</strong><br>
                  <strong>Original To:</strong> ${originalTo}<br>
                  <strong>Original CC:</strong> ${originalCc || 'None'}
                </div>` + htmlBody;
  }

  MailApp.sendEmail({
    to: finalTo,
    cc: finalCc,
    subject: finalSubject,
    htmlBody: htmlBody
  });
}

/**
 * ==========================================
 * PILLAR 1: FORM UPDATER (Nightly Trigger)
 * ==========================================
 */
function updateFormDropdowns() {
  const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_DATA_SHEET_ID);
  const form = FormApp.openById(CONFIG.FORM_ID);
  const items = form.getItems(FormApp.ItemType.LIST); 
  
  // 1. Get Academic Programs (Column E)
  const programSheet = masterSS.getSheetByName('Academic Programs');
  const programData = programSheet.getRange(2, 5, programSheet.getLastRow() - 1, 1).getValues();
  const programs = programData.map(row => row[0]).filter(String);
  
  // 2. Get Course Codes (Column A)
  const courseSheet = masterSS.getSheetByName('AD Lookup by Code');
  const courseData = courseSheet.getRange(2, 1, courseSheet.getLastRow() - 1, 1).getValues();
  const courses = courseData.map(row => row[0]).filter(String);
  
  // 3. Update the Form
  items.forEach(item => {
    const title = item.getTitle();
    if (title === CONFIG.FORM_QUESTION_PROGRAM) {
      item.asListItem().setChoiceValues(programs);
    } else if (title === CONFIG.FORM_QUESTION_COURSE) {
      item.asListItem().setChoiceValues(courses);
    }
  });
  
  Logger.log('Form dropdowns updated successfully.');
}

/**
 * ==========================================
 * PILLAR 2: THE INTAKE ROUTER (onFormSubmit Trigger)
 * ==========================================
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const responses = e.values;
  
  // Map Array Indices based on your CSV headers (0-indexed)
  const role = responses[2];
  const studentName = responses[3];
  const facultyEmail = responses[10];
  
  if (role.toLowerCase().includes('faculty')) {
    sheet.getRange(row, CONFIG.STATUS_COL_INDEX).setValue('Approved');
    sendFinalEmails(row, 'Approved', 'Approved via initial Faculty form submission.');
  } else {
    sheet.getRange(row, CONFIG.STATUS_COL_INDEX).setValue('Pending Faculty Approval');
    sendFacultyApprovalEmail(row, studentName, facultyEmail);
  }
}

/**
 * Sends the initial email to the faculty with the Web App links and IT help text.
 */
function sendFacultyApprovalEmail(row, studentName, facultyEmail) {
  const webAppUrl = ScriptApp.getService().getUrl(); 
  const approveUrl = `${webAppUrl}?row=${row}&action=Approve`;
  const denyUrl = `${webAppUrl}?row=${row}&action=Deny`;
  
  const subject = `ACTION REQUIRED: Incomplete Grade Request for ${studentName}`;
  const body = `
    <p>Dear Faculty Member,</p>
    <p>An Incomplete Grade Request has been submitted by <strong>${studentName}</strong> and requires your review.</p>
    <p>Please click one of the links below to securely approve or deny this request, and optionally add your notes:</p>
    <br>
    <p>
      <a href="${approveUrl}" style="background-color: #CFB53B; color: #000000; padding: 10px 15px; text-decoration: none; font-weight: bold; border-radius: 5px;">✅ Approve Request</a>
      &nbsp;&nbsp;&nbsp;
      <a href="${denyUrl}" style="background-color: #000000; color: #CFB53B; padding: 10px 15px; text-decoration: none; font-weight: bold; border-radius: 5px;">❌ Deny Request</a>
    </p>
    <br>
    <hr style="border: 0; border-top: 1px solid #cccccc; margin: 20px 0;">
    <div style="font-size: 12px; color: #555555; background-color: #f9f9f9; padding: 15px; border-radius: 5px;">
      <strong>🔒 WFU Security Notice & Troubleshooting:</strong><br>
      You must be logged into your <strong>@wfu.edu</strong> Google account to access this secure portal. If you click a link above and receive a "Google Drive" or "Unable to open" error, your browser is likely defaulting to your personal Gmail account.<br><br>
      <strong>How to fix:</strong> Right-click your chosen button above, select <em>"Open link in Incognito window"</em> (or Private window), and log in with your WFU credentials when prompted.
    </div>
    <p>Thank you,<br>School of Professional Studies</p>
  `;
  
  // Use the safe router instead of MailApp directly (respects TEST_MODE)
  sendRoutedEmail(facultyEmail, '', subject, body);
}

/**
 * ==========================================
 * PILLAR 3: THE WEB APP (Faculty Approver)
 * ==========================================
 */
function doGet(e) {
  const row = e.parameter.row;
  const action = e.parameter.action;
  
  const template = HtmlService.createTemplateFromFile('Index');
  template.row = row;
  template.action = action;
  
  return template.evaluate()
    .setTitle('SPS Incomplete Request')
    .setFaviconUrl('https://www.wfu.edu/wp-content/themes/wfu-theme/favicon.ico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processFacultyDecision(row, action, notes) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    
    // FIX: Force the row variable to be a strict Integer
    const rowIndex = parseInt(row, 10);
    
    // Safety check to ensure we actually got a number
    if (isNaN(rowIndex)) {
      throw new Error("Invalid row number received from the Web App URL.");
    }
    
    const status = action === 'Approve' ? 'Approved' : 'Denied';
    
    // Use the new rowIndex integer
    sheet.getRange(rowIndex, CONFIG.STATUS_COL_INDEX).setValue(status);
    
    // Pass the integer to the email function
    sendFinalEmails(rowIndex, status, notes);
    
    return `Success! The request was marked as ${status}. You may now close this tab.`;
  } catch (e) {
    Logger.log("Error processing decision: " + e.toString());
    return "Error: Could not process request. Please contact spshelp@wfu.edu.";
  } finally {
    lock.releaseLock();
  }
}

/**
 * ==========================================
 * HELPER FUNCTIONS
 * ==========================================
 */
function lookupSSMEmail(studentId, studentEmail) {
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_DATA_SHEET_ID);
  const assignSheet = ss.getSheetByName('SSMAssignments');
  const ssmSheet = ss.getSheetByName('SSMs');
  
  const assignData = assignSheet.getDataRange().getValues();
  let ssmName = null;
  
  for (let i = 1; i < assignData.length; i++) {
    const rowId = String(assignData[i][1]).trim();
    const rowEmail = String(assignData[i][7]).trim().toLowerCase();
    
    if (studentId && studentId.toUpperCase() !== 'N/A' && rowId === String(studentId).trim()) {
      ssmName = assignData[i][0];
      break;
    } else if (studentEmail && rowEmail === String(studentEmail).trim().toLowerCase()) {
      ssmName = assignData[i][0];
      break;
    }
  }
  
  if (!ssmName) return null;
  
  const ssmData = ssmSheet.getDataRange().getValues();
  for (let i = 1; i < ssmData.length; i++) {
    if (ssmData[i][0] === ssmName) {
      return ssmData[i][1];
    }
  }
  return null;
}

function sendFinalEmails(row, status, facultyNotes) {
  const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const studentName = data[3];
  const studentEmail = data[4];
  const studentId = data[5];
  const course = data[7];
  const facultyEmail = data[10];
  
  const ssmEmail = lookupSSMEmail(studentId, studentEmail);
  
  let ccEmails = `${CONFIG.KEY_PERSONNEL_EMAILS}, ${facultyEmail}`;
  if (ssmEmail) ccEmails += `, ${ssmEmail}`;
  
  const subject = `Incomplete Request ${status}: ${studentName} - ${course}`;
  const body = `
    <h2>Incomplete Request Update</h2>
    <p>The Incomplete Grade Request for <strong>${studentName}</strong> in course <strong>${course}</strong> has been <strong>${status}</strong>.</p>
    <p><strong>Faculty Notes/Comments:</strong><br>
    ${facultyNotes ? facultyNotes : 'None provided.'}</p>
    <hr>
    <p><em>This is an automated notification from the SPS Tracker.</em></p>
  `;
  
  // Use the safe router
  sendRoutedEmail(studentEmail, ccEmails, subject, body);
}

/**
 * ==========================================
 * PILLAR 4: THE 48-HOUR NUDGE (Daily Trigger)
 * ==========================================
 */
function sendNudgeEmails() {
  const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
  const data = sheet.getDataRange().getValues();
  
  const now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][0]);
    const studentName = data[i][3];
    const facultyEmail = data[i][10];
    const status = data[i][CONFIG.STATUS_COL_INDEX - 1]; 
    
    const hoursPassed = Math.abs(now - timestamp) / 36e5;
    
    if (status === 'Pending Faculty Approval' && hoursPassed > 48 && hoursPassed < 72) {
       const subject = `REMINDER: Pending Incomplete Request for ${studentName}`;
       const body = `<p>Hello,</p><p>This is an automated reminder that an incomplete request for ${studentName} has been awaiting your approval for over 48 hours. Please review your email for the approval link.</p><p>Thank you.</p>`;
       
       // Use the safe router
       sendRoutedEmail(facultyEmail, '', subject, body);
    }
  }
}