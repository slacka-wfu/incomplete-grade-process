/**
 * THE GOOGLE APPS SCRIPT PARTNER - AUTOMATED INCOMPLETE REQUESTS
 * ===============================================================
 * Enterprise Solutions Architecture: Wake Forest University (SPS)
 * @author Amy Slack & Solutions Architect
 * @version 5.0.0 - Production Master Build & Admin Triage Tooling
 */

const CONFIG = {
  // --- CORE INFRASTRUCTURE IDS ---
  MASTER_DATA_SHEET_ID: '1boHlW-x_Qr0Z11qfU40WUDSgq7aRFjVnpyJS5KZIIm8', 
  FORM_ID: '1SjsqrXuUej4zp7xMSQ5rQDSZtlGCDoVFL4Zekw31MXQ', 
  RESPONSES_SHEET_ID: '18VE3YpjqO3cII6wHCTV-JBFCRzNjqmDfMW8D9nxbMBg', 
  RESPONSES_TAB_NAME: 'Form Responses 1', 
  AUDIT_TAB_NAME: 'System_Audit_Log', 

  // --- PROGRAMMATIC LEDGER & PARITY TARGETS (1-Indexed for getRange) ---
  STATUS_COL_INDEX: 28, // Column AB: Programmatic State Manager
  NOTES_COL_INDEX: 29,  // Column AC: Denial Reasons & System Notes

  // --- STAKEHOLDER ROUTING EMAILS ---
  KEY_PERSONNEL_EMAILS: 'spsreg@wfu.edu, spshelp@wfu.edu, stuserv@wfu.edu',
  REGISTRAR_FAILSAFE_EMAIL: 'traverk@wfu.edu', 
  WATERMARK_COMPLIANCE_EMAIL: 'maguirl@wfu.edu', 

  // --- FORM FIELD STRINGS (For Nightly Sync & Dynamic Lookups) ---
  FORM_QUESTION_PROGRAM: 'Academic Program',
  FORM_QUESTION_COURSE: 'Course Code and Number',

  // --- ENVIRONMENT GOVERNANCE ---
  TEST_MODE: true, // SET TO FALSE FOR LIVE PRODUCTION
  TEST_EMAIL: 'slacka@wfu.edu' 
};

/**
 * ===============================================================
 * DYNAMIC HEADER COALESCING ENGINE (Column-Shift Immunity)
 * ===============================================================
 * Scans Row 1 headers for a partial string match and returns the first non-empty row value.
 */
function getCoalescedVal(headers, rowValues, searchString) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase().includes(searchString.toLowerCase())) {
      if (rowValues[i] && String(rowValues[i]).trim() !== '') {
        return rowValues[i];
      }
    }
  }
  return '';
}

/**
 * ===============================================================
 * PILLAR 6: ADMINISTRATIVE UI & AUDIT PANEL (MASTER MENU)
 * ===============================================================
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ SPS Admin Tools')
    .addItem('✏️ Override Student Status (Active Row)', 'openStatusOverrideModal')
    .addSeparator()
    .addItem('📊 Launch System Audit Panel', 'launchAuditPanel')
    .addItem('🔄 Force Sync Form Dropdowns', 'updateFormDropdowns')
    .addItem('📧 Retrigger Faculty Approval Email', 'uiRetriggerEmail')
    .addToUi();
}

function launchAuditPanel() {
  const html = HtmlService.createTemplateFromFile('AuditPanel').evaluate()
      .setWidth(850).setHeight(550).setTitle('SPS System Audit History');
  SpreadsheetApp.getUi().showModalDialog(html, 'System Audit & Event Tracer');
}

function getRecentAuditLogs() {
  const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
  const auditSheet = ss.getSheetByName(CONFIG.AUDIT_TAB_NAME);
  if (!auditSheet) return [];
  const data = auditSheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  return data.slice(1).reverse().slice(0, 50);
}

function logSystemAudit(eventType, details, primaryKey = 'System') {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    let auditSheet = ss.getSheetByName(CONFIG.AUDIT_TAB_NAME);
    if (!auditSheet) {
      auditSheet = ss.insertSheet(CONFIG.AUDIT_TAB_NAME);
      auditSheet.appendRow(['Timestamp', 'Event Type', 'Primary Key (ID / Row)', 'Transaction Details']);
      auditSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#000000').setFontColor('#FDC314');
      auditSheet.setFrozenRows(1);
      auditSheet.setColumnWidth(1, 150); auditSheet.setColumnWidth(2, 120); auditSheet.setColumnWidth(4, 500);
    }
    auditSheet.appendRow([new Date(), eventType, primaryKey, details]);
    console.info(`[${eventType}] ${details}`); 
  } catch (e) { console.error(`Audit Engine Failure: ${e.message}`); }
}

/**
 * ===============================================================
 * KARA TRAVERSE ACTIVE-ROW OVERRIDE ENGINE (PATTERN 1)
 * ===============================================================
 */
function openStatusOverrideModal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.RESPONSES_TAB_NAME);
  const activeRange = SpreadsheetApp.getActiveRange();
  
  if (!activeRange || activeRange.getSheet().getName() !== CONFIG.RESPONSES_TAB_NAME) {
    return SpreadsheetApp.getUi().alert('⚠️ Please click on a student row inside the "Form Responses 1" tab first.');
  }
  
  const rowNum = activeRange.getRow();
  if (rowNum < 2) {
    return SpreadsheetApp.getUi().alert('⚠️ You have selected the header row. Please click on a valid student submission row.');
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const studentName = getCoalescedVal(headers, rowData, 'Student Full Name') || 'Unknown Student';
  const currentStatus = rowData[CONFIG.STATUS_COL_INDEX - 1] || 'Pending / Blank';
  
  const template = HtmlService.createTemplateFromFile('OverrideModal');
  template.rowNum = rowNum;
  template.studentName = studentName;
  template.currentStatus = currentStatus;
  
  const html = template.evaluate().setWidth(480).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Administrative State Override');
}

function executeAdminStatusOverride(rowNum, newStatus, adminNotes) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    
    sheet.getRange(rowNum, CONFIG.STATUS_COL_INDEX).setValue(newStatus);
    
    if (adminNotes && adminNotes.trim() !== '') {
      const existingNotes = sheet.getRange(rowNum, CONFIG.NOTES_COL_INDEX).getValue();
      const updatedNotes = `[Admin Override - ${newStatus}]: ${adminNotes}\n${existingNotes}`;
      sheet.getRange(rowNum, CONFIG.NOTES_COL_INDEX).setValue(updatedNotes);
    }
    
    const adminEmail = Session.getActiveUser().getEmail() || 'Registrar Admin';
    logSystemAudit('ADMIN_STATUS_OVERRIDE', `Status manually updated to '${newStatus}' by ${adminEmail}. Notes: ${adminNotes || 'None'}`, `Row ${rowNum}`);
    
    return `Success! Row #${rowNum} updated to: ${newStatus}`;
  } catch (e) {
    logSystemAudit('ERROR', `Override failed on Row ${rowNum}: ${e.message}`);
    throw new Error(e.message);
  } finally { lock.releaseLock(); }
}

function uiRetriggerEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Retrigger Approval Email', 'Enter the exact Row Number to re-ping the instructor:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText().trim(), 10);
    if (isNaN(rowNum) || rowNum < 2) return ui.alert('Error: Invalid row number.');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    if (rowData[CONFIG.STATUS_COL_INDEX - 1] !== 'Pending Faculty Approval') {
      return ui.alert('Cannot retrigger: Record is not in a Pending state.');
    }
    
    const timestamp    = rowData[0];
    const studentName  = getCoalescedVal(headers, rowData, 'Student Full Name');
    const course       = getCoalescedVal(headers, rowData, 'Course Code');
    const facultyEmail = getCoalescedVal(headers, rowData, 'Faculty Member\'s Email') || getCoalescedVal(headers, rowData, 'Your Email Address');
    
    sendFacultyApprovalEmail(timestamp, studentName, facultyEmail, course);
    logSystemAudit('ADMIN_RETRIGGER', `Manual re-ping sent to ${facultyEmail} by Admin.`, `Row ${rowNum}`);
    ui.alert(`Success! Email re-routed to ${facultyEmail}.`);
  }
}

/**
 * ===============================================================
 * SAFE EMAIL ROUTER
 * ===============================================================
 */
function sendRoutedEmail(originalTo, originalCc, subject, htmlBody) {
  let finalTo = originalTo, finalCc = originalCc, finalSubject = subject;

  if (CONFIG.TEST_MODE) {
    finalTo = CONFIG.TEST_EMAIL; finalCc = ''; finalSubject = `[TEST MODE] ${subject}`;
    htmlBody = `<div style="background-color: #FDC314; padding: 12px; margin-bottom: 20px; border-left: 6px solid #000000; color: #000000;"><strong>⚠️ WFU STAGING ENVIRONMENT NOTICE</strong><br><strong>Intended To:</strong> ${originalTo}<br><strong>Intended CC:</strong> ${originalCc || 'None'}</div>` + htmlBody;
  }
  MailApp.sendEmail({ to: finalTo, cc: finalCc, subject: finalSubject, htmlBody: htmlBody });
  logSystemAudit('EMAIL_DISPATCH', `Notification sent to: ${finalTo}. Subject: ${subject}`);
}

/**
 * ===============================================================
 * PILLAR 1: DATA NORMALIZATION TRIGGER
 * ===============================================================
 */
function updateFormDropdowns() {
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_DATA_SHEET_ID);
    const form = FormApp.openById(CONFIG.FORM_ID);
    const items = form.getItems(FormApp.ItemType.LIST); 
    
    const programData = masterSS.getSheetByName('Academic Programs').getRange(2, 5, masterSS.getSheetByName('Academic Programs').getLastRow() - 1, 1).getValues();
    const courseData = masterSS.getSheetByName('AD Lookup by Code').getRange(2, 1, masterSS.getSheetByName('AD Lookup by Code').getLastRow() - 1, 1).getValues();
    const programs = programData.map(row => row[0]).filter(String);
    const courses = courseData.map(row => row[0]).filter(String);
    
    let pCount = 0, cCount = 0;
    items.forEach(item => {
      if (item.getTitle().includes(CONFIG.FORM_QUESTION_PROGRAM)) { item.asListItem().setChoiceValues(programs); pCount++; } 
      else if (item.getTitle().includes(CONFIG.FORM_QUESTION_COURSE)) { item.asListItem().setChoiceValues(courses); cCount++; }
    });
    logSystemAudit('DB_SYNC', `Synchronized ${pCount} Program fields and ${cCount} Course fields with Master Data.`);
  } catch (e) { logSystemAudit('ERROR', `Dropdown Sync Failure: ${e.message}`); }
}

/**
 * ===============================================================
 * PILLAR 2: THE SELF-CALIBRATING INTAKE ROUTER
 * ===============================================================
 */
function onFormSubmit(e) {
  try {
    if (!e || !e.values) throw new Error("Execution aborted: Function triggered without form payload parameters.");

    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    if (!sheet) throw new Error(`Target ledger tab '${CONFIG.RESPONSES_TAB_NAME}' unresolvable.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = e.range ? e.range.getRow() : sheet.getLastRow();
    const responses = e.values; 
    
    const timestamp     = responses[0];
    const submitterRole = getCoalescedVal(headers, responses, 'role'); 
    
    const studentName  = getCoalescedVal(headers, responses, 'Student Full Name'); 
    const studentEmail = getCoalescedVal(headers, responses, 'Student Email Address');
    const studentId    = getCoalescedVal(headers, responses, 'Student ID Number');
    const program      = getCoalescedVal(headers, responses, 'Academic Program');
    const course       = getCoalescedVal(headers, responses, 'Course Code');
    const facultyEmail = getCoalescedVal(headers, responses, 'Faculty Member\'s Email') || getCoalescedVal(headers, responses, 'Your Email Address');

    if (submitterRole.includes('Faculty')) {
      const schedulePlan = getCoalescedVal(headers, responses, 'Detailed Incomplete Grade Plan');
      sheet.getRange(row, CONFIG.STATUS_COL_INDEX).setValue('Approved');
      sheet.getRange(row, CONFIG.NOTES_COL_INDEX).setValue(`[Faculty Bypass Plan]: ${schedulePlan}`);
      
      logSystemAudit('INTAKE_BYPASS', `Faculty direct submission. Auto-approved student: ${studentName}.`, timestamp);
      sendFinalEmails(timestamp, 'Approved', schedulePlan, studentName, studentEmail, studentId, course, facultyEmail, program);
    } else {
      sheet.getRange(row, CONFIG.STATUS_COL_INDEX).setValue('Pending Faculty Approval');
      logSystemAudit('INTAKE_ROUTING', `Student petition logged. Web App dispatched to: ${facultyEmail}.`, timestamp);
      sendFacultyApprovalEmail(timestamp, studentName, facultyEmail, course);
    }
  } catch (err) {
    logSystemAudit('CRITICAL_FAILURE', `Intake Router Exception: ${err.message}`);
  }
}

function sendFacultyApprovalEmail(timestamp, studentName, facultyEmail, course) {
  const webAppUrl = ScriptApp.getService().getUrl(); 
  const safeId = encodeURIComponent(timestamp); 
  const approveUrl = `${webAppUrl}?id=${safeId}&action=Approve`, denyUrl = `${webAppUrl}?id=${safeId}&action=Deny`;
  
  const subject = `ACTION REQUIRED: Incomplete Grade Request for ${studentName} (${course})`;
  const body = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #000000;">
      <div style="border-top: 5px solid #9E7E38; padding: 20px; background-color: #FFFFFF; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        <h2 style="color: #9E7E38; margin-top: 0;">Academic Grade Plan Review</h2>
        <p>Dear Faculty Member,</p>
        <p>An Incomplete Grade Request has been initiated by <strong>${studentName}</strong> for <strong>${course}</strong>. Per WFU Academic Policy, your authorization and an outline of remaining deliverables are required.</p>
        <p>Please select an action below to securely access the compliance portal:</p>
        <br>
        <p style="text-align: center;">
          <a href="${approveUrl}" style="background-color: #9E7E38; color: #FFFFFF; padding: 12px 18px; text-decoration: none; font-weight: bold; border-radius: 4px; display: inline-block; margin-right: 15px;">✅ Approve and submit Incomplete Plan</a>
          <a href="${denyUrl}" style="background-color: #000000; color: #FDC314; padding: 12px 18px; text-decoration: none; font-weight: bold; border-radius: 4px; display: inline-block;">❌ Deny Request</a>
        </p>
        <br><hr style="border: 0; border-top: 1px solid #E0E0E0; margin: 20px 0;">
        <div style="font-size: 12px; color: #555555; background-color: #FCFBF7; padding: 15px; border-left: 4px solid #9E7E38;">
          <strong>🔒 WFU Enterprise Security Notice:</strong><br>
          You must authenticate using your official <strong>@wfu.edu</strong> Google Workspace profile. If you encounter a "Drive access denied" error, right-click your chosen button and select <em>"Open link in Incognito / Private window"</em>.
        </div>
      </div>
    </div>
  `;
  sendRoutedEmail(facultyEmail, '', subject, body);
}

/**
 * ===============================================================
 * PILLAR 3: THE WEB APP (Structural Parity Processor)
 * ===============================================================
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.id = e.parameter.id; template.action = e.parameter.action;
  return template.evaluate().setTitle('WFU SPS Academic Portal').setFaviconUrl('https://www.wfu.edu/wp-content/themes/wfu-theme/favicon.ico').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processFacultyDecision(id, action, payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getValues();
    
    let targetRowIndex = -1, rowData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { targetRowIndex = i + 1; rowData = data[i]; break; }
    }
    if (targetRowIndex === -1) throw new Error("Record unresolvable. Data moved.");
    
    const status = action === 'Approve' ? 'Approved' : 'Denied';
    sheet.getRange(targetRowIndex, CONFIG.STATUS_COL_INDEX).setValue(status);
    
    const authenticatedEmail = Session.getActiveUser().getEmail() || getCoalescedVal(headers, rowData, 'Faculty Member\'s Email');
    const facultyNameFromStudent = getCoalescedVal(headers, rowData, 'Faculty Member\'s Full Name');
    
    // Find target parity columns dynamically
    const colFacName  = headers.findIndex(h => h.toLowerCase().includes('your full name')) + 1;
    const colFacEmail = headers.findIndex(h => h.toLowerCase().includes('your email address')) + 1;
    const colDeadline = headers.findIndex(h => h.toLowerCase().includes('final completion deadline')) + 1;
    const colPlan     = headers.findIndex(h => h.toLowerCase().includes('detailed incomplete grade plan')) + 1;
    
    let finalEmailNotes = "";

    if (action === 'Approve') {
      if (colFacName > 0) sheet.getRange(targetRowIndex, colFacName).setValue(`[Web Portal Auth]: ${facultyNameFromStudent}`);
      if (colFacEmail > 0) sheet.getRange(targetRowIndex, colFacEmail).setValue(authenticatedEmail);
      if (colDeadline > 0) sheet.getRange(targetRowIndex, colDeadline).setValue(payload.deadline);
      if (colPlan > 0) sheet.getRange(targetRowIndex, colPlan).setValue(`[Web Portal Lodged]: ${payload.plan}`);
      
      // DIGITAL AGREEMENT PARITY STAMP: Explicitly stamped to master ledger
      sheet.getRange(targetRowIndex, CONFIG.NOTES_COL_INDEX).setValue("Digital Agreement Confirmed via Secure Web Portal");
      
      finalEmailNotes = `Deadline: ${payload.deadline}\n\nPlan Details:\n${payload.plan}`;
      logSystemAudit('WEB_APP_ACTION', `Faculty (${authenticatedEmail}) securely logged Approved Plan via Web Portal.`, id);
    } else {
      sheet.getRange(targetRowIndex, CONFIG.NOTES_COL_INDEX).setValue(`[Web Portal Denied]: ${payload.denialReason}`);
      finalEmailNotes = payload.denialReason;
      logSystemAudit('WEB_APP_ACTION', `Faculty (${authenticatedEmail}) recorded Denial via Web Portal.`, id);
    }
    
    const studentName = getCoalescedVal(headers, rowData, 'Student Full Name');
    const studentEmail = getCoalescedVal(headers, rowData, 'Student Email Address');
    const studentId = getCoalescedVal(headers, rowData, 'Student ID Number');
    const program = getCoalescedVal(headers, rowData, 'Academic Program');
    const course = getCoalescedVal(headers, rowData, 'Course Code');

    sendFinalEmails(id, status, finalEmailNotes, studentName, studentEmail, studentId, course, authenticatedEmail, program);
    
    return `Transaction Recorded: ${status}. The master database has been successfully updated. You may now close this window.`;
  } catch (e) {
    logSystemAudit('ERROR', `Web App Concurrency Failure: ${e.message}`, id);
    return `System Exception: ${e.message}`;
  } finally { lock.releaseLock(); }
}

function lookupSSMEmail(studentId, studentEmail) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_DATA_SHEET_ID);
    const assignData = ss.getSheetByName('SSMAssignments').getDataRange().getValues(), ssmSheet = ss.getSheetByName('SSMs').getDataRange().getValues();
    let ssmName = null;
    for (let i = 1; i < assignData.length; i++) {
      const rowId = String(assignData[i][1]).trim(), rowEmail = String(assignData[i][7]).trim().toLowerCase();
      if (studentId && studentId !== 'N/A' && rowId === String(studentId).trim()) { ssmName = assignData[i][0]; break; }
      else if (studentEmail && rowEmail === String(studentEmail).trim().toLowerCase()) { ssmName = assignData[i][0]; break; }
    }
    if (!ssmName) return null;
    for (let i = 1; i < ssmSheet.length; i++) { if (ssmSheet[i][0] === ssmName) return ssmSheet[i][1]; }
    return null;
  } catch (e) { return null; }
}

function sendFinalEmails(id, status, notes, studentName, studentEmail, studentId, course, facultyEmail, program) {
  const ssmEmail = lookupSSMEmail(studentId, studentEmail);
  let ccList = `${CONFIG.KEY_PERSONNEL_EMAILS}, ${facultyEmail}, ${CONFIG.WATERMARK_COMPLIANCE_EMAIL}`;
  if (ssmEmail) ccList += `, ${ssmEmail}`;
  
  const subject = `Official Academic Notice: Incomplete Grade Plan ${status} - ${studentName} (${course})`;
  const body = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #000000;"><div style="border-top: 5px solid ${status === 'Approved' ? '#9E7E38' : '#000000'}; padding: 20px; background-color: #FFFFFF; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"><h2 style="color: #000000; margin-top: 0;">Incomplete Grade Plan: <span style="color: ${status === 'Approved' ? '#9E7E38' : '#D32F2F'};">${status}</span></h2><p>The formal Incomplete Grade petition for <strong>${studentName}</strong> (ID: ${studentId}) in course <strong>${course}</strong> (${program}) has been marked as <strong>${status}</strong> by the Instructor of Record.</p><div style="background-color: #FCFBF7; padding: 15px; margin: 15px 0; border-left: 4px solid #9E7E38;"><strong style="color: #9E7E38;">${status === 'Approved' ? 'Authoritative Completion Plan & Schedule:' : 'Academic Justification for Denial:'}</strong><br><p style="white-space: pre-wrap; margin-bottom: 0;">${notes || 'No institutional notes provided.'}</p></div><p><strong>Next Operational Steps:</strong></p><ul><li><strong>Student:</strong> If approved, you must submit all remaining deliverables via Canvas adhering strictly to the schedule above.</li><li><strong>SPS IT / Help:</strong> Verify Canvas section access is extended through the subsequent mini-session.</li><li><strong>Registrar / Student Services:</strong> Archive record for APR reporting and Watermark compliance tracking.</li></ul><hr style="border: 0; border-top: 1px solid #E0E0E0; margin: 20px 0;"><p style="font-size: 11px; color: #777777;"><em>This is an automated governance dispatch. Primary Key ID: ${id}</em></p></div></div>`;
  sendRoutedEmail(studentEmail, ccList, subject, body);
}

function executeKaraFailSafe() {
  const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues(), now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; 
    
    const timestamp = new Date(row[0]), status = row[CONFIG.STATUS_COL_INDEX - 1]; 
    const hoursElapsed = Math.abs(now - timestamp) / 36e5;
    
    if (status === 'Pending Faculty Approval' && hoursElapsed >= 48 && hoursElapsed <= 72) {
      const studentName  = getCoalescedVal(headers, row, 'Student Full Name');
      const course       = getCoalescedVal(headers, row, 'Course Code');
      const facultyEmail = getCoalescedVal(headers, row, 'Faculty Member\'s Email'); 
      
      logSystemAudit('FAILSAFE_TRIGGER', `Stalled queue escalated to Registrar for ${studentName}.`, row[0]);

      const subject = `⚠️ ESCALATION FAIL-SAFE: Stalled Incomplete Grade Request (${studentName})`;
      const body = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #000000; border-left: 5px solid #D32F2F; padding: 15px; background-color: #FFF9F9;"><h3 style="color: #D32F2F; margin-top: 0;">Administrative Failsafe Triggered</h3><p>Hello Kara,</p><p>An Incomplete Grade Request for <strong>${studentName}</strong> (${course}) has remained stalled in the faculty review queue for over 48 hours without authorization.</p><p><strong>System Diagnostics:</strong> The listed Instructor (<em>${facultyEmail}</em>) has received the secure Web App link but has not lodged a decision.</p><br><p><a href="https://docs.google.com/spreadsheets/d/${CONFIG.RESPONSES_SHEET_ID}/edit#gid=0&range=AB${i + 1}" style="background-color: #000000; color: #FDC314; padding: 10px 15px; text-decoration: none; font-weight: bold; border-radius: 4px;">🔍 Access Master Ledger (Row #${i + 1})</a></p></div>`;
      sendRoutedEmail(CONFIG.REGISTRAR_FAILSAFE_EMAIL, '', subject, body);
    }
  }
}