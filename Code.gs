/**
 * ============================================================================
 * WAKE FOREST UNIVERSITY - SCHOOL OF PROFESSIONAL STUDIES (SPS)
 * AUTOMATED INCOMPLETE GRADE & COMPLIANCE TRACKING ENGINE
 * ============================================================================
 * @author Amy Slack, MAT - Associate Director of IT Solutions
 * @architecture Solutions Architecture Partner
 * @version 5.1.0 (Production Master Build)
 * @license Proprietary - Wake Forest University
 * * SYSTEM OVERVIEW:
 * This script serves as the authoritative backend nervous system for the SPS 
 * Incomplete Grade petition workflow. It intercepts multi-branched Google Form 
 * submissions, executes sparse-array data coalescing, manages state transitions 
 * via an immutable audit ledger, and routes authenticated HTML compliance portals.
 */

/**
 * Global Configuration Constants
 * CENTRAL MAINTENANCE NOTE: Update these IDs if underlying Google infrastructure changes.
 */
const CONFIG = {
  // --- CORE INFRASTRUCTURE IDS ---
  MASTER_DATA_SHEET_ID: '1boHlW-x_Qr0Z11qfU40WUDSgq7aRFjVnpyJS5KZIIm8', // Source of truth for SSMs & Course Catalog
  FORM_ID: '1SjsqrXuUej4zp7xMSQ5rQDSZtlGCDoVFL4Zekw31MXQ',                 // Public student/faculty petition intake form
  RESPONSES_SHEET_ID: '18VE3YpjqO3cII6wHCTV-JBFCRzNjqmDfMW8D9nxbMBg',     // Master database ledger storing transactions
  RESPONSES_TAB_NAME: 'Form Responses 1',                                 // Primary intake sheet tab
  AUDIT_TAB_NAME: 'System_Audit_Log',                                     // Immutable system audit tracking tab

  // --- AUTHORITATIVE HARDCODED WEB APP ENDPOINT ---
  // CRITICAL ARCHITECTURAL NOTE: We hardcode the base URL to bypass the legacy V8 runtime 
  // 'Ghost ID' bug associated with ScriptApp.getService().getUrl() execution in background menus.
  // Must match the published URL from Deploy > Manage Deployments (strip any /a/wfu.edu/ domain paths).
  WEBAPP_BASE_URL: 'https://script.google.com/macros/s/AKfycbxpWQq_k4IU4tz0QAWOnZZ2g_6ez6lDA3tpoaYQ8iYpq2B7Rqc/exec',

  // --- FALLBACK LEDGER TARGETS (1-Indexed for getRange) ---
  // Used only if dynamic header scanning fails to resolve destination columns.
  STATUS_COL_FALLBACK: 28, // Default column AB
  NOTES_COL_FALLBACK: 29,  // Default column AC

  // --- STAKEHOLDER COMPLIANCE ROUTING EMAILS ---
  KEY_PERSONNEL_EMAILS: 'spsreg@wfu.edu, spshelp@wfu.edu, stuserv@wfu.edu',
  REGISTRAR_FAILSAFE_EMAIL: 'traverk@wfu.edu', // Kara Traverse (Registrar Escalation)
  WATERMARK_COMPLIANCE_EMAIL: 'maguirl@wfu.edu', // Loréal Maguire (Accreditation Archive)

  // --- ENVIRONMENT GOVERNANCE ---
  TEST_MODE: false, // SET TO FALSE FOR LIVE STUDENT/FACULTY ROUTING
  TEST_EMAIL: 'slacka@wfu.edu' 
};

/**
 * ============================================================================
 * SECTION 1: CORE ENGINE UTILITIES & DATA NORMALIZATION
 * ============================================================================
 */

/**
 * Dynamic Header Coalescing Engine (Column-Shift Immunity)
 * Scans Row 1 headers for a partial substring match and returns the active row value.
 * Protects runtime against Google Forms natively inserting or rearranging spreadsheet columns.
 * * @param {Array<string>} headers - Master title strings from Row 1
 * @param {Array<any>} rowValues - Submitted data array for active transaction
 * @param {string} searchString - Substring syntax to identify target column
 * @returns {any|string} Extracted non-empty cell value or empty string
 */
function getCoalescedVal(headers, rowValues, searchString) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && String(headers[i]).toLowerCase().includes(searchString.toLowerCase())) {
      if (rowValues[i] !== undefined && rowValues[i] !== null && String(rowValues[i]).trim() !== '') {
        return rowValues[i];
      }
    }
  }
  return '';
}

/**
 * Resolves destination column index dynamically by header title.
 * Prevents writing status/notes over student data if form columns shift horizontally.
 * * @param {Array<string>} headers - Master Row 1 header array
 * @param {string} searchString - Column header title to locate
 * @param {number} fallbackIndex - 1-indexed fallback position if missing
 * @returns {number} 1-indexed column coordinate for Sheets API operations
 */
function getDynamicColIndex(headers, searchString, fallbackIndex) {
  const idx = headers.findIndex(h => h && String(h).toLowerCase().includes(searchString.toLowerCase()));
  return idx !== -1 ? idx + 1 : fallbackIndex;
}

/**
 * Converts varying Timestamp formats (JS Date vs Sheets String) into universal Epoch milliseconds.
 * Eliminates primary-key matching failures during Web App AJAX handoffs.
 * * @param {string|Date} val - Raw timestamp representation
 * @returns {number} Absolute Epoch millisecond integer
 */
function normalizeTimestampToEpoch(val) {
  if (!val) return 0;
  const d = new Date(val);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

/**
 * Safe Email Router wrapper.
 * Intercepts outgoing MailApp dispatches during staging to prevent accidental stakeholder spam.
 */
function sendRoutedEmail(originalTo, originalCc, subject, htmlBody) {
  let finalTo = originalTo, finalCc = originalCc, finalSubject = subject;
  
  if (CONFIG.TEST_MODE) {
    finalTo = CONFIG.TEST_EMAIL; 
    finalCc = ''; 
    finalSubject = `[TEST MODE] ${subject}`;
    htmlBody = `<div style="background-color: #FDC314; padding: 12px; margin-bottom: 20px; border-left: 6px solid #000000; color: #000000;"><strong>⚠️ WFU STAGING ENVIRONMENT NOTICE</strong><br><strong>Intended To:</strong> ${originalTo}<br><strong>Intended CC:</strong> ${originalCc || 'None'}</div>` + htmlBody;
  }
  
  MailApp.sendEmail({ to: finalTo, cc: finalCc, subject: finalSubject, htmlBody: htmlBody });
  logSystemAudit('EMAIL_DISPATCH', `Notification dispatched to: ${finalTo}`);
}

/**
 * ============================================================================
 * SECTION 2: PILLAR VI - ADMINISTRATIVE TOOLING SUITE & AUDIT ENGINE
 * ============================================================================
 */

/**
 * Builds custom WFU administrative menu upon spreadsheet initialization.
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

/**
 * Master Auditing Engine.
 * Logs all system events, manual overrides, and routing actions to a persistent, 
 * immutable ledger tab. Self-heals by regenerating the tab if accidentally deleted.
 */
function logSystemAudit(eventType, details, primaryKey = 'System') {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    let auditSheet = ss.getSheetByName(CONFIG.AUDIT_TAB_NAME);
    
    // Auto-construct audit tab with WFU branding if missing
    if (!auditSheet) {
      auditSheet = ss.insertSheet(CONFIG.AUDIT_TAB_NAME);
      auditSheet.appendRow(['Timestamp', 'Event Type', 'Primary Key (ID / Row)', 'Transaction Details']);
      auditSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#000000').setFontColor('#FDC314');
      auditSheet.setFrozenRows(1);
      auditSheet.setColumnWidths(1, 4, [150, 120, 120, 500]);
    }
    auditSheet.appendRow([new Date(), eventType, primaryKey, details]);
  } catch (e) {
    console.error(`CRITICAL: Audit Engine Failure: ${e.message}`);
  }
}

/**
 * Launches interactive HTML modal viewer displaying the last 50 audit logs.
 */
function launchAuditPanel() {
  const html = HtmlService.createTemplateFromFile('AuditPanel').evaluate()
      .setWidth(850).setHeight(550).setTitle('SPS System Audit History');
  SpreadsheetApp.getUi().showModalDialog(html, 'System Audit & Event Tracer');
}

/**
 * Programmatic feeder for AuditPanel.html front-end.
 */
function getRecentAuditLogs() {
  const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
  const auditSheet = ss.getSheetByName(CONFIG.AUDIT_TAB_NAME);
  if (!auditSheet) return [];
  const data = auditSheet.getDataRange().getDisplayValues();
  return data.length <= 1 ? [] : data.slice(1).reverse().slice(0, 50);
}

/**
 * Active-Cell UI Modal Launcher (Pattern 1 Override).
 * Allows Registrar staff to update student status directly from cursor focus without horizontal scrolling.
 */
function openStatusOverrideModal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.RESPONSES_TAB_NAME);
  const activeRange = SpreadsheetApp.getActiveRange();
  
  if (!activeRange || activeRange.getSheet().getName() !== CONFIG.RESPONSES_TAB_NAME) {
    return SpreadsheetApp.getUi().alert('⚠️ Click on a student row inside "Form Responses 1" first.');
  }
  const rowNum = activeRange.getRow();
  if (rowNum < 2) {
    return SpreadsheetApp.getUi().alert('⚠️ Header row selected. Click a valid student row.');
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const studentName = getCoalescedVal(headers, rowData, 'Student Full Name') || 'Unknown Student';
  const colStatusIdx = headers.findIndex(h => h && String(h).toLowerCase().includes('status'));
  const currentStatus = colStatusIdx !== -1 ? rowData[colStatusIdx] : rowData[CONFIG.STATUS_COL_FALLBACK - 1];
  
  const template = HtmlService.createTemplateFromFile('OverrideModal');
  template.rowNum = rowNum; 
  template.studentName = studentName; 
  template.currentStatus = currentStatus || 'Pending';
  
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(480).setHeight(380), 'Administrative State Override');
}

/**
 * Execution receiver for manual administrative status overrides.
 */
function executeAdminStatusOverride(rowNum, newStatus, adminNotes) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colStatus = getDynamicColIndex(headers, 'Status', CONFIG.STATUS_COL_FALLBACK);
    const colNotes  = getDynamicColIndex(headers, 'System Notes', CONFIG.NOTES_COL_FALLBACK);

    sheet.getRange(rowNum, colStatus).setValue(newStatus);
    if (adminNotes && adminNotes.trim() !== '') {
      const existingNotes = sheet.getRange(rowNum, colNotes).getValue();
      sheet.getRange(rowNum, colNotes).setValue(`[Admin Override - ${newStatus}]: ${adminNotes}\n${existingNotes}`);
    }
    logSystemAudit('ADMIN_STATUS_OVERRIDE', `Status updated to '${newStatus}'. Notes: ${adminNotes}`, `Row ${rowNum}`);
    return `Success! Row #${rowNum} updated to: ${newStatus}`;
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Manual prompt tool allowing administrators to retrigger lost faculty Web App links.
 */
function uiRetriggerEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Retrigger Approval Email', 'Enter exact Row Number:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText().trim(), 10);
    if (isNaN(rowNum) || rowNum < 2) return ui.alert('Error: Invalid row number.');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const timestamp = rowData[0];
    const studentName = getCoalescedVal(headers, rowData, 'Student Full Name');
    const course = getCoalescedVal(headers, rowData, 'Course Code');
    const facultyEmail = getCoalescedVal(headers, rowData, 'Faculty Member\'s Email') || getCoalescedVal(headers, rowData, 'Your Email Address');
    
    sendFacultyApprovalEmail(timestamp, studentName, facultyEmail, course);
    logSystemAudit('ADMIN_RETRIGGER', `Manual re-ping sent to ${facultyEmail}`, `Row ${rowNum}`);
    ui.alert(`Success! Email re-routed to ${facultyEmail}.`);
  }
}

/**
 * ============================================================================
 * SECTION 3: PILLAR I & II - PRIMARY INTAKE & COALESCING ROUTER
 * ============================================================================
 */

/**
 * Primary Form Intake Router.
 * Bound via Installable Trigger. Evaluates persona role selection, manages sparse form arrays,
 * writes initial database states, and dispatches path-specific web gateways.
 */
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30-second concurrency lock to queue simultaneous student submissions
    if (!e || !e.values) throw new Error("Execution aborted: Triggered without form payload parameters.");

    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    if (!sheet) throw new Error(`Target tab '${CONFIG.RESPONSES_TAB_NAME}' unresolvable.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = e.range ? e.range.getRow() : sheet.getLastRow();
    const responses = e.values; 
    
    const timestamp     = responses[0];
    const submitterRole = getCoalescedVal(headers, responses, 'role'); 
    const studentName   = getCoalescedVal(headers, responses, 'Student Full Name'); 
    const studentEmail  = getCoalescedVal(headers, responses, 'Student Email Address');
    const studentId     = getCoalescedVal(headers, responses, 'Student ID Number');
    const program       = getCoalescedVal(headers, responses, 'Academic Program');
    const course        = getCoalescedVal(headers, responses, 'Course Code');
    const facultyEmail  = getCoalescedVal(headers, responses, 'Faculty Member\'s Email') || getCoalescedVal(headers, responses, 'Your Email Address');

    const colStatus = getDynamicColIndex(headers, 'Status', CONFIG.STATUS_COL_FALLBACK);
    const colNotes  = getDynamicColIndex(headers, 'System Notes', CONFIG.NOTES_COL_FALLBACK);

    // PATH A: Faculty Bypass Intake (Instant Approval)
    if (submitterRole.includes('Faculty')) {
      const schedulePlan = getCoalescedVal(headers, responses, 'Detailed Incomplete Grade Plan');
      sheet.getRange(row, colStatus).setValue('Approved');
      sheet.getRange(row, colNotes).setValue(`[Faculty Bypass Plan]: ${schedulePlan}`);
      
      logSystemAudit('INTAKE_BYPASS', `Faculty direct submission. Auto-approved: ${studentName}.`, timestamp);
      sendFinalEmails(timestamp, 'Approved', schedulePlan, studentName, studentEmail, studentId, course, facultyEmail, program);
    } 
    // PATH B: Student Intake (Quarantine to Pending State)
    else {
      sheet.getRange(row, colStatus).setValue('Pending Faculty Approval');
      logSystemAudit('INTAKE_ROUTING', `Student petition logged. Portal dispatched to: ${facultyEmail}.`, timestamp);
      sendFacultyApprovalEmail(timestamp, studentName, facultyEmail, course);
    }
  } catch (err) {
    logSystemAudit('CRITICAL_FAILURE', `Intake Router Exception: ${err.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Formulates and dispatches HTML Action Email containing secure compliance portal links.
 */
function sendFacultyApprovalEmail(timestamp, studentName, facultyEmail, course) {
  const webAppUrl = CONFIG.WEBAPP_BASE_URL; 
  const safeId = encodeURIComponent(timestamp); 
  const approveUrl = `${webAppUrl}?id=${safeId}&action=Approve`;
  const denyUrl    = `${webAppUrl}?id=${safeId}&action=Deny`;
  
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
 * ============================================================================
 * SECTION 4: PILLAR III - WEB APP PARITY PROCESSOR
 * ============================================================================
 */

/**
 * HTTP GET Gateway for secure HTML portal.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.id = e.parameter.id; 
  template.action = e.parameter.action;
  return template.evaluate()
    .setTitle('WFU SPS Academic Portal')
    .setFaviconUrl('https://www.wfu.edu/wp-content/themes/wfu-theme/favicon.ico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Concurrency receiver for Web App AJAX payloads.
 * Maps Web App fields back into native Google Form schema columns (P, Q, Z, AA)
 * to maintain 100% downstream reporting parity.
 */
function processFacultyDecision(id, action, payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); 
    const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getValues();
    
    const targetEpoch = normalizeTimestampToEpoch(decodeURIComponent(id));
    let targetRowIndex = -1, rowData = null;
    
    // Locate target row mathematically
    for (let i = 1; i < data.length; i++) {
      if (normalizeTimestampToEpoch(data[i][0]) === targetEpoch && targetEpoch !== 0) {
        targetRowIndex = i + 1; 
        rowData = data[i]; 
        break; 
      }
    }
    if (targetRowIndex === -1) throw new Error("Record unresolvable or primary key expired.");
    
    const status = action === 'Approve' ? 'Approved' : 'Denied';
    const colStatus = getDynamicColIndex(headers, 'Status', CONFIG.STATUS_COL_FALLBACK);
    const colNotes  = getDynamicColIndex(headers, 'System Notes', CONFIG.NOTES_COL_FALLBACK);

    sheet.getRange(targetRowIndex, colStatus).setValue(status);
    
    const authenticatedEmail = Session.getActiveUser().getEmail() || getCoalescedVal(headers, rowData, 'Faculty Member\'s Email');
    const facultyNameFromStudent = getCoalescedVal(headers, rowData, 'Faculty Member\'s Full Name');
    
    // Dynamically locate native Section 3 Form columns
    const colFacName  = headers.findIndex(h => h && String(h).toLowerCase().includes('your full name')) + 1;
    const colFacEmail = headers.findIndex(h => h && String(h).toLowerCase().includes('your email address')) + 1;
    const colDeadline = headers.findIndex(h => h && String(h).toLowerCase().includes('final completion deadline')) + 1;
    const colPlan     = headers.findIndex(h => h && String(h).toLowerCase().includes('detailed incomplete grade plan')) + 1;
    
    let finalEmailNotes = "";

    if (action === 'Approve') {
      if (colFacName > 0) sheet.getRange(targetRowIndex, colFacName).setValue(`[Portal Auth]: ${facultyNameFromStudent}`);
      if (colFacEmail > 0) sheet.getRange(targetRowIndex, colFacEmail).setValue(authenticatedEmail);
      if (colDeadline > 0) sheet.getRange(targetRowIndex, colDeadline).setValue(payload.deadline);
      if (colPlan > 0) sheet.getRange(targetRowIndex, colPlan).setValue(`[Portal Lodged]: ${payload.plan}`);
      
      // Mandatory digital compliance stamp
      sheet.getRange(targetRowIndex, colNotes).setValue("Digital Agreement Confirmed via Secure Web Portal");
      finalEmailNotes = `Deadline: ${payload.deadline}\n\nPlan Details:\n${payload.plan}`;
      logSystemAudit('WEB_APP_ACTION', `Faculty (${authenticatedEmail}) authorized Approved Plan.`, id);
    } else {
      sheet.getRange(targetRowIndex, colNotes).setValue(`[Portal Denied]: ${payload.denialReason}`);
      finalEmailNotes = payload.denialReason;
      logSystemAudit('WEB_APP_ACTION', `Faculty (${authenticatedEmail}) recorded Denial.`, id);
    }
    
    const studentName  = getCoalescedVal(headers, rowData, 'Student Full Name');
    const studentEmail = getCoalescedVal(headers, rowData, 'Student Email Address');
    const studentId    = getCoalescedVal(headers, rowData, 'Student ID Number');
    const program      = getCoalescedVal(headers, rowData, 'Academic Program');
    const course       = getCoalescedVal(headers, rowData, 'Course Code');

    sendFinalEmails(id, status, finalEmailNotes, studentName, studentEmail, studentId, course, authenticatedEmail, program);
    return `Transaction Recorded: ${status}. Master database updated. You may now close this window.`;
  } catch (e) {
    logSystemAudit('ERROR', `Web App Concurrency Failure: ${e.message}`, id);
    return `System Exception: ${e.message}`;
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * SECTION 5: HELPER ROUTINES & AUTOMATED DOWNSTREAM TRIGGERS
 * ============================================================================
 */

/**
 * Performs relational VLOOKUP across Master Data file to retrieve Student Support Manager (SSM) email.
 */
function lookupSSMEmail(studentId, studentEmail) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_DATA_SHEET_ID);
    const assignData = ss.getSheetByName('SSMAssignments').getDataRange().getValues();
    const ssmSheet   = ss.getSheetByName('SSMs').getDataRange().getValues();
    let ssmName = null;
    
    for (let i = 1; i < assignData.length; i++) {
      const rowId = String(assignData[i][1]).trim();
      const rowEmail = String(assignData[i][7]).trim().toLowerCase();
      if (studentId && studentId !== 'N/A' && rowId === String(studentId).trim()) { 
        ssmName = assignData[i][0]; break; 
      } else if (studentEmail && rowEmail === String(studentEmail).trim().toLowerCase()) { 
        ssmName = assignData[i][0]; break; 
      }
    }
    if (!ssmName) return null;
    for (let i = 1; i < ssmSheet.length; i++) { 
      if (ssmSheet[i][0] === ssmName) return ssmSheet[i][1]; 
    }
    return null;
  } catch (e) { return null; }
}

/**
 * Formulates authoritative notice of final decision to Student, Faculty, Registrar, Help Desk, and SSM.
 */
function sendFinalEmails(id, status, notes, studentName, studentEmail, studentId, course, facultyEmail, program) {
  const ssmEmail = lookupSSMEmail(studentId, studentEmail);
  let ccList = `${CONFIG.KEY_PERSONNEL_EMAILS}, ${facultyEmail}, ${CONFIG.WATERMARK_COMPLIANCE_EMAIL}`;
  if (ssmEmail) ccList += `, ${ssmEmail}`;
  
  const subject = `Official Academic Notice: Incomplete Grade Plan ${status} - ${studentName} (${course})`;
  const body = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #000000;"><div style="border-top: 5px solid ${status === 'Approved' ? '#9E7E38' : '#000000'}; padding: 20px; background-color: #FFFFFF; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"><h2 style="color: #000000; margin-top: 0;">Incomplete Grade Plan: <span style="color: ${status === 'Approved' ? '#9E7E38' : '#D32F2F'};">${status}</span></h2><p>The formal Incomplete Grade petition for <strong>${studentName}</strong> (ID: ${studentId}) in course <strong>${course}</strong> (${program}) has been marked as <strong>${status}</strong> by the Instructor of Record.</p><div style="background-color: #FCFBF7; padding: 15px; margin: 15px 0; border-left: 4px solid #9E7E38;"><strong style="color: #9E7E38;">${status === 'Approved' ? 'Authoritative Completion Plan & Schedule:' : 'Academic Justification for Denial:'}</strong><br><p style="white-space: pre-wrap; margin-bottom: 0;">${notes || 'No institutional notes provided.'}</p></div><p><strong>Next Operational Steps:</strong></p><ul><li><strong>Student:</strong> If approved, submit remaining deliverables via Canvas adhering strictly to the schedule above.</li><li><strong>SPS IT / Help:</strong> Verify Canvas section access is extended through the subsequent mini-session.</li><li><strong>Registrar / Student Services:</strong> Archive record for APR reporting and Watermark compliance tracking.</li></ul><hr style="border: 0; border-top: 1px solid #E0E0E0; margin: 20px 0;"><p style="font-size: 11px; color: #777777;"><em>Automated governance dispatch. Key ID: ${id}</em></p></div></div>`;
  sendRoutedEmail(studentEmail, ccList, subject, body);
}

/**
 * Nightly time-driven maintenance routine.
 * Dynamically synchronizes Google Form 'Academic Program' and 'Course Code' dropdown lists 
 * against the master curriculum portfolio spreadsheet maintained by Academic Programs.
 */
function updateFormDropdowns() {
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_DATA_SHEET_ID);
    const form = FormApp.openById(CONFIG.FORM_ID);
    const items = form.getItems(FormApp.ItemType.LIST); 
    
    const programData = masterSS.getSheetByName('Academic Programs').getRange(2, 5, masterSS.getSheetByName('Academic Programs').getLastRow() - 1, 1).getValues();
    const courseData  = masterSS.getSheetByName('AD Lookup by Code').getRange(2, 1, masterSS.getSheetByName('AD Lookup by Code').getLastRow() - 1, 1).getValues();
    const programs = programData.map(row => row[0]).filter(String);
    const courses  = courseData.map(row => row[0]).filter(String);
    
    items.forEach(item => {
      if (item.getTitle().includes(CONFIG.FORM_QUESTION_PROGRAM)) item.asListItem().setChoiceValues(programs); 
      else if (item.getTitle().includes(CONFIG.FORM_QUESTION_COURSE)) item.asListItem().setChoiceValues(courses);
    });
    logSystemAudit('DB_SYNC', `Successfully synced form dropdowns with Master Curriculum catalog.`);
  } catch (e) { 
    logSystemAudit('ERROR', `Dropdown Sync Failure: ${e.message}`); 
  }
}

/**
 * Daily morning time-driven escalation engine (Pillar IV).
 * Scans active triage queue. If a student petition remains 'Pending Faculty Approval' 
 * between 48 and 72 hours, automatically dispatches an administrative alert to the Registrar.
 */
function executeKaraFailSafe() {
  const ss = SpreadsheetApp.openById(CONFIG.RESPONSES_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.RESPONSES_TAB_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues(), now = new Date();
  const colStatusIdx = headers.findIndex(h => h && String(h).toLowerCase().includes('status'));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty buffer rows
    
    const timestamp = new Date(row[0]);
    const status = colStatusIdx !== -1 ? row[colStatusIdx] : row[CONFIG.STATUS_COL_FALLBACK - 1]; 
    const hoursElapsed = Math.abs(now - timestamp) / 36e5;
    
    if (status === 'Pending Faculty Approval' && hoursElapsed >= 48 && hoursElapsed <= 72) {
      const studentName  = getCoalescedVal(headers, row, 'Student Full Name');
      const course       = getCoalescedVal(headers, row, 'Course Code');
      const facultyEmail = getCoalescedVal(headers, row, 'Faculty Member\'s Email'); 
      
      logSystemAudit('FAILSAFE_TRIGGER', `Stalled queue escalated to Registrar for ${studentName}.`, row[0]);

      const subject = `⚠️ ESCALATION FAIL-SAFE: Stalled Incomplete Grade Request (${studentName})`;
      const body = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #000000; border-left: 5px solid #D32F2F; padding: 15px; background-color: #FFF9F9;"><h3 style="color: #D32F2F; margin-top: 0;">Administrative Failsafe Triggered</h3><p>Hello Kara,</p><p>An Incomplete Grade Request for <strong>${studentName}</strong> (${course}) has remained stalled in the faculty review queue for over 48 hours without authorization.</p><p><strong>System Diagnostics:</strong> The listed Instructor (<em>${facultyEmail}</em>) received the portal link but has not lodged a decision.</p><br><p><a href="https://docs.google.com/spreadsheets/d/${CONFIG.RESPONSES_SHEET_ID}/edit#gid=0&range=AB${i + 1}" style="background-color: #000000; color: #FDC314; padding: 10px 15px; text-decoration: none; font-weight: bold; border-radius: 4px;">🔍 Access Master Ledger (Row #${i + 1})</a></p></div>`;
      sendRoutedEmail(CONFIG.REGISTRAR_FAILSAFE_EMAIL, '', subject, body);
    }
  }
}