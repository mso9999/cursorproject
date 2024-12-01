/**
 * NotificationSystem.gs
 * 1PWR Procurement System - Email Notification and Reminder System
 * Version: 1.1
 * Last Updated: 2024-11-06
 * 
 * Purpose:
 * --------
 * Handles all automated email communications for the procurement system including:
 * - Initial submission notifications
 * - Status change alerts
 * - Reminder schedules for ordered items
 * - Overdue warnings and auto-cancellation notices
 * 
 * Dependencies:
 * ------------
 * - Code.gs: Contains CONFIG and COL constants for sheet structure
 * - Master Log sheet: Source of PO data and status tracking
 * - Notification Log sheet: Records all sent notifications
 * - Reminders Log sheet: Tracks reminder schedules and intervals
 * 
 * Sheet Structure Requirements:
 * --------------------------
 * 1. Master Log sheet columns (defined in COL constant):
 *    - PO Number (A)
 *    - Status (U)
 *    - Expected Landing Date (AK)
 *    - Shipped (AL)
 *    - Customs Required (AN)
 *    - Customs Cleared (AQ)
 *    - Goods Landed (AS)
 * 
 * 2. Notification Log sheet columns:
 *    - Timestamp
 *    - Type
 *    - PR Number
 *    - Recipients
 *    - Sent By
 * 
 * 3. Reminders Log sheet columns:
 *    - PR Number
 *    - Last Reminder Date
 *    - Reminder Count
 *    - Current Interval
 *    - Blocking Item
 * 
 * Business Rules:
 * -------------
 * 1. Reminder Intervals:
 *    - First reminder: 5 days
 *    - Second reminder: 2.5 days
 *    - Third reminder: 1.25 days
 *    - Subsequent: Daily
 * 
 * 2. Auto-cancellation:
 *    - Warning at 30 days overdue
 *    - Cancellation at 40 days if no update
 * 
 * 3. Recipient Rules:
 *    - Standard notifications: Procurement, Requestor
 *    - Overdue warnings: + Finance, Approvers
 *    - Customs warnings: Procurement, Finance, Approvers
 * 
 * Change History:
 * -------------
 * 2024-11-06: 
 *   - Consolidated duplicate code
 *   - Added comprehensive template validation
 *   - Improved error handling
 */

/**
 * Required fields by template type
 * Maps each template to its required data fields
 */
const TEMPLATE_REQUIREMENTS = {
  SUBMISSION: ['number', 'description', 'submitter', 'department', 'expectedDate', 'link'],
  STATUS_CHANGE: ['number', 'description', 'oldStatus', 'newStatus', 'user', 'link'],
  CANCELLATION_WARNING: ['number', 'description', 'daysOverdue', 'daysUntilCancel', 'expectedLandingDate', 'updateLink', 'link'],
  CUSTOMS_WARNING_INITIAL: ['number', 'description', 'daysInCustoms', 'submissionDate', 'link'],
  CUSTOMS_WARNING_REPEAT: ['number', 'description', 'daysInCustoms', 'submissionDate', 'link']
};

/**
 * Email template configuration
 * Defines all notification templates used in the system
 * Templates use {placeholder} syntax for dynamic content
 */
const EMAIL_TEMPLATES = {
  // Initial submission notification
  SUBMISSION: {
    subject: 'PR {number} - {description} - New Submission',
    body: `
New Purchase Requisition submitted:
PR Number: {number}
Description: {description}
Submitted By: {submitter}
Department: {department}
Expected Date: {expectedDate}

View PR: {link}
    `
  },

  // Status change notification
  STATUS_CHANGE: {
    subject: 'PR {number} - Status Changed to {newStatus}',
    body: `
Status Update for Purchase Requisition:
PR Number: {number}
Description: {description}
Previous Status: {oldStatus}
New Status: {newStatus}
Updated By: {user}

{notes}

View PR: {link}
    `
  },

  // Cancellation warning
  CANCELLATION_WARNING: {
    subject: 'URGENT: PR {number} - Auto-cancellation in {daysUntilCancel} days',
    body: `
WARNING: This PR will be automatically canceled in {daysUntilCancel} days if no action is taken.

PR Number: {number}
Description: {description}
Days Overdue: {daysOverdue}
Expected Landing Date: {expectedLandingDate}

To prevent cancellation, please update the Expected Landing Date:
{updateLink}

View PR: {link}
    `
  },

  // Initial customs warning
  CUSTOMS_WARNING_INITIAL: {
    subject: 'PR {number} - Customs Clearance Alert',
    body: `
Customs Clearance Alert for Purchase Requisition:
PR Number: {number}
Description: {description}
Days in Customs: {daysInCustoms}
Submission Date: {submissionDate}

Please follow up on customs clearance status.
View PR: {link}
    `
  },

  // Repeat customs warning
  CUSTOMS_WARNING_REPEAT: {
    subject: 'URGENT: PR {number} - Extended Customs Clearance Time',
    body: `
URGENT: Extended Customs Clearance Time Alert
PR Number: {number}
Description: {description}
Days in Customs: {daysInCustoms}
Submission Date: {submissionDate}

Immediate follow-up required on customs clearance status.
View PR: {link}
    `
  }
};

/**
 * Core Notification Functions
 * =========================
 */

/**
 * Main notification sending function
 * --------------------------------
 * Validates data, processes template, and sends email
 * 
 * @param {string} templateKey - Key of template in EMAIL_TEMPLATES
 * @param {Object} data - Data to populate template placeholders
 * @param {Array<string>} recipients - Array of email addresses
 * @returns {boolean} Success status
 * 
 * Example:
 * sendNotification('SUBMISSION', {
 *   number: 'PR-202411-001',
 *   description: 'Office Supplies',
 *   submitter: 'John Doe',
 *   department: 'Operations',
 *   expectedDate: '2024-12-01',
 *   link: 'https://...'
 * }, ['procurement@1pwrafrica.com', 'john.doe@1pwrafrica.com']);
 */
function sendNotification(templateKey, data, recipients) {
  console.log(`Sending ${templateKey} notification to:`, recipients);
  
  try {
    // Validate template exists
    const template = EMAIL_TEMPLATES[templateKey];
    if (!template) {
      throw new Error(`Template not found: ${templateKey}`);
    }

    // Validate required fields
    const missingFields = validateRequiredFields(templateKey, data);
    if (missingFields.length > 0) {
      throw new Error(`Missing required fields: ${missingFields.join(', ')}`);
    }

    // Filter and validate recipients
    const validRecipients = validateRecipients(recipients);
    if (validRecipients.length === 0) {
      throw new Error('No valid recipients after filtering');
    }

    // Process template
    const { subject, body } = processTemplate(template, data);

    // Send email
    GmailApp.sendEmail(
        validRecipients.join(','),
        subject,
        body
    );

    // Log notification
    logNotification(templateKey, data.number, validRecipients);
    
    console.log(`Successfully sent ${templateKey} notification`);
    return true;

  } catch (error) {
    console.error('Error sending notification:', error);
    return false;
  }
}

/**
 * Validates required fields for template
 * -----------------------------------
 * Checks if all required fields are present in data
 * 
 * @param {string} templateKey - Template identifier
 * @param {Object} data - Data to validate
 * @returns {Array} List of missing field names
 */
function validateRequiredFields(templateKey, data) {
  const requiredFields = TEMPLATE_REQUIREMENTS[templateKey] || [];
  return requiredFields.filter(field => !data[field]);
}

/**
 * Validates and filters recipient email addresses
 * -------------------------------------------
 * 
 * @param {Array<string>} recipients - Array of email addresses
 * @returns {Array<string>} Array of valid email addresses
 */
function validateRecipients(recipients) {
  if (!Array.isArray(recipients)) {
    console.warn('Recipients must be an array');
    return [];
  }

  return recipients.filter(email => {
    // Remove undefined/null values
    if (!email) return false;

    // Basic email format validation
    const isValid = email.includes('@') && 
                   email.includes('.') && 
                   !email.includes(' ') &&
                   email.length > 5;
    
    if (!isValid) {
      console.warn(`Invalid email address filtered out: ${email}`);
    }
    return isValid;
  });
}

/**
 * Processes email template
 * ----------------------
 * Replaces placeholders with actual data
 * 
 * @param {Object} template - Email template
 * @param {Object} data - Data for placeholders
 * @returns {Object} Processed subject and body
 */
function processTemplate(template, data) {
  let subject = template.subject;
  let body = template.body;

  // Replace all placeholders
  Object.keys(data).forEach(key => {
    const placeholder = `{${key}}`;
    const value = data[key] || '';
    
    // Replace all occurrences in both subject and body
    subject = subject.replace(new RegExp(placeholder, 'g'), value);
    body = body.replace(new RegExp(placeholder, 'g'), value);
  });

  // Remove any remaining placeholders
  subject = subject.replace(/\{[^}]+\}/g, '');
  body = body.replace(/\{[^}]+\}/g, '');

  return { subject, body };
}

/**
 * Logs sent notification to tracking sheet
 * ------------------------------------
 * 
 * @param {string} type - Notification type
 * @param {string} prNumber - PR number
 * @param {Array<string>} recipients - Recipients
 */
function logNotification(type, prNumber, recipients) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('Notification Log');
    
    if (!logSheet) {
      console.error('Notification Log sheet not found');
      return;
    }

    // Add log entry
    logSheet.appendRow([
      new Date(),                    // Timestamp
      type,                         // Notification type
      prNumber,                     // PR number
      recipients.join(', '),        // Recipients
      Session.getActiveUser().getEmail() // Sender
    ]);

  } catch (error) {
    console.error('Error logging notification:', error);
    // Don't throw - logging failure shouldn't stop the notification process
  }
}

/**
 * Reminder Processing Functions
 * ===========================
 */

/**
 * Main reminder processing function
 * -------------------------------
 * Called daily by time-driven trigger
 * Handles daily reminder checks and notifications
 * 
 * Process Flow:
 * 1. Loads all active POs
 * 2. Identifies items needing reminders
 * 3. Processes each type of reminder
 * 4. Updates reminder logs
 */
function processReminders() {
  console.log('Starting daily reminder processing');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);
    const data = sheet.getDataRange().getValues();
    
    // Process statistics
    let stats = {
      processed: 0,
      reminders: 0,
      errors: 0
    };

    // Process each PO (skip header row)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[COL.STATUS];
      stats.processed++;
      
      try {
        // Process based on status
        if (status === 'Ordered') {
          const reminderSent = processOrderedReminders(row, i + 1);
          if (reminderSent) stats.reminders++;
        }
      } catch (error) {
        console.error(`Error processing row ${i + 1}:`, error);
        stats.errors++;
      }
    }

    console.log('Reminder processing complete:', stats);

  } catch (error) {
    console.error('Fatal error in reminder processing:', error);
    // Send admin alert
    sendAdminAlert('Reminder Processing Failed', error);
  }
}

/**
 * Processes reminders for items in Ordered status
 * -------------------------------------------
 * Implements halving interval reminder system
 * 
 * @param {Array} row - Row data from Master Log
 * @param {number} rowIndex - 1-based row index
 * @returns {boolean} Whether a reminder was sent
 */
function processOrderedReminders(row, rowIndex) {
  const poNumber = row[COL.PO_NUMBER];
  const reminderData = getReminderData(poNumber);
  
  try {
    // Determine if reminder is needed
    const { shouldSend, interval } = calculateReminderInterval(reminderData);
    if (!shouldSend) return false;

    // Get blocking item for notification
    const blockingItem = getDetailedBlockingItem(row);
    if (!blockingItem) return false;

    // Send reminder
    const success = sendReminderNotification({
      poNumber: poNumber,
      description: row[COL.DESCRIPTION],
      blockingItem: blockingItem,
      interval: interval,
      recipients: getReminderRecipients(row, blockingItem)
    });

    // Update reminder log if sent successfully
    if (success) {
      updateReminderLog(poNumber, interval, blockingItem);
      return true;
    }

    return false;

  } catch (error) {
    console.error(`Error processing reminders for ${poNumber}:`, error);
    return false;
  }
}

/**
 * Gets reminder data from Reminders Log
 * ---------------------------------
 * 
 * @param {string} poNumber - PO number to look up
 * @returns {Object} Reminder tracking data
 */
function getReminderData(poNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reminderSheet = ss.getSheetByName('Reminders Log');
    const data = reminderSheet.getDataRange().getValues();
    
    // Find latest reminder entry for this PO
    const entry = data.reverse().find(row => row[0] === poNumber);
    
    if (!entry) {
      return {
        lastReminder: null,
        count: 0,
        currentInterval: 5 // Start with 5-day interval
      };
    }

    return {
      lastReminder: entry[1], // Last reminder date
      count: entry[2],        // Reminder count
      currentInterval: entry[3] // Current interval
    };

  } catch (error) {
    console.error('Error getting reminder data:', error);
    return null;
  }
}

/**
 * Calculates if reminder should be sent and at what interval
 * ----------------------------------------------------
 * 
 * @param {Object} reminderData - Current reminder tracking data
 * @returns {Object} Decision and interval
 */
function calculateReminderInterval(reminderData) {
  if (!reminderData) return { shouldSend: false };

  const now = new Date();
  const lastReminder = reminderData.lastReminder ? new Date(reminderData.lastReminder) : null;

  // First reminder
  if (!lastReminder) {
    return {
      shouldSend: true,
      interval: 5
    };
  }

  // Calculate days since last reminder
  const daysSinceReminder = calculateBusinessDays(lastReminder, now);

  // Calculate next interval
  let nextInterval;
  if (reminderData.count <= 2) {
    // Halving interval: 5 -> 2.5 -> 1.25
    nextInterval = reminderData.currentInterval / 2;
  } else {
    // Daily after third reminder
    nextInterval = 1;
  }

  return {
    shouldSend: daysSinceReminder >= reminderData.currentInterval,
    interval: nextInterval
  };
}

/**
 * Updates reminder tracking log
 * -------------------------
 * 
 * @param {string} poNumber - PO number
 * @param {number} interval - Reminder interval used
 * @param {Object} blockingItem - Current blocking item
 */
function updateReminderLog(poNumber, interval, blockingItem) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reminderSheet = ss.getSheetByName('Reminders Log');
    
    reminderSheet.appendRow([
      poNumber,
      new Date(),
      getReminderData(poNumber)?.count + 1 || 1,
      interval,
      blockingItem.type
    ]);

  } catch (error) {
    console.error('Error updating reminder log:', error);
  }
}

/**
 * Warning and Auto-Cancellation Functions
 * ====================================
 */

/**
 * Processes customs clearance monitoring
 * ----------------------------------
 * Called as part of daily reminder processing
 * Monitors items in customs and sends appropriate warnings
 */
function processCustomsWarnings() {
  console.log('Starting customs warning processing');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);
    const data = sheet.getDataRange().getValues();
    
    // Track processing statistics
    let stats = {
      checked: 0,
      warnings: 0,
      errors: 0
    };

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Check if item requires customs clearance and is not cleared
      if (row[COL.STATUS] === 'Ordered' && 
          row[COL.CUSTOMS_REQUIRED] === 'Y' && 
          row[COL.CUSTOMS_CLEARED] !== 'Y') {
        
        stats.checked++;
        const submissionDate = row[COL.CUSTOMS_SUBMISSION_DATE];
        if (!submissionDate) continue;
        
        try {
          const daysInCustoms = calculateBusinessDays(submissionDate, new Date());
          
          // Initial 5-day warning
          if (daysInCustoms === CONFIG.CUSTOMS_WARNING_DAYS) {
            const sent = sendCustomsWarning(row, daysInCustoms, false);
            if (sent) stats.warnings++;
          } 
          // Daily warnings after 5 days
          else if (daysInCustoms > CONFIG.CUSTOMS_WARNING_DAYS) {
            const sent = sendCustomsWarning(row, daysInCustoms, true);
            if (sent) stats.warnings++;
          }
        } catch (error) {
          console.error(`Error processing customs warning for row ${i + 1}:`, error);
          stats.errors++;
        }
      }
    }

    console.log('Customs warning processing complete:', stats);

  } catch (error) {
    console.error('Fatal error in customs warning processing:', error);
    sendAdminAlert('Customs Warning Processing Failed', error);
  }
}

/**
 * Sends customs clearance warning
 * ---------------------------
 * 
 * @param {Array} row - PO data row
 * @param {number} daysInCustoms - Number of days in customs
 * @param {boolean} isRepeat - Whether this is a repeat warning
 * @returns {boolean} Success status
 */
function sendCustomsWarning(row, daysInCustoms, isRepeat) {
  const templateKey = isRepeat ? 'CUSTOMS_WARNING_REPEAT' : 'CUSTOMS_WARNING_INITIAL';
  const recipients = getCustomsWarningRecipients(row);
  
  return sendNotification(templateKey, {
    number: row[COL.PO_NUMBER],
    description: row[COL.DESCRIPTION],
    daysInCustoms: daysInCustoms,
    submissionDate: formatDate(row[COL.CUSTOMS_SUBMISSION_DATE]),
    link: generateViewLink(row[COL.PO_NUMBER])
  }, recipients);
}

/**
 * Gets recipients for customs warnings
 * --------------------------------
 * 
 * @param {Array} row - PO data row
 * @returns {Array<string>} Array of email addresses
 */
function getCustomsWarningRecipients(row) {
  return [
    CONFIG.PROCUREMENT_EMAIL,
    CONFIG.FINANCE_EMAIL,
    row[COL.APPROVER_EMAIL]
  ].filter(Boolean); // Remove any null/undefined values
}

/**
 * Processes auto-cancellation monitoring
 * ---------------------------------
 * Checks overdue items and handles warnings/cancellations
 */
function processAutoCancellation() {
  console.log('Starting auto-cancellation processing');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);
    const data = sheet.getDataRange().getValues();
    
    let stats = {
      checked: 0,
      warnings: 0,
      cancellations: 0,
      errors: 0
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[COL.STATUS] !== 'Ordered') continue;
      
      const expectedDate = row[COL.EXPECTED_LANDING_DATE];
      if (!expectedDate) continue;
      
      stats.checked++;
      try {
        const daysOverdue = calculateDaysOverdue(expectedDate);
        
        if (daysOverdue === CONFIG.AUTO_CANCEL_WARNING_DAYS) {
          // Send 30-day warning
          const sent = sendCancellationWarning(row, CONFIG.AUTO_CANCEL_DAYS - CONFIG.AUTO_CANCEL_WARNING_DAYS);
          if (sent) stats.warnings++;
        }
        else if (daysOverdue === CONFIG.AUTO_CANCEL_DAYS - 5) {
          // Send 5-day final warning
          const sent = sendCancellationWarning(row, 5);
          if (sent) stats.warnings++;
        }
        else if (daysOverdue >= CONFIG.AUTO_CANCEL_DAYS) {
          // Process auto-cancellation
          const success = processCancellation(row, i + 1);
          if (success) stats.cancellations++;
        }
      } catch (error) {
        console.error(`Error processing auto-cancellation for row ${i + 1}:`, error);
        stats.errors++;
      }
    }

    console.log('Auto-cancellation processing complete:', stats);

  } catch (error) {
    console.error('Fatal error in auto-cancellation processing:', error);
    sendAdminAlert('Auto-Cancellation Processing Failed', error);
  }
}

/**
 * Sends cancellation warning notification
 * ----------------------------------
 * 
 * @param {Array} row - PO data row
 * @param {number} daysRemaining - Days until cancellation
 * @returns {boolean} Success status
 */
function sendCancellationWarning(row, daysRemaining) {
  return sendNotification('CANCELLATION_WARNING', {
    number: row[COL.PO_NUMBER],
    description: row[COL.DESCRIPTION],
    daysOverdue: calculateDaysOverdue(row[COL.EXPECTED_LANDING_DATE]),
    daysUntilCancel: daysRemaining,
    expectedLandingDate: formatDate(row[COL.EXPECTED_LANDING_DATE]),
    updateLink: generateLandingDateUpdateLink(row[COL.PO_NUMBER]),
    link: generateViewLink(row[COL.PO_NUMBER])
  }, getAllStakeholders(row));
}

/**
 * Utility Functions
 * ===============
 */

/**
 * Sends alert to system administrators
 * --------------------------------
 * Used for critical system errors
 * 
 * @param {string} subject - Alert subject
 * @param {Error} error - Error object
 */
function sendAdminAlert(subject, error) {
  try {
    const adminEmail = CONFIG.SYSTEM_ADMIN_EMAIL;
    if (!adminEmail) return;

    GmailApp.sendEmail({
      to: adminEmail,
      subject: `SYSTEM ALERT: ${subject}`,
      body: `
Error in 1PWR Procurement System:
${error.message}

Stack Trace:
${error.stack}

Time: ${new Date().toISOString()}
      `,
      noReply: true
    });
  } catch (e) {
    console.error('Failed to send admin alert:', e);
  }
}

/**
 * Initialize notification system
 * --------------------------
 * Sets up required sheets and triggers
 * Should be run during system setup
 */
function initializeNotificationSystem() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create Notification Log if it doesn't exist
    if (!ss.getSheetByName('Notification Log')) {
      const notificationLog = ss.insertSheet('Notification Log');
      notificationLog.getRange('A1:E1').setValues([['Timestamp', 'Type', 'PR Number', 'Recipients', 'Sent By']]);
      notificationLog.setFrozenRows(1);
    }
    
    // Create Reminders Log if it doesn't exist
    if (!ss.getSheetByName('Reminders Log')) {
      const remindersLog = ss.insertSheet('Reminders Log');
      remindersLog.getRange('A1:E1').setValues([['PR Number', 'Last Reminder Date', 'Reminder Count', 'Current Interval', 'Blocking Item']]);
      remindersLog.setFrozenRows(1);
    }
    
    // Create triggers
    createNotificationTriggers();
    
    console.log('Notification system initialized successfully');
    return true;
  } catch (error) {
    console.error('Failed to initialize notification system:', error);
    return false;
  }
}