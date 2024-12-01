/**
* 1PWR Procurement System - Status Workflow Management
* ==================================================
*
* Purpose:
* -------
* This script file is responsible for managing the status transitions and business rules for purchase requisitions (PRs) and purchase orders (POs) in the 1PWR procurement system. It handles the validation of status changes, enforcement of required fields, and the execution of associated actions (such as sending notifications, updating reminders, and triggering auto-cancellation).
*
* Integration Points:
* ------------------
* - Code.gs: Provides access to the overall system configuration, column mappings, and utility functions.
* - NotificationSystem.gs: Handles the sending of email notifications for status changes and other events.
* - DashboardWeb.html: Communicates with this script to update the status of PRs and POs displayed in the dashboard.
* - PRViewWeb.html: Allows users to update the status of individual PRs through this script.
*
* Data Framework:
* --------------
* The status workflow management relies on the "Master Log" sheet in the 1PWR procurement tracking spreadsheet. It uses the column mappings defined in the Code.gs file to access the necessary data, such as current status, required fields, and related information.
*
* Core Functionality:
* -----------------
* 1. Status Transitions: Defines the valid status transitions for both PRs and POs, enforcing the allowed transitions based on the current status.
* 2. Required Field Validation: Ensures that all necessary fields are populated before allowing a status change to be processed.
* 3. Business Rule Enforcement: Implements additional business rules, such as quotes and adjudication requirements based on the PR amount and vendor status.
* 4. Status-based Actions: Triggers associated actions when a status change occurs, such as resetting auto-cancellation timers, clearing reminders, and sending notifications.
* 5. Audit Logging: Records all status changes in the "Audit Log" sheet for historical tracking and reporting.
*
* Dependencies:
* ------------
* - Configuration constants defined in Code.gs
* - Notification system implemented in NotificationSystem.gs
* - Google Apps Script environment for sheet access and data manipulation
*
*/
/**
 * Configuration Constants
 * ======================
 * Defines valid status transitions and required fields for each status.
 * These configurations govern the entire workflow behavior.
 */
const STATUS_WORKFLOW = {
  
  // PR status transitions
  PR_TRANSITIONS: {
    'Submitted': ['In Queue', 'Revise and Resubmit', 'Rejected', 'Canceled'],
    'In Queue': ['PR Ready', 'Revise and Resubmit', 'Rejected', 'Canceled'],
    'Revise and Resubmit': ['Resubmitted', 'Canceled'],
    'Resubmitted': ['Submitted', 'Canceled'],
    'PR Ready': ['PO Pending Approval', 'Canceled'],
    'Rejected': ['Canceled'],
    'Canceled': []
  },

  // PO status transitions
  PO_TRANSITIONS: {
    'PO Pending Approval': ['PO Approved', 'Rejected', 'Canceled'],
    'PO Approved': ['PO Ordered', 'Canceled'],
    'PO Ordered': ['Completed', 'In Queue', 'Canceled'],
    'Completed': ['Canceled'],
    'Rejected': ['Canceled'],
    'Canceled': []
  },

  // Valid status transitions matrix
  // Key: current status, Value: array of allowed next statuses
  VALID_TRANSITIONS: {
    'Submitted': ['In Queue', 'Rejected', 'Revise and Resubmit', 'Canceled'],
    'In Queue': ['Ordered', 'Revise and Resubmit', 'Rejected', 'Canceled'],
    'Ordered': ['Completed', 'In Queue', 'Canceled'],
    'Revise and Resubmit': ['Resubmitted', 'Canceled'],
    'Resubmitted': ['Submitted', 'Canceled'],
    'Completed': ['Canceled'],
    'Rejected': ['Canceled'],
    'Canceled': []
  },

  // Required fields for each status
  // Maps status to required field definitions
  REQUIRED_FIELDS: {
    'Ordered': {
      fields: ['Link to PoP', 'Payment Date', 'Expected Landing Date'],
      columns: [COL.LINK_TO_POP, COL.PAYMENT_DATE, COL.EXPECTED_LANDING_DATE]
    },
    'Completed': {
      fields: ['Landed Date'],
      columns: [COL.LANDED_DATE]
    },
    'Rejected': {
      fields: ['Procurement Notes'],
      columns: [COL.PROCUREMENT_NOTES]
    },
    'Revise and Resubmit': {
      fields: ['Procurement Notes'],
      columns: [COL.PROCUREMENT_NOTES]
    }
  },

  // Status-specific automated actions
  // Maps status to array of action handlers to execute
  STATUS_ACTIONS: {
    'Ordered': ['resetAutoCancellation', 'startReminderSchedule'],
    'Completed': ['clearReminders', 'sendCompletionNotification'],
    'Canceled': ['clearReminders', 'sendCancellationNotification']
  },

  // Auto-cancellation configurations
  AUTO_CANCEL_DAYS: 30, // Days after which to auto-cancel
  AUTO_CANCEL_WARNING_DAYS: 25, // Days after which to send warning
  SYSTEM_ADMIN_EMAIL: 'admin@1pwr.com', // System admin email
  PROCUREMENT_EMAIL: 'procurement@1pwr.com',
  FINANCE_EMAIL: 'finance@1pwr.com'
};

/**
 * Status Workflow Manager Class
 * ============================
 * Core class for handling all status-related operations.
 * Provides methods for status updates, validation, authorization checks, and associated actions.
 */
class StatusWorkflowManager {
  /**
   * Initializes the workflow manager
   * Sets up sheet access and validates system configuration
   * @throws {Error} If required sheets cannot be accessed
   */
  constructor() {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      this.masterSheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);
      this.auditSheet = ss.getSheetByName('Audit Log');
      this.prTrackerSheet = ss.getSheetByName('PR Number Tracker');
      this.poTrackerSheet = ss.getSheetByName('PO Number Tracker');

      if (!this.masterSheet || !this.auditSheet ||
          !this.prTrackerSheet || !this.poTrackerSheet) {
        throw new Error('Required sheets not found');
      }

      // Cache sheet headers for performance
      this.headers = this.masterSheet.getRange(1, 1, 1,
        this.masterSheet.getLastColumn()).getValues()[0];

      // Initialize cache for queue positions
      this.queueCache = CacheService.getScriptCache();

    } catch (error) {
      console.error('StatusWorkflowManager initialization failed:', error);
      throw new Error(`Failed to initialize StatusWorkflowManager: ${error.message}`);
    }
  }

  /**
   * Updates PR/PO status with full validation, authorization checks, and associated actions
   * @param {string} docNumber - PR/PO number to update
   * @param {string} newStatus - Desired new status
   * @param {string} notes - Status change notes
   * @param {boolean} skipValidation - Optional flag to skip validation
   * @returns {Object} Update result
   */
  updateStatus(docNumber, newStatus, notes = '', skipValidation = false) {
    console.log(`Starting status update for ${docNumber} to ${newStatus}`);

    try {
      // Check authorization
      const user = this.getCurrentUser();
      if (!user || !this.isAuthorized('procurement')) {
        throw new Error('Not authorized to update status');
      }
      console.log(`User authenticated: ${user.email}`);

      const lock = LockService.getScriptLock();
      try {
        lock.waitLock(30000);

        // Find document data
        const docData = this.findDocument(docNumber);
        if (!docData) {
          throw new Error(`Document ${docNumber} not found in Master Log`);
        }

        // Determine document type and status field
        const isPR = docNumber.startsWith('PR-');
        const statusColumn = isPR ? COL.PR_STATUS : COL.PO_STATUS;
        const currentStatus = docData.rowData[statusColumn];

        // Validate status change unless skipped
        if (!skipValidation) {
          this.validateStatusTransition(currentStatus, newStatus, isPR);
          this.validateRequiredFields(docData, newStatus);
          this.validateBusinessRules(docData, newStatus);
        }

        // Prepare status change data
        const statusChangeData = {
          timestamp: new Date(),
          oldStatus: currentStatus,
          newStatus: newStatus,
          notes: notes,
          user: user.username || user.email // Assuming user object has username or email
        };

        // Perform status update
        this.performStatusUpdate(docData, statusChangeData, statusColumn);

        // Handle post-update actions
        this.handlePostUpdateActions(docData, newStatus);

        return {
          success: true,
          message: `Status updated to ${newStatus}`,
          docNumber: docNumber,
          timestamp: statusChangeData.timestamp
        };

      } finally {
        lock.releaseLock();
      }

    } catch (error) {
      console.error(`Error updating status for ${docNumber}:`, error);
      return {
        success: false,
        error: error.toString(),
        docNumber: docNumber
      };
    }
  }


  /**
   * Checks if the current user has a specific role.
   * @param {string} role - The role to check (e.g., 'procurement', 'approver', 'finance').
   * @returns {boolean} True if authorized, false otherwise.
   */
  isAuthorized(role) {
    // Implementation depends on your authorization mechanism
    const user = this.getCurrentUser();
    if (!user) return false;
    return AuthService.hasRole(user, role);
  }

  /**
   * Finds a document (PR/PO) in the Master Log by its number.
   * @param {string} docNumber - The PR or PO number to find.
   * @returns {Object|null} Document data if found, otherwise null.
   */
  findDocument(docNumber) {
    const data = this.masterSheet.getDataRange().getValues();
    const searchColumn = docNumber.startsWith('PR-') ? COL.PR_NUMBER : COL.PO_NUMBER;

    for (let i = 1; i < data.length; i++) { // Skip header row
      if (data[i][searchColumn] === docNumber) {
        return {
          rowIndex: i + 1, // Sheet rows are 1-indexed
          rowData: data[i],
          docNumber: docNumber
        };
      }
    }
    return null;
  }

  /**
   * Validates the status transition based on current and new status.
   * @param {string} currentStatus - The current status of the document.
   * @param {string} newStatus - The desired new status.
   * @param {boolean} isPR - True if the document is a PR, false if PO.
   * @throws {Error} If the transition is invalid.
   */
  validateStatusTransition(currentStatus, newStatus, isPR) {
    const transitions = isPR ? 
      STATUS_WORKFLOW.PR_TRANSITIONS : 
      STATUS_WORKFLOW.PO_TRANSITIONS;

    if (!transitions.hasOwnProperty(currentStatus)) {
      throw new Error(`Invalid current status: ${currentStatus}`);
    }

    const allowedTransitions = transitions[currentStatus];
    if (!allowedTransitions?.includes(newStatus)) {
      throw new Error(
        `Invalid status transition from ${currentStatus} to ${newStatus}. ` +
        `Allowed transitions: ${allowedTransitions?.join(', ')}`
      );
    }
  }

  /**
   * Validates that all required fields for the new status are present.
   * @param {Object} docData - The document data.
   * @param {string} newStatus - The desired new status.
   * @throws {Error} If required fields are missing.
   */
  validateRequiredFields(docData, newStatus) {
    const requirements = STATUS_WORKFLOW.REQUIRED_FIELDS[newStatus];
    if (!requirements) return;

    const missingFields = [];
    requirements.fields.forEach(field => {
      const value = docData.rowData[COL[field.toUpperCase()]];
      if (!value || (typeof value === 'string' && value.trim() === '')) {
        missingFields.push(field);
      }
    });

    // Check conditional fields
    if (requirements.conditionalFields) {
      if (this.isQuotesRequired(docData) && 
          requirements.conditionalFields.quotesRequired) {
        requirements.conditionalFields.quotesRequired.forEach(field => {
          const value = docData.rowData[COL[field.toUpperCase()]];
          if (!value || (typeof value === 'string' && value.trim() === '')) {
            missingFields.push(field);
          }
        });
      }

      if (this.isAdjudicationRequired(docData) && 
          requirements.conditionalFields.adjudicationRequired) {
        requirements.conditionalFields.adjudicationRequired.forEach(field => {
          const value = docData.rowData[COL[field.toUpperCase()]];
          if (!value || (typeof value === 'string' && value.trim() === '')) {
            missingFields.push(field);
          }
        });
      }
    }

    if (missingFields.length > 0) {
      throw new Error(
        `Missing required fields for ${newStatus} status: ${missingFields.join(', ')}`
      );
    }
  }

  /**
   * Validates additional business rules for the status transition.
   * @param {Object} docData - The document data.
   * @param {string} newStatus - The desired new status.
   * @throws {Error} If business rules are violated.
   */
  validateBusinessRules(docData, newStatus) {
    const amount = docData.rowData[COL.PR_AMOUNT];
    const isApprovedVendor = this.checkApprovedVendor(docData.rowData[COL.VENDOR]);

    if (newStatus === 'PR Ready') {
      // Quotes validation
      if (amount > 5000 && !isApprovedVendor) {
        if (!docData.rowData[COL.QUOTES_LINK]) {
          throw new Error('Three quotes are required for amounts over 5,000 LSL with non-approved vendor');
        }
      }

      if (amount > 50000) {
        if (!docData.rowData[COL.QUOTES_LINK]) {
          throw new Error('Three quotes are required for amounts over 50,000 LSL');
        }
        if (!docData.rowData[COL.ADJ_NOTES]) {
          throw new Error('Adjudication is required for amounts over 50,000 LSL');
        }
      }
    }

    if (newStatus === 'PO Ordered') {
      const expectedDate = docData.rowData[COL.EXPECTED_LANDING_DATE];
      if (expectedDate) {
        const maxDate = new Date();
        maxDate.setMonth(maxDate.getMonth() + 6);
        if (new Date(expectedDate) > maxDate) {
          throw new Error('Expected Landing Date cannot be more than 6 months in the future');
        }
      } else {
        throw new Error('Expected Landing Date is required for PO Ordered status');
      }
    }
  }
    /**
   * Performs the actual status update in the Master Log.
   * @param {Object} docData - The document data.
   * @param {Object} statusChangeData - The status change details.
   * @param {number} statusColumn - The column index for status.
   */
  performStatusUpdate(docData, statusChangeData, statusColumn) {
    try {
      const rowRange = this.masterSheet.getRange(docData.rowIndex, 1, 1,
        this.masterSheet.getLastColumn());
      const rowValues = rowRange.getValues()[0];

      // Update status
      rowValues[statusColumn] = statusChangeData.newStatus;

      // Update notes if provided
      if (statusChangeData.notes) {
        const currentNotes = rowValues[COL.PROCUREMENT_NOTES] || '';
        const updatedNotes = this.addTimestampToNote(statusChangeData.notes, currentNotes);
        rowValues[COL.PROCUREMENT_NOTES] = updatedNotes;
      }

      // Update modification tracking
      rowValues[COL.LAST_MODIFIED] = statusChangeData.timestamp;
      rowValues[COL.LAST_MODIFIED_BY] = statusChangeData.user;

      // Write all updates in single operation
      rowRange.setValues([rowValues]);

      // Log in audit trail
      this.logStatusChange(docData, statusChangeData);

    } catch (error) {
      console.error('Error performing status update:', error);
      throw new Error(`Failed to update status in sheet: ${error.message}`);
    }
  }

  /**
   * Handles post-update actions based on the new status.
   * @param {Object} docData - The document data.
   * @param {string} newStatus - The new status after update.
   */
  handlePostUpdateActions(docData, newStatus) {
    try {
      // Update queue positions if needed
      if (docData.rowData[COL.PR_STATUS] === 'In Queue' || 
          newStatus === 'In Queue') {
        this.updateQueuePositions();
      }

      // Handle PR to PO conversion
      if (docData.docNumber.startsWith('PR-') && newStatus === 'PR Ready') {
        this.initiatePOCreation(docData);
      }

      // Initialize delivery tracking for ordered items
      if (newStatus === 'PO Ordered') {
        this.initializeDeliveryTracking(docData);
      }

      // Clear reminders for completed/canceled items
      if (['Completed', 'Canceled'].includes(newStatus)) {
        this.clearAllReminders(docData.docNumber);
      }

      // Send notifications
      this.sendStatusNotification(docData, newStatus);

      // Update completion percentage
      this.updateCompletionPercentage(docData.rowIndex);

    } catch (error) {
      console.error('Error in post-update actions:', error);
      throw new Error(`Failed to complete post-update actions: ${error.message}`);
    }
  }

  /**
   * Logs the status change to the Audit Log.
   * @param {Object} docData - The document data.
   * @param {Object} statusChangeData - The status change details.
   */
  logStatusChange(docData, statusChangeData) {
    try {
      this.auditSheet.appendRow([
        statusChangeData.timestamp.toISOString(),
        statusChangeData.user,
        'Status Change',
        docData.docNumber,
        `Status changed from ${statusChangeData.oldStatus} to ${statusChangeData.newStatus}`,
        statusChangeData.oldStatus,
        statusChangeData.newStatus,
        statusChangeData.notes || ''
      ]);
    } catch (error) {
      console.error('Error logging status change:', error);
      // Non-critical error, continue execution
    }
  }

  /**
   * Adds timestamp to note
   * @param {string} newNote - New note content
   * @param {string} existingNotes - Existing notes
   * @returns {string} Combined notes with timestamp
   */
  addTimestampToNote(newNote, existingNotes) {
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm'
    );
    const formattedNote = `[${timestamp}] ${newNote}`;
    return existingNotes ? `${existingNotes}\n${formattedNote}` : formattedNote;
  }

  /**
   * Checks if quotes are required
   * @param {Object} docData - Document data
   * @returns {boolean} Whether quotes are required
   */
  isQuotesRequired(docData) {
    const amount = docData.rowData[COL.PR_AMOUNT];
    const isApprovedVendor = this.checkApprovedVendor(docData.rowData[COL.VENDOR]);
    return (amount > 5000 && !isApprovedVendor) || amount > 50000;
  }

  /**
   * Checks if adjudication is required
   * @param {Object} docData - Document data
   * @returns {boolean} Whether adjudication is required
   */
  isAdjudicationRequired(docData) {
    return docData.rowData[COL.PR_AMOUNT] > 50000;
  }

  /**
   * Checks if vendor is approved
   * @param {string} vendor - Vendor name
   * @returns {boolean} Whether vendor is approved
   */
  checkApprovedVendor(vendor) {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const vendorSheet = ss.getSheetByName('Vendor List');
      if (!vendorSheet) return false;

      const vendorData = vendorSheet.getDataRange().getValues();
      const headers = vendorData[0].map(header => header.toString().toLowerCase());
      const nameCol = headers.indexOf('vendor name');
      const statusCol = headers.indexOf('approved status (y/n)');

      if (nameCol === -1 || statusCol === -1) return false;

      // Search for vendor
      for (let i = 1; i < vendorData.length; i++) {
        if (vendorData[i][nameCol] === vendor &&
            vendorData[i][statusCol].toString().toUpperCase() === 'Y') {
          return true;
        }
      }
      return false;
    } catch (error) {
      console.error('Error checking vendor status:', error);
      return false;
    }
  }

  /**
   * Initializes delivery tracking for ordered items
   * @param {Object} docData - Document data
   * @private
   */
  initializeDeliveryTracking(docData) {
    try {
      const rowRange = this.masterSheet.getRange(docData.rowIndex, 1, 1,
        this.masterSheet.getLastColumn());
      const rowValues = rowRange.getValues()[0];

      // Initialize tracking fields
      rowValues[COL.DAYS_OPEN] = this.calculateDaysOpen(rowValues[COL.TIMESTAMP]);
      rowValues[COL.COMPLETION] = this.calculateCompletion(docData);
      
      // Set up shipping tracking if needed
      if (rowValues[COL.EXPECTED_SHIPPING_DATE]) {
        rowValues[COL.TIME_TO_SHIP] = 0;
        rowValues[COL.SHIPPED] = 'N';
      }

      // Set up customs tracking if needed
      if (rowValues[COL.CUSTOMS_REQUIRED] === 'Y') {
        rowValues[COL.CUSTOMS_CLEARED] = 'N';
        rowValues[COL.TIME_IN_CUSTOMS] = 0;
      }

      rowValues[COL.GOODS_LANDED] = 'N';
      rowValues[COL.TIME_TO_LAND] = 0;

      rowRange.setValues([rowValues]);
    } catch (error) {
      console.error('Error initializing delivery tracking:', error);
      throw new Error(`Failed to initialize delivery tracking: ${error.message}`);
    }
  }

  /**
   * Calculates completion percentage
   * @param {Object} docData - Document data
   * @returns {number} Completion percentage
   * @private
   */
  calculateCompletion(docData) {
    try {
      const row = docData.rowData;
      const isPR = docData.docNumber.startsWith('PR-');
      let requiredFields = [];

      if (isPR) {
        requiredFields = [
          COL.PR_NUMBER,
          COL.REQUESTOR_NAME,
          COL.DEPARTMENT,
          COL.DESCRIPTION,
          COL.PR_AMOUNT
        ];

        // Add conditional requirements
        if (this.isQuotesRequired(docData)) {
          requiredFields.push(COL.QUOTES_LINK, COL.QUOTES_DATE);
        }
        if (this.isAdjudicationRequired(docData)) {
          requiredFields.push(COL.ADJ_NOTES, COL.ADJ_DATE);
        }
      } else {
        requiredFields = [
          COL.PO_NUMBER,
          COL.LINK_TO_POP,
          COL.PAYMENT_DATE,
          COL.EXPECTED_LANDING_DATE
        ];

        if (row[COL.CUSTOMS_REQUIRED] === 'Y') {
          requiredFields.push(
            COL.CUSTOMS_DOCS,
            COL.CUSTOMS_SUBMISSION_DATE
          );
        }
      }

      // Count populated required fields
      const filledCount = requiredFields.filter(col => {
        const value = row[col];
        return value !== null && value !== undefined && value.toString().trim() !== '';
      }).length;

      return Math.round((filledCount / requiredFields.length) * 100);
    } catch (error) {
      console.error('Error calculating completion:', error);
      return 0;
    }
  }

  /**
   * Calculates days open
   * @param {Date} timestamp - Creation timestamp
   * @returns {number} Number of business days open
   * @private
   */
  calculateDaysOpen(timestamp) {
    if (!timestamp) return 0;

    const today = new Date();
    let count = 0;
    let current = new Date(timestamp);

    while (current <= today) {
      // Skip weekends (0 = Sunday, 6 = Saturday)
      if (current.getDay() !== 0 && current.getDay() !== 6) {
        count++;
      }
      current.setDate(current.getDate() + 1);
    }

    return count;
  }

  /**
   * Clears all reminders for a document
   * @param {string} docNumber - Document number
   * @private
   */
  clearAllReminders(docNumber) {
    try {
      // Clear reminder cache
      this.queueCache.remove(`reminders_${docNumber}`);

      // Clear from reminder log if exists
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const reminderSheet = ss.getSheetByName('Reminders Log');
      if (!reminderSheet) return;

      const data = reminderSheet.getDataRange().getValues();
      const docCol = 0; // Assuming document number is in first column

      // Find and delete all reminder rows for this document
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][docCol] === docNumber) {
          reminderSheet.deleteRow(i + 1);
        }
      }
    } catch (error) {
      console.error('Error clearing reminders:', error);
      // Non-critical error, continue execution
    }
  }

  /**
   * Sends status change notification
   * @param {Object} docData - Document data
   * @param {string} newStatus - New status
   * @private
   */
  sendStatusNotification(docData, newStatus) {
    try {
      const notificationData = {
        docNumber: docData.docNumber,
        description: docData.rowData[COL.DESCRIPTION],
        requestor: docData.rowData[COL.REQUESTOR_NAME],
        department: docData.rowData[COL.DEPARTMENT],
        status: newStatus,
        previousStatus: docData.rowData[docData.docNumber.startsWith('PR-') ? 
          COL.PR_STATUS : COL.PO_STATUS],
        link: this.generateDocumentLink(docData.docNumber)
      };

      const recipients = [
        CONFIG.PROCUREMENT_EMAIL,
        docData.rowData[COL.EMAIL],
        docData.rowData[COL.APPROVER]
      ].filter(email => email);

      NotificationSystem.sendNotification(
        'STATUS_CHANGE',
        notificationData,
        recipients
      );
    } catch (error) {
      console.error('Error sending notification:', error);
      // Non-critical error, continue execution
    }
  }

  /**
   * Generates document link
   * @param {string} docNumber - Document number
   * @returns {string} Document view URL
   * @private
   */
  generateDocumentLink(docNumber) {
    const baseUrl = ScriptApp.getService().getUrl();
    return `${baseUrl}?doc=${encodeURIComponent(docNumber)}`;
  }

  /**
   * Updates completion percentage for a specific row
   * @param {number} rowIndex - Row index in the sheet
   * @private
   */
  updateCompletionPercentage(rowIndex) {
    try {
      const sheet = this.masterSheet;
      const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
      const docNumber = row[COL.PR_NUMBER] || row[COL.PO_NUMBER];
      const docData = {
        rowIndex: rowIndex,
        rowData: row,
        docNumber: docNumber
      };
      const completion = this.calculateCompletion(docData);
      sheet.getRange(rowIndex, COL.COMPLETION + 1).setValue(completion);
    } catch (error) {
      console.error(`Error updating completion percentage for row ${rowIndex}:`, error);
      throw new Error(`Failed to update completion percentage: ${error.message}`);
    }
  }

  /**
   * Initiates PO creation from a PR
   * @param {Object} docData - Document data
   * @private
   */
  initiatePOCreation(docData) {
    try {
      // Implement PO creation logic here
      // Example: Generate PO number, create new PO entry, link to PR
      const newPONumber = `PO-${Date.now()}`;
      this.poTrackerSheet.appendRow([newPONumber, docData.docNumber, new Date().toISOString()]);
      console.log(`Created new PO ${newPONumber} from PR ${docData.docNumber}`);
    } catch (error) {
      console.error('Error initiating PO creation:', error);
      throw new Error(`Failed to initiate PO creation: ${error.message}`);
    }
  }

  /**
   * Initializes delivery tracking for a PO
   * @param {Object} docData - Document data
   * @private
   */
  initializeDeliveryTracking(docData) {
    try {
      const rowRange = this.masterSheet.getRange(docData.rowIndex, 1, 1,
        this.masterSheet.getLastColumn());
      const rowValues = rowRange.getValues()[0];

      // Initialize tracking fields
      rowValues[COL.DAYS_OPEN] = this.calculateDaysOpen(rowValues[COL.TIMESTAMP]);
      rowValues[COL.COMPLETION] = this.calculateCompletion(docData);
      
      // Set up shipping tracking if needed
      if (rowValues[COL.EXPECTED_SHIPPING_DATE]) {
        rowValues[COL.TIME_TO_SHIP] = 0;
        rowValues[COL.SHIPPED] = 'N';
      }

      // Set up customs tracking if needed
      if (rowValues[COL.CUSTOMS_REQUIRED] === 'Y') {
        rowValues[COL.CUSTOMS_CLEARED] = 'N';
        rowValues[COL.TIME_IN_CUSTOMS] = 0;
      }

      rowValues[COL.GOODS_LANDED] = 'N';
      rowValues[COL.TIME_TO_LAND] = 0;

      rowRange.setValues([rowValues]);
    } catch (error) {
      console.error('Error initializing delivery tracking:', error);
      throw new Error(`Failed to initialize delivery tracking: ${error.message}`);
    }
  }
} //end class

/**
 * Factory Functions and Global Handlers
 * ===================================
 */

/**
 * Factory function for StatusWorkflowManager
 * ----------------------------------------
 * Provides single entry point for other scripts to access workflow manager
 * 
 * @returns {StatusWorkflowManager} New instance of workflow manager
 * 
 * Usage:
 * const workflow = getStatusWorkflowManager();
 * workflow.updateStatus(...);
 */
function getStatusWorkflowManager() {
  return new StatusWorkflowManager();
}

/**
 * Automated Processing Functions
 * ============================
 */

/**
 * Initializes triggers for automated status checks
 * ---------------------------------------------
 * Should be run once during system setup
 * Creates daily and hourly triggers for status monitoring
 */
function createStatusWorkflowTriggers() {
  try {
    // Remove existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction().startsWith('processStatus')) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Create daily trigger for status checks
    ScriptApp.newTrigger('processStatusChecks')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();

    // Create hourly trigger for urgent checks
    ScriptApp.newTrigger('processUrgentStatusChecks')
      .timeBased()
      .everyHours(1)
      .create();

    console.log('Status workflow triggers created successfully');
  } catch (error) {
    console.error('Error creating status workflow triggers:', error);
    throw new Error(`Failed to create status workflow triggers: ${error.message}`);
  }
}

/**
 * Daily automated status checks
 * --------------------------
 * Handles overdue items and reminder scheduling
 * Triggered daily at 1 AM
 */
function processStatusChecks() {
  console.log('Starting daily status checks');
  const workflow = new StatusWorkflowManager();
  
  try {
    // Process ordered items
    processOrderedItems(workflow);
    
    // Check for auto-cancellation
    processAutoCancellation(workflow);
    
    // Update completion percentages
    updateAllCompletionPercentages(workflow);
    
    console.log('Daily status checks completed successfully');
  } catch (error) {
    console.error('Error in daily status checks:', error);
    // Send alert to system administrators
    NotificationSystem.sendNotification(
      'SYSTEM_ALERT',
      {
        type: 'Error',
        process: 'Daily Status Checks',
        error: error.message,
        timestamp: new Date()
      },
      [STATUS_WORKFLOW.SYSTEM_ADMIN_EMAIL]
    );
  }
}

/**
 * Hourly urgent status checks
 * ------------------------
 * Processes time-sensitive status updates
 * Triggered every hour
 */
function processUrgentStatusChecks() {
  console.log('Starting urgent status checks');
  const workflow = new StatusWorkflowManager();

  try {
    // Check overdue ordered items
    checkOverdueOrders(workflow);
    
    // Process urgent reminders
    processUrgentReminders(workflow);
    
    console.log('Urgent status checks completed successfully');
  } catch (error) {
    console.error('Error in urgent status checks:', error);
    // Send alert to system administrators
    NotificationSystem.sendNotification(
      'SYSTEM_ALERT',
      {
        type: 'Error',
        process: 'Urgent Status Checks',
        error: error.message,
        timestamp: new Date()
      },
      [STATUS_WORKFLOW.SYSTEM_ADMIN_EMAIL]
    );
  }
}

/**
 * Status Processing Functions
 * =========================
 * Implements specific business logic for status-based processing
 */

/**
 * Processes items in Ordered status
 * ------------------------------
 * Handles reminder scheduling and overdue checks for ordered items
 * 
 * @param {StatusWorkflowManager} workflow - Workflow manager instance
 */
function processOrderedItems(workflow) {
  try {
    const sheet = workflow.masterSheet;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Find ordered items
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[COL.STATUS] !== 'Ordered') continue;

      const poNumber = row[COL.PO_NUMBER];
      console.log(`Processing ordered item: ${poNumber}`);

      try {
        // Check if landing date is set
        const expectedLandingDate = row[COL.EXPECTED_LANDING_DATE];
        if (!expectedLandingDate) {
          console.warn(`No landing date for ${poNumber}`);
          continue;
        }

        // Calculate days overdue
        const daysOverdue = calculateDaysOverdue(expectedLandingDate);
        
        if (daysOverdue > 0) {
          handleOverdueOrder({
            poNumber: poNumber,
            daysOverdue: daysOverdue,
            rowIndex: i + 1,
            workflow: workflow,
            poData: {
              submitter: row[COL.REQUESTOR_NAME],
              submitterEmail: row[COL.EMAIL],
              approver: row[COL.APPROVER],
              approverEmail: row[COL.APPROVER_EMAIL],
              description: row[COL.DESCRIPTION]
            }
          });
        }

        // Update reminder schedule
        updateReminderSchedule(poNumber, daysOverdue);

      } catch (itemError) {
        console.error(`Error processing ordered item ${poNumber}:`, itemError);
        // Continue with next item
      }
    }
  } catch (error) {
    console.error('Error in processOrderedItems:', error);
    throw new Error(`Failed to process ordered items: ${error.message}`);
  }
}

/**
 * Handles overdue ordered items
 * --------------------------
 * Manages notifications and auto-cancellation for overdue items
 * 
 * @param {Object} params - Processing parameters
 * @param {string} params.poNumber - PO number
 * @param {number} params.daysOverdue - Number of days overdue
 * @param {number} params.rowIndex - Sheet row index
 * @param {StatusWorkflowManager} params.workflow - Workflow manager
 * @param {Object} params.poData - PO data for notifications
 */
function handleOverdueOrder(params) {
  const { poNumber, daysOverdue, rowIndex, workflow, poData } = params;

  try {
    // Check auto-cancellation threshold
    if (daysOverdue >= STATUS_WORKFLOW.AUTO_CANCEL_DAYS) {
      console.log(`Auto-canceling PO ${poNumber} - ${daysOverdue} days overdue`);
      
      // Update status to Canceled
      workflow.updateStatus(
        poNumber, 
        'Canceled',
        `Automatically canceled after ${daysOverdue} days overdue`,
        true // Skip validation
      );

    } else if (daysOverdue >= STATUS_WORKFLOW.AUTO_CANCEL_WARNING_DAYS) {
      // Send warning notification
      const daysUntilCancel = STATUS_WORKFLOW.AUTO_CANCEL_DAYS - daysOverdue;
      
      NotificationSystem.sendNotification(
        'CANCELLATION_WARNING',
        {
          poNumber: poNumber,
          description: poData.description,
          daysOverdue: daysOverdue,
          daysUntilCancel: daysUntilCancel,
          expectedLandingDate: workflow.findDocument(poNumber).rowData[COL.EXPECTED_LANDING_DATE],
          updateLink: generateLandingDateUpdateLink(poNumber)
        },
        [
          poData.submitterEmail,
          poData.approverEmail,
          STATUS_WORKFLOW.PROCUREMENT_EMAIL,
          STATUS_WORKFLOW.FINANCE_EMAIL
        ]
      );
    }
  } catch (error) {
    console.error(`Error handling overdue order ${poNumber}:`, error);
    throw new Error(`Failed to handle overdue order: ${error.message}`);
  }
}

/**
 * Processes auto-cancellation checks
 * ------------------------------
 * Identifies and handles items requiring auto-cancellation
 * 
 * @param {StatusWorkflowManager} workflow - Workflow manager instance
 */
function processAutoCancellation(workflow) {
  try {
    const sheet = workflow.masterSheet;
    const data = sheet.getDataRange().getValues();

    // Find items in Ordered status
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[COL.STATUS] !== 'Ordered') continue;

      const poNumber = row[COL.PO_NUMBER];
      console.log(`Checking auto-cancellation for ${poNumber}`);

      try {
        const expectedLandingDate = row[COL.EXPECTED_LANDING_DATE];
        if (!expectedLandingDate) continue;

        const daysOverdue = calculateDaysOverdue(expectedLandingDate);
        if (daysOverdue >= STATUS_WORKFLOW.AUTO_CANCEL_DAYS) {
          // Auto-cancel the PO
          workflow.updateStatus(
            poNumber,
            'Canceled',
            `Automatically canceled after ${daysOverdue} days overdue`,
            true // Skip validation
          );
        }
      } catch (itemError) {
        console.error(`Error processing auto-cancellation for ${poNumber}:`, itemError);
        // Continue with next item
      }
    }
  } catch (error) {
    console.error('Error in processAutoCancellation:', error);
    throw new Error(`Failed to process auto-cancellation: ${error.message}`);
  }
}

/**
 * Updates completion percentages for all active POs
 * --------------------------------------------
 * Recalculates completion percentages for reporting
 * 
 * @param {StatusWorkflowManager} workflow - Workflow manager instance
 */
function updateAllCompletionPercentages(workflow) {
  try {
    const sheet = workflow.masterSheet;
    const data = sheet.getDataRange().getValues();

    // Process all active POs
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[COL.STATUS];

      // Skip completed/canceled POs
      if (['Completed', 'Canceled', 'Rejected'].includes(status)) continue;

      try {
        workflow.updateCompletionPercentage(i + 1);
      } catch (itemError) {
        console.error(`Error updating completion for row ${i + 1}:`, itemError);
        // Continue with next item
      }
    }
  } catch (error) {
    console.error('Error in updateAllCompletionPercentages:', error);
    throw new Error(`Failed to update completion percentages: ${error.message}`);
  }
}

/**
 * Utility Functions
 * ===============
 */

/**
 * Calculates days overdue for a landing date
 * --------------------------------------
 * 
 * @param {Date} expectedLandingDate - Expected landing date
 * @returns {number} Number of business days overdue (negative if not overdue)
 */
function calculateDaysOverdue(expectedLandingDate) {
  if (!expectedLandingDate) return 0;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const expected = new Date(expectedLandingDate);
  expected.setHours(0, 0, 0, 0);
  
  if (expected >= today) return 0;
  
  return calculateBusinessDays(expected, today);
}

/**
 * Calculates business days between dates
 * ----------------------------------
 * 
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {number} Number of business days
 */
function calculateBusinessDays(startDate, endDate) {
  let count = 0;
  const current = new Date(startDate);
  
  while (current <= endDate) {
    // Skip weekends (0 = Sunday, 6 = Saturday)
    if (current.getDay() !== 0 && current.getDay() !== 6) {
      count++;
    }
    current.setDate(current.getDate() + 1);
  }
  
  return count;
}

/**
 * Generates a landing date update link
 * @param {string} poNumber - PO number
 * @returns {string} Landing date update URL
 * @private
 */
function generateLandingDateUpdateLink(poNumber) {
  const baseUrl = ScriptApp.getService().getUrl();
  return `${baseUrl}?action=updateLandingDate&doc=${encodeURIComponent(poNumber)}`;
}
