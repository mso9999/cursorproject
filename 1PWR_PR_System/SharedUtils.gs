/**
 * SharedUtils.gs (Refactored)
 * Part 1 of 3: Core Configuration and Session Management
 * Last Updated: December 1, 2024 22:00 GMT+2
 * 
 * This file contains consolidated utility functions for the PR system.
 * Major changes:
 * 1. Consolidated duplicate session management functions
 * 2. Unified configuration objects
 * 3. Streamlined authentication logic
 * 4. Enhanced error handling and logging
 * 
 * Integration Notes:
 * - Previous getCurrentUserShared() calls should now use getCurrentUser()
 * - Previous validateSheetSession() calls should now use validateSession()
 * - Previous isAuthorizedShared() calls should now use isAuthorized()
 */


/**
 * Session management configuration
 * Consolidated from multiple session configs
 */
const SESSION_CONFIG = {
  // Sheet configuration
  SHEET_NAME: 'Session Log',
  CLEANUP_FREQUENCY: 60, // minutes
  
  // Timeouts
  INACTIVITY_TIMEOUT: 30, // minutes
  SESSION_DURATION: 1440, // 24 hours
  
  // Column indices
  COLUMNS: {
    SESSION_ID: 0,
    TIMESTAMP: 1,
    USERNAME: 2,
    EMAIL: 3,
    ROLE: 4,
    LAST_ACTIVITY: 5,
    STATUS: 6
  },
  
  // Status values
  STATUS: {
    ACTIVE: 'Active',
    EXPIRED: 'Expired'
  }
};

/**
 * Role hierarchy configuration
 * Consolidated from multiple auth checks
 */
const ROLE_HIERARCHY = {
  'procurement': ['procurement'],
  'finance': ['procurement', 'finance'],
  'approver': ['procurement', 'approver'],
  'requestor': ['procurement', 'finance', 'approver', 'requestor']
};





/**
 * Checks if user is authorized for a role - Consolidated from isAuthorizedShared
 * @param {Object} user - User object with email and optional username
 * @param {string} requiredRole - Role to check
 * @returns {boolean} Whether user is authorized
 */
function isAuthorized(user, requiredRole) {
  if (!user || !requiredRole) return false;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.REQUESTOR_SHEET_NAME);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    
    // Find user row
    const userRow = data.find(row =>
      (row[1].toString().trim().toLowerCase() === user.email.toLowerCase()) ||
      (row[0].toString().trim().toLowerCase() === user.username?.toLowerCase())
    );

    if (!userRow || userRow[3].toString().toUpperCase() !== 'Y') {
      return false;
    }

    const userRole = userRow[5].toString().trim().toLowerCase();
    return ROLE_HIERARCHY[requiredRole]?.includes(userRole) || false;

  } catch (error) {
    console.error('Error in isAuthorized:', error);
    return false;
  }
}

/**
 * SharedUtils.gs (Refactored)
 * Part 2 of 3: Utility Functions and Error Handling
 * Last Updated: 2024-11-21
 * 
 * This section contains utility functions for:
 * - Error handling and display
 * - Date formatting and calculations
 * - Loading state management
 * - URL parameter handling
 * - Input validation
 * 
 * Integration Notes:
 * - Previous showFieldError() calls remain unchanged
 * - Previous calculateBusinessDays() calls remain unchanged
 * - formatDate() and formatDateTime() have been enhanced with better error handling
 */

/**
 * UI Error Display Configuration
 */
const ERROR_DISPLAY = {
  TYPES: {
    ERROR: 'error',
    WARNING: 'warning',
    INFO: 'info'
  },
  CLASSES: {
    ERROR: 'invalid-input',
    MESSAGE: 'error-message',
    SYSTEM: 'system-message',
    HIDDEN: 'hidden',
    SHAKE: 'shake'
  }
};

/**
 * Shows error message for specific form field
 * @param {string} fieldId - ID of form field
 * @param {string} message - Error message
 */
function showFieldError(fieldId, message) {
  const field = document.getElementById(fieldId);
  if (!field) return;

  // Add error styling to field
  field.classList.add(ERROR_DISPLAY.CLASSES.ERROR);
  field.setAttribute('aria-invalid', 'true');

  // Show error message
  const errorSpan = document.getElementById(`${fieldId}Error`);
  if (errorSpan) {
    errorSpan.textContent = message;
    errorSpan.classList.remove(ERROR_DISPLAY.CLASSES.HIDDEN);
  }

  // Add shake animation
  field.classList.add(ERROR_DISPLAY.CLASSES.SHAKE);
  setTimeout(() => field.classList.remove(ERROR_DISPLAY.CLASSES.SHAKE), 500);
}

/**
 * Clears all error messages and styling
 */
function clearErrors() {
  // Clear field errors
  document.querySelectorAll(`.${ERROR_DISPLAY.CLASSES.ERROR}`).forEach(field => {
    field.classList.remove(ERROR_DISPLAY.CLASSES.ERROR);
    field.removeAttribute('aria-invalid');
  });

  // Hide error messages
  document.querySelectorAll(`.${ERROR_DISPLAY.CLASSES.MESSAGE}`).forEach(message => {
    message.classList.add(ERROR_DISPLAY.CLASSES.HIDDEN);
    message.textContent = '';
  });

  // Remove system messages
  document.querySelectorAll(`.${ERROR_DISPLAY.CLASSES.SYSTEM}`).forEach(message => {
    message.remove();
  });
}

/**
 * Date formatting configuration
 */
const DATE_FORMAT = {
  OPTIONS: {
    short: { year: 'numeric', month: 'short', day: 'numeric' },
    long: { year: 'numeric', month: 'long', day: 'numeric' },
    iso: { year: 'numeric', month: '2-digit', day: '2-digit' }
  },
  TIME_OPTIONS: {
    standard: {
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    },
    detailed: {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true
    }
  }
};

/**
 * Formats date for display
 * @param {Date|string} date - Date to format
 * @param {string} [format='short'] - Format style (short, long, iso)
 * @returns {string} Formatted date string
 */
function formatDate(date, format = 'short') {
  if (!date) return '';

  try {
    const dateObj = new Date(date);
    if (isNaN(dateObj.getTime())) {
      console.warn('Invalid date input:', date);
      return '';
    }

    if (format === 'iso') {
      return dateObj.toISOString().split('T')[0];
    }

    return dateObj.toLocaleDateString(undefined, DATE_FORMAT.OPTIONS[format] || DATE_FORMAT.OPTIONS.short);
  } catch (error) {
    console.error('Date formatting error:', error);
    return date.toString();
  }
}

/**
 * Formats date and time for display
 * @param {Date|string} datetime - Date/time to format
 * @param {boolean} [includeSeconds=false] - Whether to include seconds
 * @returns {string} Formatted date/time string
 */
function formatDateTime(datetime, includeSeconds = false) {
  if (!datetime) return '';

  try {
    const dateObj = new Date(datetime);
    if (isNaN(dateObj.getTime())) {
      console.warn('Invalid datetime input:', datetime);
      return '';
    }

    const timeOptions = includeSeconds ? DATE_FORMAT.TIME_OPTIONS.detailed : DATE_FORMAT.TIME_OPTIONS.standard;
    const datePart = dateObj.toLocaleDateString(undefined, DATE_FORMAT.OPTIONS.short);
    const timePart = dateObj.toLocaleTimeString(undefined, timeOptions);

    return `${datePart} at ${timePart}`;
  } catch (error) {
    console.error('DateTime formatting error:', error);
    return datetime.toString();
  }
}

/**
 * Calculates business days between two dates
 * @param {Date|string} startDate - Start date
 * @param {Date|string} endDate - End date
 * @returns {number} Number of business days
 */
function calculateBusinessDays(startDate, endDate) {
  try {
    const start = new Date(startDate);
    const end = new Date(endDate);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      console.warn('Invalid date input for business days calculation');
      return 0;
    }

    let count = 0;
    let current = new Date(start);

    while (current <= end) {
      // Skip weekends (0 = Sunday, 6 = Saturday)
      if (current.getDay() !== 0 && current.getDay() !== 6) {
        count++;
      }
      current.setDate(current.getDate() + 1);
    }

    return count;
  } catch (error) {
    console.error('Error calculating business days:', error);
    return 0;
  }
}

/**
 * Loading state management functions
 */
const LOADING_STATE = {
  OVERLAY_CLASS: 'loading-overlay',
  SPINNER_CLASS: 'loading-spinner',
  LOADING_CLASS: 'loading'
};

/**
 * Shows loading state for element
 * @param {string} elementId - ID of element to show loading state for
 */
function showLoading(elementId) {
  const element = document.getElementById(elementId);
  if (!element) return;

  element.classList.add(LOADING_STATE.LOADING_CLASS);

  // Add loading overlay
  const overlay = document.createElement('div');
  overlay.className = `${LOADING_STATE.OVERLAY_CLASS} absolute inset-0 bg-white bg-opacity-75 flex items-center justify-center`;
  overlay.innerHTML = `<div class="${LOADING_STATE.SPINNER_CLASS}"></div>`;

  element.style.position = 'relative';
  element.appendChild(overlay);
}

/**
 * Hides loading state for element
 * @param {string} elementId - ID of element to hide loading state for
 */
function hideLoading(elementId) {
  const element = document.getElementById(elementId);
  if (!element) return;

  element.classList.remove(LOADING_STATE.LOADING_CLASS);

  // Remove loading overlay
  const overlay = element.querySelector(`.${LOADING_STATE.OVERLAY_CLASS}`);
  if (overlay) {
    overlay.remove();
  }
}

/**
 * URL parameter handling functions
 */

/**
 * Safely gets URL parameter
 * @param {string} param - Parameter name
 * @returns {string|null} Parameter value
 */
function getUrlParam(param) {
  try {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
  } catch (error) {
    console.error('Error getting URL parameter:', error);
    return null;
  }
}

/**
 * Validates email address format
 * @param {string} email - Email address to validate
 * @returns {boolean} Whether email is valid
 */
function isValidEmail(email) {
  if (!email) return false;
  
  try {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(String(email).toLowerCase());
  } catch (error) {
    console.error('Email validation error:', error);
    return false;
  }
}

/**
 * Handles server errors
 * @param {Error|string} error - Error object or message
 * @param {string} [context] - Error context
 */
function handleServerError(error, context = '') {
  const errorMessage = error.message || error.toString();
  const contextMessage = context ? ` in ${context}` : '';
  
  console.error(`Server error${contextMessage}:`, error);
  showError(`An error occurred${context ? ` while ${context}` : ''}: ${errorMessage}`);
  
  // Log to server if available
  try {
    if (typeof logAuthEvent === 'function') {
      logAuthEvent(
        getCurrentUser()?.username || 'Unknown',
        false,
        `Server error${contextMessage}: ${errorMessage}`
      );
    }
  } catch (loggingError) {
    console.error('Error logging server error:', loggingError);
  }
}

/**
 * SharedUtils.gs (Refactored)
 * Part 3 of 3: PR Utilities and Session Logging
 * Last Updated: 2024-11-21
 * 
 * This section contains:
 * - PR retrieval and formatting functions
 * - Session logging and management
 * - Audit log functionality
 * - Authentication utilities
 * 
 * Integration Notes:
 * - Previous getPRDetailsShared() calls should now use getPRDetails()
 * - Previous getAuditHistoryShared() calls should now use getAuditHistory()
 * - All session logging now uses enhanced error handling and validation
 */

/**
 * PR Status and Management Configuration
 */
const PR_CONFIG = {
  STATUS_CLASSES: {
    OVERDUE: 'overdue',
    WARNING: 'warning',
    CURRENT: 'current'
  },
  WARNING_THRESHOLD: 30, // days for warning status
  AUDIT_COLUMNS: {
    TIMESTAMP: 0,
    USER: 1,
    ACTION: 2,
    PR_NUMBER: 3,
    DETAILS: 4,
    OLD_STATUS: 5,
    NEW_STATUS: 6
  }
};

/**
 * Gets detailed PR information from Master Log sheet
 * @param {string} prNumber - PR number to retrieve
 * @returns {Object} PR details and success status
 */
function getPRDetails(prNumber) {
  console.log('Getting PR details for:', prNumber);
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);

    if (!sheet) {
      throw new Error('Master Log sheet not found');
    }

    // Find PR row
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[CONFIG.COLUMNS.PR_NUMBER] === prNumber);

    if (rowIndex === -1) {
      throw new Error(`PR ${prNumber} not found`);
    }

    const row = data[rowIndex];

    // Log the raw line item data
    console.log('Raw line item data:', {
      items: row[CONFIG.COLUMNS.ITEM_LIST],
      quantities: row[CONFIG.COLUMNS.QTY_LIST],
      uoms: row[CONFIG.COLUMNS.UOM_LIST],
      urls: row[CONFIG.COLUMNS.URL_LIST]
    });

    // Parse line items and get history
    const lineItems = parseLineItems(row);
    const history = getAuditHistory(prNumber);

    return {
      success: true,
      data: formatPRData(row, lineItems, history, rowIndex)
    };

  } catch (error) {
    console.error('Error getting PR details:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Formats PR data for response
 * @param {Array} row - Raw data row
 * @param {Array} lineItems - Parsed line items
 * @param {Array} history - Audit history
 * @param {number} rowIndex - Row index in sheet
 * @returns {Object} Formatted PR data
 */
function formatPRData(row, lineItems, history, rowIndex) {
  return {
    // Basic information
    prNumber: row[CONFIG.COLUMNS.PR_NUMBER],
    timestamp: row[CONFIG.COLUMNS.TIMESTAMP],
    status: row[CONFIG.COLUMNS.PR_STATUS],
    requestorName: row[CONFIG.COLUMNS.REQUESTOR_NAME],
    department: row[CONFIG.COLUMNS.DEPARTMENT],
    requestorEmail: row[CONFIG.COLUMNS.EMAIL],
    description: row[CONFIG.COLUMNS.DESCRIPTION],
    
    // Project details
    projectCategory: row[CONFIG.COLUMNS.PROJECT_CATEGORY],
    organization: row[CONFIG.COLUMNS.ORGANIZATION],
    currency: row[CONFIG.COLUMNS.CURRENCY],
    paymentFormat: row[CONFIG.COLUMNS.PAYMENT_FORMAT],
    
    // Location and type
    siteLocation: row[CONFIG.COLUMNS.SITE_LOCATION],
    expenseType: row[CONFIG.COLUMNS.EXPENSE_TYPE],
    vehicle: row[CONFIG.COLUMNS.VEHICLE],
    
    // Budget and timeline
    budgetStatus: row[CONFIG.COLUMNS.BUDGET_STATUS],
    deadlineDate: row[CONFIG.COLUMNS.DEADLINE_DATE],
    vendor: row[CONFIG.COLUMNS.VENDOR],
    urgencyStatus: row[CONFIG.COLUMNS.URGENCY],
    approver: row[CONFIG.COLUMNS.APPROVER],
    
    // Processing details
    requestorNotes: row[CONFIG.COLUMNS.REQUESTOR_NOTES],
    procurementNotes: row[CONFIG.COLUMNS.PROCUREMENT_NOTES],
    
    // PO details
    poNumber: row[CONFIG.COLUMNS.PO_NUMBER],
    poDate: row[CONFIG.COLUMNS.PO_DATE],
    poStatus: row[CONFIG.COLUMNS.PO_STATUS],
    poApprovedDate: row[CONFIG.COLUMNS.PO_APPROVED_DATE],
    
    // Payment and shipping
    paymentDate: row[CONFIG.COLUMNS.PAYMENT_DATE],
    expectedShippingDate: row[CONFIG.COLUMNS.EXPECTED_SHIPPING_DATE],
    expectedLandingDate: row[CONFIG.COLUMNS.EXPECTED_LANDING_DATE],
    shipped: row[CONFIG.COLUMNS.SHIPPED],
    shipmentDate: row[CONFIG.COLUMNS.SHIPMENT_DATE],
    
    // Customs and delivery
    customsRequired: row[CONFIG.COLUMNS.CUSTOMS_REQUIRED],
    customsDocs: row[CONFIG.COLUMNS.CUSTOMS_DOCS],
    customsSubmissionDate: row[CONFIG.COLUMNS.CUSTOMS_SUBMISSION_DATE],
    customsCleared: row[CONFIG.COLUMNS.CUSTOMS_CLEARED],
    dateCleared: row[CONFIG.COLUMNS.DATE_CLEARED],
    goodsLanded: row[CONFIG.COLUMNS.GOODS_LANDED],
    landedDate: row[CONFIG.COLUMNS.LANDED_DATE],
    
    // Tracking
    daysOpen: calculateDaysOpen(row[CONFIG.COLUMNS.TIMESTAMP]),
    completion: row[CONFIG.COLUMNS.COMPLETION],
    timeToShip: row[CONFIG.COLUMNS.TIME_TO_SHIP],
    timeInCustoms: row[CONFIG.COLUMNS.TIME_IN_CUSTOMS],
    timeToLand: row[CONFIG.COLUMNS.TIME_TO_LAND],
    
    // Line items and history
    lineItems: lineItems,
    history: history,
    
    // Debug info
    _debug: {
      rawLineItems: {
        items: row[CONFIG.COLUMNS.ITEM_LIST],
        quantities: row[CONFIG.COLUMNS.QTY_LIST],
        uoms: row[CONFIG.COLUMNS.UOM_LIST],
        urls: row[CONFIG.COLUMNS.URL_LIST]
      }
    },
    
    // Metadata
    rowIndex: rowIndex + 1,
    lastModified: row[CONFIG.COLUMNS.LAST_MODIFIED],
    lastModifiedBy: row[CONFIG.COLUMNS.LAST_MODIFIED_BY]
  };
}

/**
 * Gets PR audit history
 * @param {string} prNumber - PR number
 * @returns {Array} History entries
 */
function getAuditHistory(prNumber) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const auditSheet = ss.getSheetByName('Audit Log');

    if (!auditSheet) {
      console.warn('Audit Log sheet not found');
      return [];
    }

    const data = auditSheet.getDataRange().getValues();
    
    return data
      .slice(1) // Skip header
      .filter(row => row[PR_CONFIG.AUDIT_COLUMNS.PR_NUMBER] === prNumber)
      .map(row => ({
        timestamp: row[PR_CONFIG.AUDIT_COLUMNS.TIMESTAMP],
        user: row[PR_CONFIG.AUDIT_COLUMNS.USER],
        action: row[PR_CONFIG.AUDIT_COLUMNS.ACTION],
        details: row[PR_CONFIG.AUDIT_COLUMNS.DETAILS],
        oldStatus: row[PR_CONFIG.AUDIT_COLUMNS.OLD_STATUS],
        newStatus: row[PR_CONFIG.AUDIT_COLUMNS.NEW_STATUS]
      }))
      .sort((a, b) => b.timestamp - a.timestamp);

  } catch (error) {
    console.error('Error getting audit history:', error);
    return [];
  }
}

/**
 * Gets status class for styling
 * @param {Object} poData - PO data object
 * @returns {string} CSS class for status
 */
function getStatusClass(poData) {
  if (!poData?.daysOverdue) return PR_CONFIG.STATUS_CLASSES.CURRENT;
  
  return poData.daysOverdue >= PR_CONFIG.WARNING_THRESHOLD 
    ? PR_CONFIG.STATUS_CLASSES.OVERDUE 
    : PR_CONFIG.STATUS_CLASSES.WARNING;
}

/**
 * Generates view link for PO
 * @param {string} poNumber - PO number
 * @returns {string} URL to view PO
 */
function generateViewLink(poNumber) {
  return `viewPO.html?po=${encodeURIComponent(poNumber)}`;
}

/**
 * Generates update link for PO
 * @param {string} poNumber - PO number
 * @returns {string} URL to update PO landing date
 */
function generateUpdateLink(poNumber) {
  return `updateLandingDate.html?po=${encodeURIComponent(poNumber)}`;
}

/**
 * Enhanced session logging configuration
 */
const LOGGING_CONFIG = {
  MAX_LOG_ROWS: 10000,
  ARCHIVE_PREFIX: 'Auth Log Archive'
};

/**
 * Archives old auth log entries
 * Creates monthly archive sheets when main log gets too large
 */
function archiveAuthLogs() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(CONFIG.AUTH_LOG_TAB);
    if (!logSheet) return;

    const today = new Date();
    const archiveName = `${LOGGING_CONFIG.ARCHIVE_PREFIX} ${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;

    // Create new archive sheet
    let archiveSheet = ss.getSheetByName(archiveName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveName);
      logSheet.getRange('1:1').copyTo(archiveSheet.getRange('1:1'), {contentsOnly: true});
      archiveSheet.setFrozenRows(1);
    }

    // Keep last 1000 rows in main log
    const lastRow = logSheet.getLastRow();
    const rowsToKeep = 1000;
    const rowsToMove = lastRow - rowsToKeep - 1;

    if (rowsToMove > 0) {
      const rangeToCopy = logSheet.getRange(2, 1, rowsToMove, logSheet.getLastColumn());
      const values = rangeToCopy.getValues();
      
      archiveSheet.getRange(
        archiveSheet.getLastRow() + 1, 
        1, 
        rowsToMove, 
        logSheet.getLastColumn()
      ).setValues(values);

      logSheet.deleteRows(2, rowsToMove);
      console.log(`Archived ${rowsToMove} rows to ${archiveName}`);
    }

  } catch (error) {
    console.error('Error archiving auth logs:', error);
  }
}

/**
 * Sets up cleanup trigger to run every hour
 */
function setupCleanupTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'cleanupSessions') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Create new hourly trigger
    ScriptApp.newTrigger('cleanupSessions')
      .timeBased()
      .everyHours(1)
      .create();

    console.log('Cleanup trigger set up successfully');
  } catch (error) {
    console.error('Error setting up cleanup trigger:', error);
  }
}

/**
 * Authentication Service Supplement
 * Contains core authentication functions for the PR system
 * 
 * Integration Notes:
 * - This should be placed at the end of Part 1 (Core Configuration and Session Management)
 * - Previous verifyUser() calls remain unchanged as this is the main entry point
 * - Enhanced with better session management and logging
 */

/**
 * Authentication configuration
 */
const AUTH_CONFIG = {
  VALID_ROLES: ['requestor', 'procurement', 'finance', 'approver'],
  DEFAULT_ROLE: 'requestor',
  COLUMN_INDICES: {
    NAME: 0,      // Column A: Name
    EMAIL: 1,     // Column B: Email
    DEPT: 2,      // Column C: Department
    ACTIVE: 3,    // Column D: Active (Y/N)
    PASSWORD: 4,  // Column E: Password
    ROLE: 5       // Column F: Role
  }
};

/**
 * Verifies user credentials and creates session
 * @param {string} username - Username to verify
 * @param {string} password - Password to verify
 * @returns {Object} Authentication result with session data
 */
function verifyUser(username, password) {
    Logger.log('==================== START verifyUser ====================');
    Logger.log('Verifying user:', username);

    try {
        // Get Requestor List sheet
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(CONFIG.REQUESTOR_SHEET_NAME);
        if (!sheet) {
            Logger.log('ERROR: User list sheet not found');
            throw new Error('User list not found');
        }
        Logger.log('Accessed Requestor List sheet');

        // Get all data and validate structure
        const data = sheet.getDataRange().getValues();
        const headers = data[0].map(h => h.toString().trim().toLowerCase());
        Logger.log('Sheet headers:', headers);

        // Validate required columns exist
        const requiredColumns = ['name', 'email', 'department', 'active (y/n)', 'password', 'role'];
        const missingColumns = requiredColumns.filter(col => !headers.includes(col));

        if (missingColumns.length > 0) {
            Logger.log('ERROR: Missing columns:', missingColumns);
            throw new Error(`Required columns not found: ${missingColumns.join(', ')}`);
        }

        // Find user row
        const userRow = data.find(row => row[AUTH_CONFIG.COLUMN_INDICES.NAME] === username);
        if (!userRow) {
            Logger.log('User not found:', username);
            return {
                success: false,
                message: 'Invalid username or password'
            };
        }
        Logger.log('Found user row:', JSON.stringify(userRow));

        // Check if user is active
        if (userRow[AUTH_CONFIG.COLUMN_INDICES.ACTIVE].toString().toUpperCase() !== 'Y') {
            Logger.log('Inactive account attempt:', username);
            return {
                success: false,
                message: 'Account is inactive. Please contact procurement team.'
            };
        }

        // Verify password
        if (userRow[AUTH_CONFIG.COLUMN_INDICES.PASSWORD] !== password) {
            Logger.log('Invalid password attempt for:', username);
            logAuthEvent(username, false, 'Failed login attempt - invalid password');
            return {
                success: false,
                message: 'Invalid username or password'
            };
        }
        Logger.log('Password verified successfully');

        // Get user's role and validate
        const userRole = userRow[AUTH_CONFIG.COLUMN_INDICES.ROLE]?.toString().toLowerCase() || AUTH_CONFIG.DEFAULT_ROLE;
        if (!AUTH_CONFIG.VALID_ROLES.includes(userRole)) {
            Logger.log(`Warning: Invalid role "${userRole}" for user ${username}, defaulting to ${AUTH_CONFIG.DEFAULT_ROLE}`);
        }

        // Create user session
        const userSession = {
            username: username,
            role: AUTH_CONFIG.VALID_ROLES.includes(userRole) ? userRole : AUTH_CONFIG.DEFAULT_ROLE,
            department: userRow[AUTH_CONFIG.COLUMN_INDICES.DEPT],
            email: userRow[AUTH_CONFIG.COLUMN_INDICES.EMAIL],
            timestamp: new Date().toISOString()
        };
        Logger.log('Created user session:', JSON.stringify(userSession));

        // Store session in cache
        const userCache = CacheService.getUserCache();
        userCache.put('userSession', JSON.stringify(userSession), 86400);
        Logger.log('Stored session in cache');

        let sessionId = null;

        // Create sheet session
        try {
            sessionId = createSheetSession(userSession);
            Logger.log('Sheet session created:', sessionId);
        } catch (sessionError) {
            Logger.log('Warning: Sheet session creation failed:', sessionError);
        }

        // Log successful login
        logAuthEvent(username, true, 'Successful login');
        Logger.log('Login successful for:', username);

        // Get web app URL and state token for secure redirect
        const redirectUrl = getWebAppUrl();
        const stateToken = Utilities.getUuid();

        // Store state token in cache briefly
        const stateCache = CacheService.getScriptCache();
        stateCache.put(`state_${stateToken}`, sessionId, 300); // 5 minute expiry

        Logger.log('Generated redirect URL:', redirectUrl);
        Logger.log('Generated state token for secure redirect');

        const response = {
            success: true,
            redirectUrl: redirectUrl,
            sessionId: sessionId,
            stateToken: stateToken,
            user: {
                username: userSession.username,
                role: userSession.role,
                department: userSession.department
            }
        };
        Logger.log('Returning response:', JSON.stringify(response));
        Logger.log('==================== END verifyUser ====================');
        return response;

    } catch (error) {
        Logger.log('ERROR in verifyUser:', error.toString());
        Logger.log('Error stack:', error.stack);
        logAuthEvent(username, false, `Login error: ${error.message}`);
        Logger.log('==================== END verifyUser (with error) ====================');
        return {
            success: false,
            message: error.message || 'System error during login. Please try again later.'
        };
    }
}

/**
 * Creates sheet-based session entry with enhanced error handling
 * @param {Object} user - User session data
 * @returns {string} Session ID
 */
function createSheetSession(user) {
  const sessionId = Utilities.getUuid();
  const timestamp = new Date().toISOString();

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sessionsSheet = ss.getSheetByName(SESSION_CONFIG.SHEET_NAME);

    if (!sessionsSheet) {
      sessionsSheet = createSessionLog();
    }

    const deviceInfo = {
      userAgent: HtmlService.getUserAgent(),
      timestamp: timestamp
    };

    sessionsSheet.appendRow([
      sessionId,
      user.username,
      user.email,
      user.role,
      timestamp,
      timestamp,
      SESSION_CONFIG.STATUS.ACTIVE,
      JSON.stringify(deviceInfo)
    ]);

    return sessionId;

  } catch (error) {
    console.error('Error creating sheet session:', error);
    throw new Error('Failed to create session: ' + error.message);
  }
}


/**
 * Authentication Logging Supplement
 * Enhanced logging functionality for authentication and access events
 * 
 * Integration Notes:
 * - This should be placed in Part 3 with other logging functions
 * - Works with existing archiveAuthLogs() function
 * - Integrates with the session management system
 */







/**
 * Gets context information for the current event
 * @returns {Object} Context information including IP and user agent
 */
function getEventContext() {
  let ipAddress = '';
  let userAgent = '';

  try {
    const request = HtmlService.getCurrentRequest();
    if (request) {
      userAgent = request.getUserAgent() || '';
      
      // Note: Direct IP address access is not available in Apps Script
      // Instead, we might get it from headers if available in your setup
      // This would need to be customized based on your deployment
      // ipAddress = request.getHeader('X-Forwarded-For') || '';
    }
  } catch (e) {
    console.warn('Could not get complete request details:', e);
  }

  return {
    ipAddress,
    userAgent
  };
}

/**
 * Determines the event type from the details string
 * @param {string} details - Event details
 * @returns {string} Event type
 */
function determineEventType(details) {
  const detailsLower = (details || '').toLowerCase();
  
  if (detailsLower.includes('login')) {
    return AUTH_LOG_CONFIG.EVENT_TYPES.LOGIN;
  }
  if (detailsLower.includes('logout')) {
    return AUTH_LOG_CONFIG.EVENT_TYPES.LOGOUT;
  }
  if (detailsLower.includes('session')) {
    return AUTH_LOG_CONFIG.EVENT_TYPES.SESSION;
  }
  return AUTH_LOG_CONFIG.EVENT_TYPES.ACCESS;
}


/**
 * Corrected event context handling
 * Removes incorrect HtmlService.getCurrentRequest() usage
 */

/**
 * Gets context information for the current event
 * @returns {Object} Context information including user agent
 */
function getEventContext() {
  let userAgent = '';

  try {
    // In Google Apps Script, we can only get limited context
    // User agent can be accessed when serving HTML
    userAgent = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    console.warn('Could not get user context:', e);
  }

  return {
    userAgent,
    timestamp: new Date().toISOString()
  };
}

/**
 * Logs authentication and access events - Corrected Version
 * @param {string} username - Username involved in event
 * @param {boolean} success - Whether event was successful
 * @param {string} details - Additional event details
 * @returns {boolean} Success status of logging
 */
function logAuthEvent(username, success, details) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(AUTH_LOG_CONFIG.SHEET_NAME);

    // Create Auth Log sheet if it doesn't exist
    if (!logSheet) {
      logSheet = createAuthLogSheet(ss);
    }

    // Get event context (simplified)
    const eventContext = getEventContext();
    
    // Determine event type from details
    const eventType = determineEventType(details);

    // Add log entry with available information
    const logEntry = [
      new Date(), // Timestamp
      username || 'Unknown',
      eventType,
      success ? 'Y' : 'N',
      details || '',
      eventContext.userAgent, // User email instead of IP
      eventContext.timestamp  // Timestamp instead of user agent
    ];

    // Add the log entry
    logSheet.appendRow(logEntry);

    // Check if archiving is needed
    const currentRows = logSheet.getLastRow();
    if (currentRows > AUTH_LOG_CONFIG.MAX_ROWS) {
      archiveAuthLogs();
    }

    // Log to console for debugging
    console.log('Auth event logged:', {
      username,
      eventType,
      success,
      details,
      context: eventContext
    });

    return true;

  } catch (error) {
    console.error('Error logging auth event:', error, {
      username,
      success,
      details
    });
    return false;
  }
}

// Update AUTH_LOG_CONFIG columns to match new structure
const AUTH_LOG_CONFIG = {
  SHEET_NAME: 'Auth Log',
  COLUMNS: {
    TIMESTAMP: 0,
    USERNAME: 1,
    EVENT_TYPE: 2,
    SUCCESS: 3,
    DETAILS: 4,
    USER_EMAIL: 5,  // Changed from IP_ADDRESS
    EVENT_TIME: 6   // Changed from USER_AGENT
  },
  EVENT_TYPES: {
    LOGIN: 'Login',
    LOGOUT: 'Logout',
    ACCESS: 'Access',
    SESSION: 'Session'
  },
  MAX_ROWS: 10000,
  COLUMN_WIDTHS: {
    TIMESTAMP: 180,
    USERNAME: 150,
    EVENT_TYPE: 120,
    DETAILS: 300,
    USER_EMAIL: 200,    // Adjusted width
    EVENT_TIME: 180     // Adjusted width
  }
};

/**
 * Creates the Auth Log sheet with proper structure - Updated Version
 * @param {Spreadsheet} ss - Spreadsheet object
 * @returns {Sheet} The created sheet
 */
function createAuthLogSheet(ss) {
  const sheet = ss.insertSheet(AUTH_LOG_CONFIG.SHEET_NAME);
  
  // Set up headers with corrected column names
  const headers = [
    'Timestamp',
    'Username',
    'Event Type',
    'Success',
    'Details',
    'User Email',     // Updated header
    'Event Time'      // Updated header
  ];

  // Apply header formatting
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E8EAED');

  // Set frozen rows
  sheet.setFrozenRows(1);

  // Set column widths
  Object.entries(AUTH_LOG_CONFIG.COLUMN_WIDTHS).forEach(([col, width], index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  // Add data validation for Success column
  const successRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y', 'N'])
    .build();
  sheet.getRange(2, AUTH_LOG_CONFIG.COLUMNS.SUCCESS + 1, sheet.getMaxRows() - 1, 1)
    .setDataValidation(successRule);

  // Add data validation for Event Type column
  const eventTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(AUTH_LOG_CONFIG.EVENT_TYPES))
    .build();
  sheet.getRange(2, AUTH_LOG_CONFIG.COLUMNS.EVENT_TYPE + 1, sheet.getMaxRows() - 1, 1)
    .setDataValidation(eventTypeRule);

  return sheet;
}


function getCurrentUser(sessionId) {
    console.log('Getting current user for session:', sessionId);
    
    try {
        if (!sessionId) {
            console.log('No session ID provided');
            return null;
        }

        // Only check sheet session
        const sheetSession = validateSession(sessionId);
        if (sheetSession) {
            console.log('Found valid sheet session for:', sheetSession.username);
            return sheetSession;
        }

        console.log('No valid session found');
        return null;
    } catch (error) {
        console.error('Error in getCurrentUser:', error);
        return null;
    }
}

function validateSession(sessionId) {
    if (!sessionId) return null;
    
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SESSION_CONFIG.SHEET_NAME);
        if (!sheet) {
            console.warn('Session sheet not found');
            return null;
        }

        const data = sheet.getDataRange().getValues();
        const sessionRow = data.find(row => 
            row[SESSION_CONFIG.COLUMNS.SESSION_ID] === sessionId &&
            row[SESSION_CONFIG.COLUMNS.STATUS] === SESSION_CONFIG.STATUS.ACTIVE
        );

        if (!sessionRow) {
            console.log('No active session found for ID:', sessionId);
            return null;
        }

        // Check session expiry
        const lastActivity = new Date(sessionRow[SESSION_CONFIG.COLUMNS.LAST_ACTIVITY]);
        const now = new Date();
        const inactiveMinutes = (now - lastActivity) / (1000 * 60);

        if (inactiveMinutes > SESSION_CONFIG.INACTIVITY_TIMEOUT) {
            console.log('Session expired due to inactivity');
            // Update session status to expired
            markSessionExpired(sessionId);
            return null;
        }

        return {
            username: sessionRow[SESSION_CONFIG.COLUMNS.USERNAME],
            email: sessionRow[SESSION_CONFIG.COLUMNS.EMAIL],
            role: sessionRow[SESSION_CONFIG.COLUMNS.ROLE],
            sessionId: sessionId
        };

    } catch (error) {
        console.error('Error validating session:', error);
        return null;
    }
}

function markSessionExpired(sessionId) {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SESSION_CONFIG.SHEET_NAME);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        const rowIndex = data.findIndex(row => row[SESSION_CONFIG.COLUMNS.SESSION_ID] === sessionId);
        
        if (rowIndex > -1) {
            sheet.getRange(rowIndex + 1, SESSION_CONFIG.COLUMNS.STATUS + 1)
                .setValue(SESSION_CONFIG.STATUS.EXPIRED);
            console.log('Session marked as expired:', sessionId);
        }
    } catch (error) {
        console.error('Error marking session expired:', error);
    }
}

/**
 * Gets redirect URL for secure navigation
 * @returns {Object} Object containing URL and timestamp
 */
function getRedirectUrl() {
    try {
        const baseUrl = ScriptApp.getService().getUrl();
        const webAppUrl = ScriptApp.getService()
            .getUrl()
            .replace('/exec', '/dev'); // Use /dev for development, /exec for production
            
        return {
            url: webAppUrl,
            timestamp: new Date().getTime()
        };
    } catch (error) {
        console.error('Error getting redirect URL:', error);
        throw new Error('Failed to generate redirect URL');
    }
}

function handleLogin(event) {
    event.preventDefault();

    if (isLoginInProgress) {
        console.log('Login already in progress');
        return;
    }

    console.log('Starting login process');
    isLoginInProgress = true;

    // Check if SessionUtils is available
    if (typeof SessionUtils === 'undefined') {
        console.error('SessionUtils not loaded');
        showError('Session management system not loaded. Please refresh the page.');
        isLoginInProgress = false;
        return;
    }

    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    const errorDiv = document.getElementById('error');
    const loadingDiv = document.getElementById('loading');
    const loginButton = document.getElementById('loginButton');

    // Clear previous errors and show loading
    errorDiv.classList.add('hidden');
    loadingDiv.classList.remove('hidden');
    loginButton.disabled = true;

    // First step: Verify user credentials
    google.script.run
        .withFailureHandler(handleLoginError)
        .withSuccessHandler(handleLoginSuccess)
        .verifyUser(username, password);

    function handleLoginError(error) {
        console.error('Login failed:', error);
        showError(error.message || 'Login failed. Please try again.');
        loadingDiv.classList.add('hidden');
        loginButton.disabled = false;
        isLoginInProgress = false;
    }

    function handleLoginSuccess(result) {
        console.log('Received authentication response');

        if (!result.success) {
            handleLoginError(result);
            return;
        }

        try {
            // Store session ID using SessionUtils
            if (result.sessionId && typeof SessionUtils !== 'undefined') {
                SessionUtils.setSessionId(result.sessionId);
                
                // Get redirect URL and handle navigation
                google.script.run
                    .withSuccessHandler(function(redirectInfo) {
                        if (!redirectInfo || !redirectInfo.url) {
                            handleLoginError({ message: 'Failed to get redirect URL' });
                            return;
                        }

                        try {
                            let redirectUrl = redirectInfo.url;
                            
                            // Append session ID to URL if SessionUtils is available
                            if (typeof SessionUtils !== 'undefined') {
                                redirectUrl = SessionUtils.appendSessionToUrl(redirectUrl);
                            } else {
                                // Fallback: manually append session ID
                                const separator = redirectUrl.includes('?') ? '&' : '?';
                                redirectUrl = `${redirectUrl}${separator}sessionId=${encodeURIComponent(result.sessionId)}`;
                            }

                            // Use the closeOrRedirect utility if available, otherwise direct navigation
                            if (typeof closeOrRedirect !== 'undefined') {
                                closeOrRedirect(redirectUrl);
                            } else {
                                // Fallback navigation
                                if (google.script.host) {
                                    google.script.host.close();
                                }
                                window.top.location.href = redirectUrl;
                            }
                        } catch (e) {
                            console.error('Error during navigation:', e);
                            handleLoginError({ message: 'Navigation error occurred' });
                        }
                    })
                    .withFailureHandler(function(error) {
                        console.error('Redirect URL request failed:', error);
                        handleLoginError({ message: 'Failed to complete login process' });
                    })
                    .getRedirectUrl();
            } else {
                handleLoginError({ message: 'Session management not available' });
            }
        } catch (e) {
            console.error('Error during login success handling:', e);
            handleLoginError({ message: 'Error during login process' });
        }
    }
}

function handleLoginError(error) {
  console.log('Login failed:', error);
  const errorDiv = document.getElementById('error');
  const loadingDiv = document.getElementById('loading');
  const loginButton = document.getElementById('loginButton');

  // Show error message
  errorDiv.textContent = error.message || 'Login failed. Please try again.';
  errorDiv.classList.remove('hidden');

  // Hide loading and re-enable button
  loadingDiv.classList.add('hidden');
  loginButton.disabled = false;
  isLoginInProgress = false;
}

// Load user options using google.script.run
function loadUserOptions() {
  google.script.run
    .withFailureHandler(function(error) {
      console.error('Failed to load users:', error);
      const errorOption = document.createElement('option');
      errorOption.text = 'Error loading users';
      errorOption.disabled = true;
      document.getElementById('username').add(errorOption);
    })
    .withSuccessHandler(function(users) {
      const select = document.getElementById('username');
      // Clear existing options except default
      while (select.options.length > 1) {
        select.remove(1);
      }
      // Add user options
      users.forEach(function(user) {
        const option = document.createElement('option');
        option.value = user.name;
        option.text = user.name;
        select.add(option);
      });
    })
    .getActiveRequestors();
}

