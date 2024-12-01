/**
 * 241107_PRView.gs
 * ================
 * Part of 1PWR Procurement System
 * Version: 1.0
 * Last Updated: 2024-11-07
 * 
 * Purpose:
 * --------
 * Handles the retrieval and display of individual Purchase Request (PR) details.
 * Provides the server-side functionality for the PR view page, which shows
 * comprehensive information about a specific PR including its status history
 * and all related documentation.
 * 
 * Integration Points:
 * ------------------
 * - Code.gs: Uses CONFIG and COL constants for sheet structure
 * - StatusWorkflow.gs: References status information and history
 * - Master Log Sheet: Primary data source for PR details
 * - Audit Log Sheet: Source for PR status history
 * 
 * Sheet Dependencies:
 * -----------------
 * Master Log Tab Columns:
 * - Uses column mapping from COL constant in Code.gs
 * - Requires all standard PR fields as defined in Master Sheet Structure
 * 
 * Audit Log Tab Columns:
 * - Timestamp (A)
 * - User (B)
 * - Action (C)
 * - PR Number (D)
 * - Details (E)
 * - Old Status (F)
 * - New Status (G)
 * 
 * Usage:
 * ------
 * Called by doGet() when URL includes pr parameter
 * Example: ?pr=PR-202411-001
 */




/**
 * Calculates business days a PR has been open
 * @param {Date} timestamp - PR submission date
 * @returns {number} Number of business days
 */
function calculateDaysOpen(timestamp) {
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
 * Enhanced PRView.gs with session management
 * Maintains existing functionality while adding session validation
 */

/**
 * Serves the PR view page with role-based access and session validation
 * @param {Object} e - Event object containing PR number and session ID
 * @returns {HtmlOutput} The formatted PR view page
 */
function servePRView(e) {
  console.log('Starting PR view request');
  
  try {
    // Validate session
    const sessionId = e.parameter.sessionId;
    const user = getCurrentUser(sessionId);
    
    if (!user) {
      console.log('No valid session found');
      return createLoginRedirect();
    }

    const prNumber = e.parameter.pr;
    if (!prNumber) {
      throw new Error('No PR number provided');
    }

    const prDetails = getPRDetails(prNumber);
    if (!prDetails.success) {
      throw new Error(prDetails.error);
    }

    // Create and populate template with user context
    const template = HtmlService.createTemplateFromFile('PRView');
    template.prData = prDetails.data;
    template.userRole = user.role;
    template.canUpdateStatus = isAuthorized(user, 'procurement');
    template.sessionId = sessionId; // Pass session ID to template

    return template.evaluate()
      .setTitle(`PR ${prNumber}`)
      .setFaviconUrl('https://1pwrafrica.com/wp-content/uploads/2018/11/logo.png');

  } catch (error) {
    console.error('Error serving PR view:', error);
    return createErrorPage('Error displaying Purchase Request details.');
  }
}

/**
 * Enhanced status update handler with session validation
 */
function updateStatus(prNumber, newStatus, notes, sessionId) {
  console.log('Processing status update request');
  
  try {
    const user = getCurrentUser(sessionId);
    if (!user) {
      return {
        success: false,
        sessionExpired: true
      };
    }

    if (!isAuthorized(user, 'procurement')) {
      throw new Error('Not authorized to update status');
    }

    // Use existing workflow manager for status update
    const workflow = getStatusWorkflowManager();
    const result = workflow.updateStatus(prNumber, newStatus, notes);

    if (result.success) {
      // Log the status change
      logStatusChange(prNumber, newStatus, notes, user);
    }

    return result;

  } catch (error) {
    console.error('Error updating status:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Creates login redirect page
 * @returns {HtmlOutput} Login page redirect
 */
function createLoginRedirect() {
  const template = HtmlService.createTemplate(
    `<!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <script>
          window.top.location.href = '<?= getWebAppUrl() ?>?page=login';
        </script>
      </head>
    </html>`
  );
  
  return template.evaluate();
}

/**
 * Logs status changes with user context
 */
function logStatusChange(prNumber, newStatus, notes, user) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('Status Change Log');
    
    if (!logSheet) {
      console.warn('Status Change Log sheet not found');
      return;
    }

    logSheet.appendRow([
      new Date(),
      prNumber,
      user.username,
      user.email,
      newStatus,
      notes
    ]);

  } catch (error) {
    console.error('Error logging status change:', error);
  }
}