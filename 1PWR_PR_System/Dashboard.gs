/**
 * Dashboard.gs
 * 1PWR Procurement System - Dashboard Backend
 * Version: 2.3
 * Last Updated: 2024-11-19
 * 
 * Purpose:
 * --------
 * Provides server-side functionality for the procurement dashboard, handling data
 * retrieval, organization filtering, and status-based presentation of Purchase
 * Requisitions (PRs) and Purchase Orders (POs). Supports a grid-based view of PRs
 * and POs organized by their current status categories. Additionally, incorporates
 * role-based permissions to control user capabilities within the dashboard.
 * 
 * Core Functions:
 * --------------
 * 1. Data Retrieval and Organization:
 *    - Fetches PR and PO data from the Master Log sheet.
 *    - Filters data by organization (1PWR LESOTHO/SMP/PUECO).
 *    - Groups PRs and POs by status categories.
 *    - Calculates status counts and system-wide metrics.
 * 
 * 2. Status Categories Managed:
 *    - PRs:
 *      - Submitted: New PRs awaiting processing.
 *      - In Queue: PRs under active processing.
 *      - Ordered: PRs that have been ordered.
 *      - Completed: Fully processed PRs.
 *      - R&R: PRs requiring revision and resubmission.
 *      - Rejected: Declined PRs.
 *      - Canceled: Terminated PRs.
 *    - POs:
 *      - PO Pending Approval
 *      - PO Approved
 *      - PO Ordered
 *      - Completed
 *      - Rejected
 *      - Canceled
 * 
 * 3. Role-Based Permissions:
 *    - Determines user permissions based on roles.
 *    - Controls actions such as submitting PRs, updating statuses, approving requests, and viewing financial data.
 * 
 * Dependencies:
 * ------------
 * - Code.gs: System configuration and column definitions (CONFIG, COL constants).
 * - StatusWorkflow.gs: Status management and validation.
 * - NotificationSystem.gs: Email notifications.
 * - PRView.gs: Individual PR view functionality.
 * - AuthService.gs: User authentication and authorization.
 * 
 * Sheet Structure Dependencies:
 * ----------------------------
 * Master Log Tab:
 * - PR Number (COL.PR_NUMBER)
 * - PO Number (COL.PO_NUMBER)
 * - Status (COL.PR_STATUS / COL.PO_STATUS)
 * - Organization (COL.ORGANIZATION)
 * - Description (COL.DESCRIPTION)
 * - Timestamp (COL.TIMESTAMP)
 * - Requestor (COL.REQUESTOR_NAME)
 * - Additional tracking fields as defined in COL constant.
 * 
 * Data Return Structure:
 * ---------------------
 * {
 *   success: boolean,
 *   data: {
 *     countsByStatus: {
 *       'Submitted': number,
 *       'In Queue': number,
 *       'Ordered': number,
 *       'Completed': number,
 *       'R&R': number,
 *       'Rejected': number,
 *       'Canceled': number
 *     },
 *     prs: {
 *       'Submitted': Array<PRData>,
 *       'In Queue': Array<PRData>,
 *       'Ordered': Array<PRData>,
 *       'Completed': Array<PRData>,
 *       'R&R': Array<PRData>,
 *       'Rejected': Array<PRData>,
 *       'Canceled': Array<PRData>
 *     },
 *     metrics: {
 *       totalPRs: number,
 *       urgentPRs: number,
 *       avgDaysOpen: number,
 *       overduePRs: number,
 *       quotesRequired: number,
 *       adjudicationRequired: number,
 *       customsClearanceRequired: number,
 *       completionRate: number
 *     },
 *     permissions: {
 *       canSubmitPR: boolean,
 *       canUpdateStatus: boolean,
 *       canApprove: boolean,
 *       canViewFinance: boolean
 *     }
 *   },
 *   timestamp: string (ISO date)
 * }
 * 
 * Integration Points:
 * -------------------
 * - Called by DashboardWeb.html for initial load and refresh.
 * - Supports organization filtering from UI.
 * - Provides data for PR view navigation.
 * - Integrates with status workflow system.
 * - Utilizes authentication service for role-based permissions.
 * 
 * Error Handling:
 * ---------------
 * - Returns structured error responses.
 * - Logs errors for debugging.
 * - Maintains data integrity during filtering.
 * 
 * Usage:
 * ------
 * Called from client-side via google.script.run:
 * google.script.run
 *   .withSuccessHandler(handleSuccess)
 *   .withFailureHandler(handleError)
 *   .getDashboardData(organization);
 */


/**
 * getDashboardData
 * Retrieves and processes PR/PO data for the procurement dashboard, filtering by organization
 * and computing metrics. Includes user permissions based on roles in Requestor List.
 * 
 * @param {string} organization - Optional organization filter (e.g., '1PWR LESOTHO', 'SMP', 'PUECO')
 * @returns {Object} Dashboard data object with structure:
 * {
 *   success: boolean,
 *   data: {
 *     countsByStatus: {
 *       'Submitted': number,
 *       'In Queue': number,
 *       'Ordered': number,
 *       'Completed': number,
 *       'R&R': number,
 *       'Rejected': number,
 *       'Canceled': number
 *     },
 *     prs: {
 *       'Submitted': Array<PRData>,
 *       'In Queue': Array<PRData>,
 *       'Ordered': Array<PRData>,
 *       'Completed': Array<PRData>,
 *       'R&R': Array<PRData>,
 *       'Rejected': Array<PRData>,
 *       'Canceled': Array<PRData>
 *     },
 *     metrics: {
 *       totalPRs: number,
 *       urgentPRs: number,
 *       avgDaysOpen: number,
 *       overduePRs: number,
 *       quotesRequired: number,
 *       adjudicationRequired: number,
 *       customsClearanceRequired: number,
 *       completionRate: number
 *     },
 *     permissions: {
 *       canSubmitPR: boolean,
 *       canUpdateStatus: boolean,
 *       canApprove: boolean,
 *       canViewFinance: boolean
 *     }
 *   },
 *   timestamp: string // ISO date string
 * }
 * 
 * Dependencies:
 * - getCurrentUser(): Gets authenticated user info
 * - isAuthorized(user, role): Checks user role permissions
 * - parseLineItems(row): Processes line items from row data
 * - calculateDaysOpen(timestamp): Calculates days PR has been open
 * - calculateDaysOverdue(date): Calculates days overdue from expected date
 * - calculateBusinessDays(start, end): Calculates business days between dates
 * - getDetailedBlockingItem(row): Gets blocking item info for ordered PRs
 * 
 * Required Sheet Structure:
 * - Master Log Tab: Main PR/PO tracking sheet (see COL constant for mapping)
 * - Requestor List: User roles and permissions (Name, Email, Department, Active, Password, Role)
 */
function getDashboardData(organization) {
  Logger.log('Starting getDashboardData with organization: ' + organization);

  try {
    // Authenticate user and retrieve permissions
    const user = getCurrentUser();
    if (!user) {
      throw new Error('Authentication required');
    }
    Logger.log('User authenticated: ' + user.email);

    // Get user permissions
    const userPermissions = {
      canSubmitPR: true, // All authenticated users can submit
      canUpdateStatus: isAuthorized(user, 'procurement'),
      canApprove: isAuthorized(user, 'approver'),
      canViewFinance: isAuthorized(user, 'finance')
    };

    // Open the spreadsheet using the configured SPREADSHEET_ID
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      throw new Error('Failed to open spreadsheet with ID: ' + SPREADSHEET_ID);
    }
    Logger.log('Spreadsheet opened successfully.');

    // Access the "Master Log" sheet as defined in CONFIG
    const sheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);
    if (!sheet) {
      throw new Error('Master Log sheet not found in spreadsheet.');
    }
    Logger.log('Master Log sheet accessed successfully.');

    // Initialize data structure with all possible statuses
    const statusData = {
      countsByStatus: {
        'Submitted': 0,
        'In Queue': 0,
        'Ordered': 0,
        'Completed': 0,
        'R&R': 0,
        'Rejected': 0,
        'Canceled': 0
      },
      prs: {
        'Submitted': [],
        'In Queue': [],
        'Ordered': [],
        'Completed': [],
        'R&R': [],
        'Rejected': [],
        'Canceled': []
      },
      metrics: {
        totalPRs: 0,
        urgentPRs: 0,
        avgDaysOpen: 0,
        overduePRs: 0,
        quotesRequired: 0,
        adjudicationRequired: 0,
        customsClearanceRequired: 0,
        completionRate: 0
      }
    };

    // Retrieve all data from the sheet
    const data = sheet.getDataRange().getValues();
    Logger.log('Retrieved rows: ' + (data.length - 1)); // Subtract header row

    // If only header row exists, return empty statusData with permissions
    if (data.length <= 1) {
      Logger.log('No PR data found in sheet. Returning empty statusData.');
      return {
        success: true,
        data: {
          ...statusData,
          permissions: userPermissions
        },
        timestamp: new Date().toISOString()
      };
    }

    // Debug first row data
    if (data.length > 1) {
      Logger.log('First row data sample: ' + JSON.stringify(data[1]));
      Logger.log('Status field value: ' + data[1][COL.PR_STATUS]);
      Logger.log('Organization field value: ' + data[1][COL.ORGANIZATION]);
    }

    // Initialize variables for metrics calculation
    let totalDaysOpen = 0;
    let activePRs = 0;
    let completedPRs = 0;
    let totalCompletionPercentage = 0;

    // Process each row (skip header)
    data.slice(1).forEach((row, index) => {
      const prOrg = row[COL.ORGANIZATION];
      const status = row[COL.PR_STATUS];

      Logger.log(`Checking PR in row ${index + 2}: Status=${status}, Organization=${prOrg}`);

      // Check if the PR should be included based on organization filter
      if (!organization || prOrg === organization) {
        Logger.log(`Processing row ${index + 2}, status: ${status}`);

        // Only process if status is recognized
        if (statusData.countsByStatus.hasOwnProperty(status)) {
          // Increment status count
          statusData.countsByStatus[status]++;
          statusData.metrics.totalPRs++;

          // Calculate metrics based on PR attributes
          if (row[COL.URGENCY] === 'Y') {
            statusData.metrics.urgentPRs++;
          }
          if (row[COL.QUOTES_REQUIRED] === 'Y') {
            statusData.metrics.quotesRequired++;
          }
          if (row[COL.ADJ_REQUIRED] === 'Y') {
            statusData.metrics.adjudicationRequired++;
          }
          if (row[COL.CUSTOMS_REQUIRED] === 'Y') {
            statusData.metrics.customsClearanceRequired++;
          }

          // Calculate days open
          const daysOpen = calculateDaysOpen(row[COL.TIMESTAMP]);
          totalDaysOpen += daysOpen;
          activePRs++;

          // Calculate completion percentage
          const completion = row[COL.COMPLETION] || 0;
          totalCompletionPercentage += completion;

          // Track completed PRs
          if (status === 'Completed') {
            completedPRs++;
          }

          // Parse line items
          const lineItems = parseLineItems(row);
          Logger.log('Parsed line items:', lineItems);

          // Create PR data object with status-specific fields
          const prData = {
            prNumber: row[COL.PR_NUMBER],
            timestamp: row[COL.TIMESTAMP] ? new Date(row[COL.TIMESTAMP]).toISOString() : 'No timestamp provided',
            description: row[COL.DESCRIPTION] || 'No description provided',
            requestor: row[COL.REQUESTOR_NAME] || 'Unknown',
            department: row[COL.DEPARTMENT] || 'N/A',
            daysOpen: daysOpen,
            urgency: row[COL.URGENCY] === 'Y' ? 'Urgent' : 'Normal',
            amount: row[COL.PR_AMOUNT],
            currency: row[COL.CURRENCY],
            vendor: row[COL.VENDOR],
            approver: row[COL.APPROVER],
            completion: completion,
            procurementNotes: row[COL.PROCUREMENT_NOTES] || '',
            lineItems: lineItems,
            _debug: {
              rawLineItems: {
                items: row[COL.ITEM_LIST],
                quantities: row[COL.QTY_LIST],
                uoms: row[COL.UOM_LIST],
                urls: row[COL.URL_LIST]
              }
            }
          };

          // Add status-specific fields
          switch (status) {
            case 'Submitted':
              prData.submittedDate = row[COL.TIMESTAMP] ? new Date(row[COL.TIMESTAMP]).toISOString() : 'No submission date';
              break;

            case 'In Queue':
              prData.queuePosition = calculateQueuePosition(row, data);
              prData.quotesRequired = row[COL.QUOTES_REQUIRED] === 'Y';
              prData.quotesStatus = row[COL.QUOTES_LINK] ? 'Received' : 'Pending';
              prData.adjudicationRequired = row[COL.ADJ_REQUIRED] === 'Y';
              prData.adjudicationStatus = row[COL.ADJ_NOTES] ? 'Completed' : 'Pending';
              break;

            case 'Ordered':
              prData.orderedDate = row[COL.PO_DATE] ? new Date(row[COL.PO_DATE]).toISOString() : 'No ordered date';
              prData.expectedLanding = row[COL.EXPECTED_LANDING_DATE] ? new Date(row[COL.EXPECTED_LANDING_DATE]).toISOString() : 'No expected landing date';
              prData.daysOverdue = calculateDaysOverdue(row[COL.EXPECTED_LANDING_DATE]);
              prData.blockingItem = getDetailedBlockingItem(row);
              prData.paymentDate = row[COL.PAYMENT_DATE] ? new Date(row[COL.PAYMENT_DATE]).toISOString() : 'No payment date';
              prData.shipmentStatus = {
                shipped: row[COL.SHIPPED] === 'Y',
                shipmentDate: row[COL.SHIPMENT_DATE] ? new Date(row[COL.SHIPMENT_DATE]).toISOString() : 'No shipment date',
                customsRequired: row[COL.CUSTOMS_REQUIRED] === 'Y',
                customsCleared: row[COL.CUSTOMS_CLEARED] === 'Y',
                customsClearanceDate: row[COL.DATE_CLEARED] ? new Date(row[COL.DATE_CLEARED]).toISOString() : 'No customs clearance date',
                goodsLanded: row[COL.GOODS_LANDED] === 'Y',
                landedDate: row[COL.LANDED_DATE] ? new Date(row[COL.LANDED_DATE]).toISOString() : 'No landed date'
              };

              // Track overdue PRs
              if (prData.daysOverdue > 0) {
                statusData.metrics.overduePRs++;
              }
              break;

            case 'R&R':
              prData.rAndRDate = row[COL.OVERRIDE_DATE] ? new Date(row[COL.OVERRIDE_DATE]).toISOString() : new Date().toISOString();
              prData.rAndRBy = row[COL.OVERRIDE_BY] || '';
              prData.rAndRJustification = row[COL.OVERRIDE_JUST] || '';
              break;

            case 'Completed':
              prData.completedDate = row[COL.LANDED_DATE] ? new Date(row[COL.LANDED_DATE]).toISOString() : 'No completed date';
              prData.timeToComplete = calculateBusinessDays(row[COL.TIMESTAMP], row[COL.LANDED_DATE]);
              prData.timeToShip = row[COL.TIME_TO_SHIP] || null;
              prData.timeInCustoms = row[COL.TIME_IN_CUSTOMS] || null;
              prData.timeToLand = row[COL.TIME_TO_LAND] || null;
              break;

            case 'Rejected':
              prData.rejectionDate = row[COL.OVERRIDE_DATE] ? new Date(row[COL.OVERRIDE_DATE]).toISOString() : 'No rejection date';
              prData.rejectedBy = row[COL.OVERRIDE_BY] || '';
              prData.rejectionReason = row[COL.OVERRIDE_JUST] || '';
              break;

            case 'Canceled':
              prData.cancelationDate = row[COL.OVERRIDE_DATE] ? new Date(row[COL.OVERRIDE_DATE]).toISOString() : 'No cancelation date';
              prData.canceledBy = row[COL.OVERRIDE_BY] || '';
              prData.cancelationReason = row[COL.OVERRIDE_JUST] || '';
              break;
          }

          // Add PR data to appropriate status list
          statusData.prs[status].push(prData);
          Logger.log(`Added PR ${prData.prNumber} to status '${status}'.`);
        } else {
          Logger.warn(`Unrecognized status '${status}' in row ${index + 2}. Skipping.`);
        }
      }
    });

    // Calculate final metrics
    if (activePRs > 0) {
      statusData.metrics.avgDaysOpen = Math.round(totalDaysOpen / activePRs);
      statusData.metrics.completionRate = Math.round(totalCompletionPercentage / activePRs);
    }

    Logger.log('Successfully processed data: ' + JSON.stringify(statusData));

    // Return the processed data with permissions
    return {
      success: true,
      data: {
        ...statusData,
        permissions: userPermissions
      },
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    Logger.log('ERROR: Error in getDashboardData: ' + error.toString());
    console.error('Error in getDashboardData:', error);
    return {
      success: false,
      error: error.toString(),
      timestamp: new Date().toISOString()
    };
  }
}



/**
 * calculateDaysOpen
 * Calculates the number of days a PR has been open.
 * @param {string|Date} timestamp - The timestamp when the PR was created.
 * @returns {number} Days open.
 */
function calculateDaysOpen(timestamp) {
    const createdDate = new Date(timestamp);
    const currentDate = new Date();
    const diffTime = currentDate - createdDate;
    return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * calculateCompletionPercentage
 * Calculates the completion percentage of a PR based on its fields.
 * @param {Array} row - The row data from the spreadsheet.
 * @returns {number} Completion percentage.
 */
function calculateCompletionPercentage(row) {
    // Example implementation; adjust based on actual criteria
    let completedTasks = 0;
    const totalTasks = 5; // Adjust based on actual number of tasks
    if (row[COL.TASK_1] === 'Completed') completedTasks++;
    if (row[COL.TASK_2] === 'Completed') completedTasks++;
    if (row[COL.TASK_3] === 'Completed') completedTasks++;
    if (row[COL.TASK_4] === 'Completed') completedTasks++;
    if (row[COL.TASK_5] === 'Completed') completedTasks++;
    return (completedTasks / totalTasks) * 100;
}

/**
 * calculateQueuePosition
 * Determines the position of a PR in the processing queue.
 * @param {Array} row - The row data from the spreadsheet.
 * @param {Array} data - All data rows from the spreadsheet.
 * @returns {number} Queue position.
 */
function calculateQueuePosition(row, data) {
    // Example implementation; adjust based on actual queue logic
    const prNumber = row[COL.PR_NUMBER];
    const queue = data.slice(1).filter(r => r[COL.PR_STATUS] === 'In Queue');
    return queue.findIndex(r => r[COL.PR_NUMBER] === prNumber) + 1;
}

/**
 * calculateDaysOverdue - Helper Function
 * --------------------------------------
 * Calculates the number of days a PR is overdue based on the expected landing date.
 * Returns 0 if the PR is not overdue.
 * 
 * @param {Date|string} expectedDate - Expected landing date.
 * @returns {number} - Number of days overdue.
 */
function calculateDaysOverdue(expectedDate) {
    if (!expectedDate) return 0;
    const today = new Date();
    const expected = new Date(expectedDate);
    if (today <= expected) return 0;

    return calculateBusinessDays(expected, today);
}

/**
 * getDetailedBlockingItem - Helper Function
 * ------------------------------------------
 * Determines the current blocking item for an ordered PR based on its attributes.
 * Returns an object with type and message if there's a blocking item, or null otherwise.
 * 
 * @param {Array} row - PR data row.
 * @returns {Object|null} - Blocking item details or null.
 */
function getDetailedBlockingItem(row) {
    if (row[COL.SHIPPED] !== 'Y') {
        return {
            type: 'shipping',
            message: 'Awaiting shipping confirmation',
            days: calculateBusinessDays(row[COL.PAYMENT_DATE], new Date()),
            date: row[COL.PAYMENT_DATE] ? new Date(row[COL.PAYMENT_DATE]).toISOString() : 'No payment date'
        };
    }

    if (row[COL.CUSTOMS_REQUIRED] === 'Y' && row[COL.CUSTOMS_CLEARED] !== 'Y') {
        return {
            type: 'customs',
            message: 'Awaiting customs clearance',
            days: row[COL.CUSTOMS_SUBMISSION_DATE] ? 
                  calculateBusinessDays(row[COL.CUSTOMS_SUBMISSION_DATE], new Date()) : 0,
            date: row[COL.CUSTOMS_SUBMISSION_DATE] ? new Date(row[COL.CUSTOMS_SUBMISSION_DATE]).toISOString() : 'No customs submission date'
        };
    }

    if (row[COL.GOODS_LANDED] !== 'Y') {
        return {
            type: 'delivery',
            message: 'Awaiting delivery confirmation',
            days: row[COL.SHIPMENT_DATE] ? 
                  calculateBusinessDays(row[COL.SHIPMENT_DATE], new Date()) : 0,
            date: row[COL.SHIPMENT_DATE] ? new Date(row[COL.SHIPMENT_DATE]).toISOString() : 'No shipment date'
        };
    }

    return null;
}

/**
 * calculateBusinessDays - Helper Function
 * ----------------------------------------
 * Calculates the number of business days between two dates.
 * 
 * @param {Date|string} startDate - Start date.
 * @param {Date|string} endDate - End date.
 * @returns {number} - Number of business days.
 */
function calculateBusinessDays(startDate, endDate) {
    let count = 0;
    let current = new Date(startDate);
    endDate = new Date(endDate);
    
    while (current <= endDate) {
        const day = current.getDay();
        if (day !== 0 && day !== 6) { // Not Sunday or Saturday
            count++;
        }
        current.setDate(current.getDate() + 1);
    }

    return count;
}

/**
 * calculateQueuePosition - Helper Function
 * ----------------------------------------
 * Calculates the queue position for a PR in the "In Queue" status based on its
 * submission timestamp relative to other PRs in the same status.
 * 
 * @param {Array} row - PR data row.
 * @param {Array} data - All PR data.
 * @returns {number} - Queue position.
 */
function calculateQueuePosition(row, data) {
    // Determine submission date
    const submissionDate = new Date(row[COL.TIMESTAMP]);

    // Filter PRs that are in "In Queue" and submitted before or at the same time
    const queuePRs = data.slice(1).filter(r => {
        return r[COL.PR_STATUS] === 'In Queue' && new Date(r[COL.TIMESTAMP]) <= submissionDate;
    });

    return queuePRs.length;
}

/**
 * sendNotification - Helper Function
 * -----------------------------------
 * Sends notifications (e.g., email) based on specific events or status changes.
 * 
 * @param {string} type - Type of notification (e.g., 'STATUS_CHANGE', 'SUBMISSION').
 * @param {Object} data - Data relevant to the notification.
 * @param {Array<string>} recipients - List of email recipients.
 */
function sendNotification(type, data, recipients) {
    // Implement email or in-app notifications based on 'type'
    // Example: Send email for status changes
    if (type === 'STATUS_CHANGE') {
        const subject = `PR/PO Status Updated to ${data.newStatus}`;
        const body = `
Number: ${data.number}
Description: ${data.description}
Submitted By: ${data.submitter}
Department: ${data.department}
Previous Status: ${data.oldStatus}
New Status: ${data.newStatus}
Notes: ${data.notes || 'N/A'}
View Details: ${data.link}
        `;
        MailApp.sendEmail(recipients.join(','), subject, body);
        Logger.log(`Sent STATUS_CHANGE notification to ${recipients.join(', ')}`);
    }

    if (type === 'SUBMISSION') {
        const subject = `New PR Submitted: ${data.number}`;
        const body = `
A new Purchase Requisition has been submitted.

PR Number: ${data.number}
Description: ${data.description}
Submitted By: ${data.submitter}
Department: ${data.department}
Expected Deadline: ${data.expectedDate}
View Details: ${data.link}
        `;
        MailApp.sendEmail(recipients.join(','), subject, body);
        Logger.log(`Sent SUBMISSION notification to ${recipients.join(', ')}`);
    }

    // Add more notification types as needed
}

/**
 * updatePOStatus - Function to Update PO Status
 * ---------------------------------------------
 * Updates the status of a Purchase Order (PO) and handles related actions such
 * as sending notifications and recalculating queue positions.
 * 
 * @param {string} poNumber - PO number to update.
 * @param {string} newStatus - New status to set.
 * @param {string} notes - Notes related to the status change.
 * @returns {Object} - Result of the update operation.
 */
function updatePOStatus(poNumber, newStatus, notes) {
    Logger.log(`Updating status for PO ${poNumber} to ${newStatus}`);

    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        if (!ss) {
            throw new Error('Failed to open spreadsheet with ID: ' + SPREADSHEET_ID);
        }

        const sheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);
        if (!sheet) {
            throw new Error('Master Log sheet not found in spreadsheet.');
        }

        const data = sheet.getDataRange().getValues();
        Logger.log(`Retrieved ${data.length} rows from Master Log sheet.`);

        // Find the PO row by PO number
        const rowIndex = data.findIndex(row => row[COL.PO_NUMBER_TEXT] === poNumber);
        if (rowIndex === -1) {
            throw new Error(`PO number ${poNumber} not found in Master Log.`);
        }

        Logger.log(`Found PO ${poNumber} at row ${rowIndex + 1}.`);

        // Get current status
        const currentStatus = data[rowIndex][COL.PO_STATUS];
        Logger.log(`Current status of PO ${poNumber}: ${currentStatus}`);

        // Validate status change
        validateStatusChange(currentStatus, newStatus, notes);

        // Update status in the sheet
        sheet.getRange(rowIndex + 1, COL.PO_STATUS + 1).setValue(newStatus);
        Logger.log(`Updated status of PO ${poNumber} to ${newStatus}.`);

        // Add notes if provided
        if (notes) {
            const currentNotes = data[rowIndex][COL.PROCUREMENT_NOTES] || '';
            const updatedNotes = addTimestampToNote(notes, currentNotes);
            sheet.getRange(rowIndex + 1, COL.PROCUREMENT_NOTES + 1).setValue(updatedNotes);
            Logger.log(`Added notes to PO ${poNumber}: ${notes}`);
        }

        // Handle status-specific actions
        handleStatusChangeActions(poNumber, data[rowIndex], currentStatus, newStatus, notes);

        return {
            success: true,
            message: `PO ${poNumber} status updated to ${newStatus}.`
        };

    } catch (error) {
        Logger.log(`ERROR: Error in updatePOStatus: ${error.toString()}`);
        // Use console.error for explicit error logging
        console.error(`Error in updatePOStatus:`, error);
        return {
            success: false,
            error: error.toString()
        };
    }
}

/**
 * validateStatusChange - Helper Function
 * ---------------------------------------
 * Validates whether a status change is allowed based on current status and rules.
 * Throws an error if the status change is invalid.
 * 
 * @param {string} currentStatus - Current status of the PO.
 * @param {string} newStatus - New status to set.
 * @param {string} notes - Notes related to the status change.
 * @throws {Error} - If the status change is invalid.
 */
function validateStatusChange(currentStatus, newStatus, notes) {
    const allowedTransitions = {
        'PO Pending Approval': ['PO Approved', 'Rejected', 'Canceled'],
        'PO Approved': ['PO Ordered', 'Canceled'],
        'PO Ordered': ['Completed', 'Canceled'],
        'Completed': [], // No transitions allowed from Completed
        'Rejected': [], // No transitions allowed from Rejected
        'Canceled': []  // No transitions allowed from Canceled
    };

    // Any status can transition to 'Canceled' except 'Completed', 'Rejected', 'Canceled'
    if (newStatus === 'Canceled') {
        if (['Completed', 'Rejected', 'Canceled'].includes(currentStatus)) {
            throw new Error(`Cannot cancel PO from status '${currentStatus}'.`);
        }
    }

    // Validate if the transition is allowed
    if (!allowedTransitions[currentStatus].includes(newStatus)) {
        throw new Error(`Invalid status change from '${currentStatus}' to '${newStatus}'.`);
    }

    // Validate required notes for certain status changes
    const requiresNotes = ['Rejected', 'Revise and Resubmit'];
    if (requiresNotes.includes(newStatus) && (!notes || notes.trim() === '')) {
        throw new Error(`Notes are required when changing status to '${newStatus}'.`);
    }
}

/**
 * addTimestampToNote - Helper Function
 * -------------------------------------
 * Adds a timestamp to the provided notes.
 * 
 * @param {string} newNote - New note to add.
 * @param {string} existingNotes - Existing notes in the sheet.
 * @returns {string} - Combined notes with timestamp.
 */
function addTimestampToNote(newNote, existingNotes) {
    const timestamp = new Date().toISOString();
    return existingNotes + `\n[${timestamp}] ${newNote}`;
}

/**
 * handleStatusChangeActions - Helper Function
 * --------------------------------------------
 * Handles additional actions required when a PO's status changes, such as sending
 * notifications and updating queue positions.
 * 
 * @param {string} poNumber - PO number.
 * @param {Array} poData - PO data row.
 * @param {string} oldStatus - Previous status.
 * @param {string} newStatus - New status.
 * @param {string} notes - Notes related to the status change.
 */
function handleStatusChangeActions(poNumber, poData, oldStatus, newStatus, notes) {
    Logger.log(`Handling actions for PO ${poNumber} status change from '${oldStatus}' to '${newStatus}'.`);

    // Update queue positions if needed
    if (oldStatus === 'PO Pending Approval' || newStatus === 'PO Pending Approval') {
        updateQueuePositions();
    }

    // Prepare notification data
    const notificationData = {
        number: poNumber,
        description: poData[COL.DESCRIPTION],
        submitter: poData[COL.REQUESTOR_NAME],
        newStatus: newStatus,
        oldStatus: oldStatus,
        notes: notes,
        link: generateViewLink(poNumber)
    };

    // Define recipients based on PO data (ensure these columns exist)
    const recipients = [
        CONFIG.PROCUREMENT_EMAIL,
        poData[COL.EMAIL],            // COL.EMAIL is defined as 2 in Code.gs
        poData[COL.APPROVER_EMAIL]    // COL.APPROVER_EMAIL is defined as 28 in Code.gs
    ];

    // Send notification
    sendNotification('STATUS_CHANGE', notificationData, recipients);

    // Handle auto-cancellation reset if applicable
    if (newStatus === 'PO Ordered') {
        resetAutoCancellationTimer(poNumber);
    }
}

/**
 * updateQueuePositions - Helper Function
 * ---------------------------------------
 * Updates the queue positions for PRs or POs as necessary.
 * 
 * Note: Implement the actual logic based on your system's requirements.
 */
function updateQueuePositions() {
    Logger.log('Updating queue positions.');
    // Implement queue position update logic here
    // This could involve recalculating positions based on submission times or priorities
}

/**
 * resetAutoCancellationTimer - Helper Function
 * --------------------------------------------
 * Resets the auto-cancellation timer for a PO when it reaches a certain status.
 * 
 * @param {string} poNumber - PO number.
 */
function resetAutoCancellationTimer(poNumber) {
    Logger.log(`Resetting auto-cancellation timer for PO ${poNumber}.`);
    // Implement auto-cancellation logic here
    // This could involve setting a time-based trigger to auto-cancel the PO if not completed
}

/**
 * filterAndFormatPRData - Helper Function
 * ----------------------------------------
 * Filters and formats PR data for display on the dashboard.
 * 
 * @param {Array} data - Raw PR data from the sheet.
 * @param {string} organization - Organization filter.
 * @param {string} status - Status filter.
 * @returns {Array<Object>} - Filtered and formatted PR data.
 */
function filterAndFormatPRData(data, organization, status) {
    return data
        .filter(row => {
            // Apply organization filter
            if (organization && row[COL.ORGANIZATION] !== organization) return false;
            
            // Apply status filter
            if (status && row[COL.PR_STATUS] !== status) return false;
            
            return true;
        })
        .map(row => {
            // Format data for display
            const prData = {
                prNumber: row[COL.PR_NUMBER],
                timestamp: row[COL.TIMESTAMP],
                description: row[COL.DESCRIPTION],
                requestor: row[COL.REQUESTOR_NAME],
                department: row[COL.DEPARTMENT],
                daysOpen: calculateDaysOpen(row[COL.TIMESTAMP]),
                urgency: row[COL.URGENCY] === 'Y' ? 'Urgent' : 'Normal',
                amount: row[COL.PR_AMOUNT],
                currency: row[COL.CURRENCY],
                vendor: row[COL.VENDOR],
                approver: row[COL.APPROVER],
                completion: calculateCompletionPercentage(row),
                procurementNotes: row[COL.PROCUREMENT_NOTES]
            };

            // Add status-specific fields
            switch (row[COL.PR_STATUS]) {
                case 'Submitted':
                    prData.submittedDate = row[COL.TIMESTAMP];
                    break;

                case 'In Queue':
                    prData.queuePosition = calculateQueuePosition(row, data);
                    prData.quotesRequired = row[COL.QUOTES_REQUIRED] === 'Y';
                    prData.quotesStatus = row[COL.QUOTES_LINK] ? 'Received' : 'Pending';
                    prData.adjudicationRequired = row[COL.ADJ_REQUIRED] === 'Y';
                    prData.adjudicationStatus = row[COL.ADJ_NOTES] ? 'Completed' : 'Pending';
                    break;

                case 'Ordered':
                    prData.orderedDate = row[COL.PO_DATE];
                    prData.expectedLanding = row[COL.EXPECTED_LANDING_DATE];
                    prData.daysOverdue = calculateDaysOverdue(row[COL.EXPECTED_LANDING_DATE]);
                    prData.blockingItem = getDetailedBlockingItem(row);
                    prData.paymentDate = row[COL.PAYMENT_DATE];
                    prData.shipmentStatus = {
                        shipped: row[COL.SHIPPED] === 'Y',
                        shipmentDate: row[COL.SHIPMENT_DATE],
                        customsRequired: row[COL.CUSTOMS_REQUIRED] === 'Y',
                        customsCleared: row[COL.CUSTOMS_CLEARED] === 'Y',
                        customsClearanceDate: row[COL.DATE_CLEARED],
                        goodsLanded: row[COL.GOODS_LANDED] === 'Y',
                        landedDate: row[COL.LANDED_DATE]
                    };
                    break;

                case 'R&R':
                    prData.rAndRDate = row[COL.OVERRIDE_DATE] || new Date();
                    prData.rAndRBy = row[COL.OVERRIDE_BY] || '';
                    prData.rAndRJustification = row[COL.OVERRIDE_JUST] || '';
                    break;

                case 'Completed':
                    prData.completedDate = row[COL.LANDED_DATE];
                    prData.timeToComplete = calculateBusinessDays(row[COL.TIMESTAMP], row[COL.LANDED_DATE]);
                    prData.timeToShip = row[COL.TIME_TO_SHIP] || null;
                    prData.timeInCustoms = row[COL.TIME_IN_CUSTOMS] || null;
                    prData.timeToLand = row[COL.TIME_TO_LAND] || null;
                    break;

                case 'Rejected':
                    prData.rejectionDate = row[COL.OVERRIDE_DATE];
                    prData.rejectedBy = row[COL.OVERRIDE_BY];
                    prData.rejectionReason = row[COL.OVERRIDE_JUST];
                    break;

                case 'Canceled':
                    prData.cancelationDate = row[COL.OVERRIDE_DATE];
                    prData.canceledBy = row[COL.OVERRIDE_BY];
                    prData.cancelationReason = row[COL.OVERRIDE_JUST];
                    break;
            }

            return prData;
        });
}

/**
 * calculateDashboardMetrics - Helper Function
 * -------------------------------------------
 * Calculates various metrics for the dashboard based on PR and PO data.
 * 
 * @param {Array} data - Raw data from the sheet.
 * @param {string} organization - Organization filter.
 * @returns {Object} - Dashboard metrics.
 */
function calculateDashboardMetrics(data, organization) {
    Logger.log('Calculating dashboard metrics for organization: ' + organization);

    const metrics = {
        totalPRs: 0,
        urgentPRs: 0,
        avgDaysOpen: 0,
        overduePRs: 0,
        quotesRequired: 0,
        adjudicationRequired: 0,
        customsClearanceRequired: 0,
        completionRate: 0
    };

    let totalDaysOpen = 0;
    let activePRs = 0;
    let completedPRs = 0;
    let totalCompletionPercentage = 0;

    data.slice(1).forEach(row => {
        // Apply organization filter
        if (!organization || row[COL.ORGANIZATION] === organization) {
            const status = row[COL.PR_STATUS];

            // Only process if status is recognized
            if (statusData.countsByStatus.hasOwnProperty(status)) {
                metrics.totalPRs++;
            }

            // Calculate metrics based on PR attributes
            if (row[COL.URGENCY] === 'Y') {
                metrics.urgentPRs++;
            }
            if (row[COL.QUOTES_REQUIRED] === 'Y') {
                metrics.quotesRequired++;
            }
            if (row[COL.ADJ_REQUIRED] === 'Y') {
                metrics.adjudicationRequired++;
            }
            if (row[COL.CUSTOMS_REQUIRED] === 'Y') {
                metrics.customsClearanceRequired++;
            }

            // Calculate days open
            const daysOpen = calculateDaysOpen(row[COL.TIMESTAMP]);
            totalDaysOpen += daysOpen;
            activePRs++;

            // Calculate completion percentage
            const completion = calculateCompletionPercentage(row);
            totalCompletionPercentage += completion;

            // Track completed PRs
            if (status === 'Completed') {
                completedPRs++;
            }

            // Track overdue PRs
            if (status === 'Ordered' && calculateDaysOverdue(row[COL.EXPECTED_LANDING_DATE]) > 0) {
                metrics.overduePRs++;
            }
        }
    });

    // Calculate average days open and completion rate
    if (activePRs > 0) {
        metrics.avgDaysOpen = Math.round(totalDaysOpen / activePRs);
        metrics.completionRate = Math.round(totalCompletionPercentage / activePRs);
    }

    Logger.log('Updated dashboard metrics: ' + JSON.stringify(metrics));
    return metrics;
}

/**
 * viewPRDetails - Function to Navigate to PR Details
 * ---------------------------------------------------
 * Redirects the user to the detailed view of a specific PR.
 * 
 * @param {string} prNumber - PR number to view.
 */
function viewPRDetails(prNumber) {
    // This function is intended to be called from the client-side
    // Implement navigation logic as needed, possibly by setting a redirect URL
}

/**
 * getActiveOrganizations - Helper Function
 * ----------------------------------------
 * Retrieves a list of active organizations from the Organizations sheet.
 * 
 * @returns {Array<string>} - List of active organizations.
 */
function getActiveOrganizations() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        if (!ss) {
            throw new Error('Failed to open spreadsheet with ID: ' + SPREADSHEET_ID);
        }

        const sheet = ss.getSheetByName(CONFIG.ORGANIZATIONS_TAB);
        if (!sheet) {
            throw new Error('Organizations sheet not found.');
        }

        const data = sheet.getDataRange().getValues();
        Logger.log(`Retrieved ${data.length} rows from Organizations sheet.`);

        // Assuming organizations are listed in the first column, excluding header
        const organizations = data.slice(1).map(row => row[0]).filter(org => org && org.trim() !== '');
        Logger.log(`Active organizations: ${organizations.join(', ')}`);

        return organizations;
    } catch (error) {
        Logger.log(`ERROR: Error in getActiveOrganizations: ${error.toString()}`);
        console.error('Error in getActiveOrganizations:', error);
        return [];
    }
}