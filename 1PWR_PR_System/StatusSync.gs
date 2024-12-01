

/**
 * Synchronizes Submitted sheet with Master Log data
 * Headers: PR Number, Description, Submitted By, Submitted Date, 
 *          Resubmitted Date, Days Open, Days Since Resubmission, Link
 */
function syncSubmittedSheet() {
    console.log('Starting Submitted sheet synchronization');
    
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        
        // Get Master Log data
        const masterSheet = ss.getSheetByName(CONFIG.MASTER_LOG_TAB);
        if (!masterSheet) throw new Error('Master Log sheet not found');
        
        // Get or create Submitted sheet
        let submittedSheet = ss.getSheetByName('Submitted');
        if (!submittedSheet) {
            submittedSheet = ss.insertSheet('Submitted');
            console.log('Created new Submitted sheet');
        }
        
        // Headers are correct - no change needed
        const headers = [
            'PR Number',
            'Description',
            'Submitted By',
            'Submitted Date',
            'Resubmitted Date',
            'Days Open',
            'Days Since Resubmission',
            'Link'
        ];
        
        const masterData = masterSheet.getDataRange().getValues();
        const masterHeaders = masterData[0];
        
        // Update column indices to use COL constant for consistency
        const colIndices = {
            prNumber: COL.PR_NUMBER,
            description: COL.DESCRIPTION,
            submittedBy: COL.REQUESTOR_NAME,
            submittedDate: COL.TIMESTAMP,
            resubmittedDate: COL.OVERRIDE_DATE,
            status: COL.PR_STATUS
        };
        
        // Validation stays the same
        for (const [key, index] of Object.entries(colIndices)) {
            if (index === -1) throw new Error(`Column not found in Master Log: ${key}`);
        }
        
        // Update the data mapping to fix Days Since Resubmission calculation
        const submittedData = masterData.slice(1) // Skip header row
            .filter(row => row[colIndices.status] === 'Submitted')
            .map(row => {
                const submittedDate = row[colIndices.submittedDate];
                const resubmittedDate = row[colIndices.resubmittedDate];
                
                // Calculate days since resubmission if applicable
                let daysSinceResubmission = '';
                if (resubmittedDate) {
                    const now = new Date();
                    daysSinceResubmission = Math.floor(
                        (now - new Date(resubmittedDate)) / (1000 * 60 * 60 * 24)
                    );
                }
                
                return [
                    row[colIndices.prNumber],
                    row[colIndices.description],
                    row[colIndices.submittedBy],
                    submittedDate,
                    resubmittedDate || '',  // Ensure empty string if null
                    calculateDaysOpen(submittedDate),
                    daysSinceResubmission,  // Updated calculation
                    generateViewLink(row[colIndices.prNumber])
                ];
            });
            
        // Clear and set headers
        submittedSheet.clear();
        submittedSheet.getRange(1, 1, 1, headers.length)
            .setValues([headers])
            .setFontWeight('bold')
            .setBackground('#E8EAED');
        
        // Write data if exists
        if (submittedData.length > 0) {
            submittedSheet.getRange(2, 1, submittedData.length, headers.length)
                .setValues(submittedData);
                
            // Format date columns
            const dateColumns = [4, 5]; // Submitted Date and Resubmitted Date columns
            dateColumns.forEach(col => {
                submittedSheet.getRange(2, col, submittedData.length, 1)
                    .setNumberFormat('yyyy-mm-dd hh:mm:ss');
            });
            
            // Format numeric columns
            const numericColumns = [6, 7]; // Days Open and Days Since Resubmission
            numericColumns.forEach(col => {
                submittedSheet.getRange(2, col, submittedData.length, 1)
                    .setNumberFormat('#,##0');
            });
        }
        
        // Auto-size columns
        submittedSheet.autoResizeColumns(1, headers.length);
        
        // Freeze header row
        submittedSheet.setFrozenRows(1);
        
        console.log(`Synchronized Submitted sheet with ${submittedData.length} records`);
        return true;
        
    } catch (error) {
        console.error('Error synchronizing Submitted sheet:', error);
        throw error;
    }
}

// Add helper function if not already present
function calculateDaysOpen(startDate) {
    if (!startDate) return 0;
    const start = new Date(startDate);
    const now = new Date();
    return Math.floor((now - start) / (1000 * 60 * 60 * 24));
}

/**
 * Sets up time-based trigger for synchronization
 */
function setupSyncTrigger() {
    // Delete any existing sync triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'syncSubmittedSheet') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    
    // Create new trigger to run every minute
    ScriptApp.newTrigger('syncSubmittedSheet')
        .timeBased()
        .everyMinutes(1)
        .create();
    
    console.log('Sync trigger set up successfully');
}

/**
 * Sets up trigger to sync on edit of Master Log
 */
function setupEditTrigger() {
    // Delete any existing edit triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onMasterLogEdit') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    
    // Create new edit trigger
    ScriptApp.newTrigger('onMasterLogEdit')
        .forSpreadsheet(SPREADSHEET_ID)
        .onEdit()
        .create();
    
    console.log('Edit trigger set up successfully');
}

/**
 * Handles edits to Master Log
 */
function onMasterLogEdit(e) {
    if (!e || !e.range) return;
    
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.MASTER_LOG_TAB) return;
    
    // Get the edited row data
    const row = e.range.getRow();
    const masterData = sheet.getDataRange().getValues();
    const headers = masterData[0];
    
    // Check if status column was edited
    const statusCol = headers.indexOf('PR Status');
    if (statusCol === -1) return;
    
    // If status is or was 'Submitted', sync the Submitted sheet
    const oldValue = e.oldValue;
    const newValue = e.value;
    if (oldValue === 'Submitted' || newValue === 'Submitted') {
        console.log('Status change detected affecting Submitted items, syncing...');
        syncSubmittedSheet();
    }
}

/**
 * Function to manually trigger sync for testing
 */
function testSubmittedSync() {
    try {
        syncSubmittedSheet();
        console.log('Manual sync completed successfully');
    } catch (error) {
        console.error('Manual sync failed:', error);
    }
}