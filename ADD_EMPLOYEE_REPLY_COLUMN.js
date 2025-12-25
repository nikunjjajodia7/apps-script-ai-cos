/**
 * Migration Script: Add Employee_Reply Column to Existing Tasks_DB Sheet
 * 
 * Run this ONCE if your Tasks_DB sheet already exists and doesn't have the Employee_Reply column.
 * This will add the column in the correct position (after Boss_Reply_Draft, before Created_Date).
 */
function ADD_EMPLOYEE_REPLY_COLUMN() {
  try {
    Logger.log('=== Adding Employee_Reply Column to Tasks_DB ===');
    
    const spreadsheet = getSpreadsheet();
    const tasksSheet = spreadsheet.getSheetByName('Tasks_DB');
    
    if (!tasksSheet) {
      Logger.log('ERROR: Tasks_DB sheet not found. Run createAllSheets() first.');
      return;
    }
    
    // Get all headers
    const headers = tasksSheet.getRange(1, 1, 1, tasksSheet.getLastColumn()).getValues()[0];
    
    // Check if Employee_Reply column already exists
    const employeeReplyIndex = headers.indexOf('Employee_Reply');
    if (employeeReplyIndex !== -1) {
      Logger.log('✓ Employee_Reply column already exists at column ' + (employeeReplyIndex + 1));
      Logger.log('No action needed.');
      return;
    }
    
    // Find the position where Employee_Reply should be inserted
    // It should be after Boss_Reply_Draft and before Created_Date
    const bossReplyIndex = headers.indexOf('Boss_Reply_Draft');
    const createdDateIndex = headers.indexOf('Created_Date');
    
    let insertColumn = createdDateIndex + 1; // Default: before Created_Date
    
    if (bossReplyIndex !== -1) {
      insertColumn = bossReplyIndex + 2; // After Boss_Reply_Draft
    } else if (createdDateIndex === -1) {
      // If neither exists, add at the end
      insertColumn = headers.length + 1;
    }
    
    Logger.log(`Inserting Employee_Reply column at position ${insertColumn} (column ${String.fromCharCode(64 + insertColumn)})`);
    
    // Insert a new column
    tasksSheet.insertColumnAfter(insertColumn - 1);
    
    // Set the header
    tasksSheet.getRange(1, insertColumn).setValue('Employee_Reply');
    tasksSheet.getRange(1, insertColumn).setFontWeight('bold');
    
    Logger.log('✓ Employee_Reply column added successfully!');
    Logger.log('Column position: ' + insertColumn + ' (' + String.fromCharCode(64 + insertColumn) + ')');
    
    // Verify the column was added
    const newHeaders = tasksSheet.getRange(1, 1, 1, tasksSheet.getLastColumn()).getValues()[0];
    const verifyIndex = newHeaders.indexOf('Employee_Reply');
    if (verifyIndex !== -1) {
      Logger.log('✓ Verification: Employee_Reply column confirmed at position ' + (verifyIndex + 1));
    } else {
      Logger.log('✗ Warning: Could not verify Employee_Reply column was added');
    }
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

