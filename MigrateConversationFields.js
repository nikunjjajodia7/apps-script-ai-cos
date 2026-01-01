/**
 * Migration Script: Add Conversation and Approval Fields to Existing Tasks_DB
 * 
 * This script backfills existing tasks with default values for the new conversation
 * and approval tracking fields added in the interaction system update.
 * 
 * Run this ONCE after updating the schema to initialize existing tasks.
 */

function migrateConversationFields() {
  try {
    Logger.log('=== Starting Conversation Fields Migration ===');
    
    const spreadsheet = getSpreadsheet();
    const tasksSheet = spreadsheet.getSheetByName('Tasks_DB');
    
    if (!tasksSheet) {
      Logger.log('ERROR: Tasks_DB sheet not found. Run createAllSheets() first.');
      return;
    }
    
    // Get all headers
    const headers = tasksSheet.getRange(1, 1, 1, tasksSheet.getLastColumn()).getValues()[0];
    
    // Check if new columns exist
    const newFields = [
      'Conversation_History',
      'Approval_State',
      'Pending_Decision',
      'Progress_Update',
      'Progress_Percentage',
      'Last_Progress_Update',
      'Last_Boss_Message',
      'Last_Employee_Message',
      'Message_Count',
      'Negotiation_History'
    ];
    
    const missingFields = newFields.filter(field => headers.indexOf(field) === -1);
    
    if (missingFields.length > 0) {
      Logger.log(`ERROR: Missing columns: ${missingFields.join(', ')}`);
      Logger.log('Please run createAllSheets() first to add the new columns.');
      return;
    }
    
    // Get column indices
    const conversationHistoryIndex = headers.indexOf('Conversation_History');
    const approvalStateIndex = headers.indexOf('Approval_State');
    const pendingDecisionIndex = headers.indexOf('Pending_Decision');
    const progressUpdateIndex = headers.indexOf('Progress_Update');
    const progressPercentageIndex = headers.indexOf('Progress_Percentage');
    const lastProgressUpdateIndex = headers.indexOf('Last_Progress_Update');
    const lastBossMessageIndex = headers.indexOf('Last_Boss_Message');
    const lastEmployeeMessageIndex = headers.indexOf('Last_Employee_Message');
    const messageCountIndex = headers.indexOf('Message_Count');
    const negotiationHistoryIndex = headers.indexOf('Negotiation_History');
    
    // Get all task data (skip header row)
    const lastRow = tasksSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No tasks to migrate.');
      return;
    }
    
    const dataRange = tasksSheet.getRange(2, 1, lastRow - 1, tasksSheet.getLastColumn());
    const data = dataRange.getValues();
    
    let migratedCount = 0;
    let skippedCount = 0;
    
    // Process each task
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const taskId = row[headers.indexOf('Task_ID')];
      const status = row[headers.indexOf('Status')];
      const employeeReply = row[headers.indexOf('Employee_Reply')] || '';
      const interactionLog = row[headers.indexOf('Interaction_Log')] || '';
      
      // Skip if already migrated (has non-empty Conversation_History or Approval_State)
      const existingConversationHistory = row[conversationHistoryIndex];
      const existingApprovalState = row[approvalStateIndex];
      
      if (existingConversationHistory || existingApprovalState) {
        skippedCount++;
        continue;
      }
      
      // Initialize conversation history from Employee_Reply if exists
      let conversationHistory = [];
      if (employeeReply) {
        conversationHistory.push({
          id: Utilities.getUuid(),
          timestamp: new Date().toISOString(),
          senderEmail: row[headers.indexOf('Assignee_Email')] || '',
          senderName: row[headers.indexOf('Assignee_Name')] || '',
          type: 'employee_reply',
          content: employeeReply,
          metadata: {}
        });
      }
      
      // Determine initial approval state based on status
      let approvalState = APPROVAL_STATE.NONE;
      if (status === TASK_STATUS.REVIEW_DATE || 
          status === TASK_STATUS.REVIEW_SCOPE || 
          status === TASK_STATUS.REVIEW_ROLE) {
        approvalState = APPROVAL_STATE.AWAITING_BOSS;
      }
      
      // Update row with default values
      const rowNumber = i + 2; // +2 because we start from row 2 (skip header)
      
      tasksSheet.getRange(rowNumber, conversationHistoryIndex + 1).setValue(
        conversationHistory.length > 0 ? JSON.stringify(conversationHistory) : ''
      );
      tasksSheet.getRange(rowNumber, approvalStateIndex + 1).setValue(approvalState);
      tasksSheet.getRange(rowNumber, pendingDecisionIndex + 1).setValue('');
      tasksSheet.getRange(rowNumber, progressUpdateIndex + 1).setValue('');
      tasksSheet.getRange(rowNumber, progressPercentageIndex + 1).setValue('');
      tasksSheet.getRange(rowNumber, lastProgressUpdateIndex + 1).setValue('');
      tasksSheet.getRange(rowNumber, lastBossMessageIndex + 1).setValue('');
      tasksSheet.getRange(rowNumber, lastEmployeeMessageIndex + 1).setValue(
        employeeReply ? new Date().toISOString() : ''
      );
      tasksSheet.getRange(rowNumber, messageCountIndex + 1).setValue(conversationHistory.length);
      tasksSheet.getRange(rowNumber, negotiationHistoryIndex + 1).setValue('');
      
      migratedCount++;
      
      if (migratedCount % 10 === 0) {
        Logger.log(`Migrated ${migratedCount} tasks...`);
      }
    }
    
    Logger.log(`=== Migration Complete ===`);
    Logger.log(`Migrated: ${migratedCount} tasks`);
    Logger.log(`Skipped: ${skippedCount} tasks (already migrated)`);
    Logger.log(`Total: ${data.length} tasks`);
    
  } catch (error) {
    Logger.log(`ERROR in migrateConversationFields: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
    throw error;
  }
}

/**
 * Preview migration without making changes
 * Shows what would be migrated
 */
function previewConversationFieldsMigration() {
  try {
    Logger.log('=== Preview: Conversation Fields Migration ===');
    
    const spreadsheet = getSpreadsheet();
    const tasksSheet = spreadsheet.getSheetByName('Tasks_DB');
    
    if (!tasksSheet) {
      Logger.log('ERROR: Tasks_DB sheet not found.');
      return;
    }
    
    const headers = tasksSheet.getRange(1, 1, 1, tasksSheet.getLastColumn()).getValues()[0];
    
    const newFields = [
      'Conversation_History',
      'Approval_State',
      'Pending_Decision'
    ];
    
    const missingFields = newFields.filter(field => headers.indexOf(field) === -1);
    
    if (missingFields.length > 0) {
      Logger.log(`Missing columns: ${missingFields.join(', ')}`);
      Logger.log('Run createAllSheets() first.');
      return;
    }
    
    const lastRow = tasksSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No tasks to migrate.');
      return;
    }
    
    const data = tasksSheet.getRange(2, 1, lastRow - 1, tasksSheet.getLastColumn()).getValues();
    
    let needsMigration = 0;
    let alreadyMigrated = 0;
    
    const conversationHistoryIndex = headers.indexOf('Conversation_History');
    const approvalStateIndex = headers.indexOf('Approval_State');
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const taskId = row[headers.indexOf('Task_ID')];
      const existingConversationHistory = row[conversationHistoryIndex];
      const existingApprovalState = row[approvalStateIndex];
      
      if (existingConversationHistory || existingApprovalState) {
        alreadyMigrated++;
      } else {
        needsMigration++;
        if (needsMigration <= 5) {
          Logger.log(`Would migrate: ${taskId}`);
        }
      }
    }
    
    if (needsMigration > 5) {
      Logger.log(`... and ${needsMigration - 5} more tasks`);
    }
    
    Logger.log(`\nSummary:`);
    Logger.log(`  Needs migration: ${needsMigration}`);
    Logger.log(`  Already migrated: ${alreadyMigrated}`);
    Logger.log(`  Total: ${data.length}`);
    
  } catch (error) {
    Logger.log(`ERROR in preview: ${error.toString()}`);
  }
}

