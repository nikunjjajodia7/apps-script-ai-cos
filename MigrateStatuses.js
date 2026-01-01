/**
 * Database Migration Script v2.0
 * Migrates Tasks_DB to new bucket and conversation state system
 * 
 * MIGRATION STEPS:
 * 1. Add new columns (Conversation_State, Pending_Changes, AI_Summary, Initial_Parameters, Conversation_History)
 * 2. Migrate old status values to new lifecycle statuses
 * 3. Set initial Conversation_State based on old status
 * 4. Create Initial_Parameters snapshot for existing tasks
 * 
 * Run migrations in order:
 * 1. runFullMigration() - Does everything in one go (recommended)
 * OR run individual steps:
 * 1. addNewColumns() - Add missing columns
 * 2. migrateTaskStatuses() - Migrate status values
 * 3. setInitialConversationStates() - Set conversation states
 * 4. createInitialParametersSnapshots() - Create parameter snapshots
 */

// ============================================
// FULL MIGRATION (RECOMMENDED)
// ============================================

/**
 * Run the complete migration - adds columns, migrates statuses, and sets conversation states
 * Run this from Apps Script editor: Run > runFullMigration
 */
function runFullMigration() {
  Logger.log('╔════════════════════════════════════════════════════════════╗');
  Logger.log('║       FULL DATABASE MIGRATION TO NEW BUCKET SYSTEM         ║');
  Logger.log('╚════════════════════════════════════════════════════════════╝');
  Logger.log(`Started at: ${new Date().toISOString()}\n`);
  
  try {
    // Step 1: Add new columns
    Logger.log('━━━ Step 1/4: Adding New Columns ━━━');
    const columnsResult = addNewColumns();
    Logger.log(`Columns added: ${columnsResult.added.join(', ') || 'None needed'}\n`);
    
    // Step 2: Migrate statuses
    Logger.log('━━━ Step 2/4: Migrating Status Values ━━━');
    const statusResult = migrateTaskStatuses();
    Logger.log(`Migrated: ${statusResult.migrated}, Skipped: ${statusResult.skipped}\n`);
    
    // Step 3: Set conversation states
    Logger.log('━━━ Step 3/4: Setting Conversation States ━━━');
    const convResult = setInitialConversationStates();
    Logger.log(`Set: ${convResult.set}, Skipped: ${convResult.skipped}\n`);
    
    // Step 4: Create initial parameters snapshots
    Logger.log('━━━ Step 4/4: Creating Initial Parameters Snapshots ━━━');
    const paramsResult = createInitialParametersSnapshots();
    Logger.log(`Created: ${paramsResult.created}, Skipped: ${paramsResult.skipped}\n`);
    
    Logger.log('╔════════════════════════════════════════════════════════════╗');
    Logger.log('║                 MIGRATION COMPLETE                          ║');
    Logger.log('╚════════════════════════════════════════════════════════════╝');
    Logger.log(`Finished at: ${new Date().toISOString()}`);
    
    return {
      success: true,
      columns: columnsResult,
      statuses: statusResult,
      conversationStates: convResult,
      initialParams: paramsResult
    };
    
  } catch (error) {
    Logger.log(`\n❌ FATAL ERROR: ${error.toString()}`);
    Logger.log(error.stack);
    return { success: false, error: error.toString() };
  }
}

// ============================================
// STEP 1: ADD NEW COLUMNS
// ============================================

/**
 * Add new columns to Tasks_DB sheet if they don't exist
 */
function addNewColumns() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID not found in Script Properties. Please run quickSetup() first.');
  }
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(SHEETS.TASKS_DB);
  
  if (!sheet) {
    throw new Error('Tasks_DB sheet not found');
  }
  
  // Get existing headers
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log(`Existing columns: ${headers.length}`);
  
  // New columns to add
  const newColumns = [
    'Conversation_State',
    'Pending_Changes',
    'AI_Summary',
    'Initial_Parameters',
    'Conversation_History'
  ];
  
  const added = [];
  
  newColumns.forEach(colName => {
    if (!headers.includes(colName)) {
      // Add column at the end
      const newColIndex = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColIndex).setValue(colName);
      added.push(colName);
      Logger.log(`  ✓ Added column: ${colName}`);
    } else {
      Logger.log(`  - Column exists: ${colName}`);
    }
  });
  
  return { added, existing: headers.length };
}

// ============================================
// STEP 2: MIGRATE STATUS VALUES
// ============================================

/**
 * Main status migration function - run this from Apps Script editor
 * Go to Run > migrateTaskStatuses
 */
function migrateTaskStatuses() {
  try {
    Logger.log('=== Starting Task Status Migration ===');
    Logger.log(`Migration started at: ${new Date().toISOString()}`);
    
    const tasks = getSheetData(SHEETS.TASKS_DB);
    Logger.log(`Found ${tasks.length} tasks to process`);
    
    let migratedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    
    const migrationLog = [];
    
    tasks.forEach((task, index) => {
      const taskId = task.Task_ID;
      const oldStatus = task.Status;
      
      if (!taskId) {
        Logger.log(`Skipping row ${index + 2}: No Task_ID`);
        skippedCount++;
        return;
      }
      
      // Check if status is already in new format
      const newStatuses = Object.values(TASK_STATUS);
      if (newStatuses.includes(oldStatus)) {
        Logger.log(`Task ${taskId}: Already using new status "${oldStatus}" - skipping`);
        skippedCount++;
        return;
      }
      
      // Get new status from legacy mapping
      const newStatus = normalizeStatus(oldStatus);
      
      if (newStatus !== oldStatus) {
        try {
          updateTask(taskId, { Status: newStatus });
          migratedCount++;
          migrationLog.push({
            taskId: taskId,
            taskName: task.Task_Name,
            oldStatus: oldStatus,
            newStatus: newStatus,
            success: true
          });
          Logger.log(`Task ${taskId}: "${oldStatus}" → "${newStatus}"`);
        } catch (error) {
          errorCount++;
          migrationLog.push({
            taskId: taskId,
            taskName: task.Task_Name,
            oldStatus: oldStatus,
            newStatus: newStatus,
            success: false,
            error: error.toString()
          });
          Logger.log(`ERROR migrating task ${taskId}: ${error.toString()}`);
        }
      } else {
        skippedCount++;
        Logger.log(`Task ${taskId}: Status "${oldStatus}" unchanged`);
      }
    });
    
    Logger.log('');
    Logger.log('=== Migration Complete ===');
    Logger.log(`Total tasks: ${tasks.length}`);
    Logger.log(`Migrated: ${migratedCount}`);
    Logger.log(`Skipped: ${skippedCount}`);
    Logger.log(`Errors: ${errorCount}`);
    Logger.log('');
    Logger.log('Migration Log:');
    Logger.log(JSON.stringify(migrationLog, null, 2));
    
    return {
      success: true,
      total: tasks.length,
      migrated: migratedCount,
      skipped: skippedCount,
      errors: errorCount,
      log: migrationLog
    };
    
  } catch (error) {
    Logger.log(`FATAL ERROR: ${error.toString()}`);
    Logger.log(error.stack);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Dry run - see what would be migrated without making changes
 */
function previewStatusMigration() {
  Logger.log('=== Status Migration Preview (DRY RUN) ===');
  Logger.log('No changes will be made - just showing what would happen\n');
  
  const tasks = getSheetData(SHEETS.TASKS_DB);
  Logger.log(`Found ${tasks.length} tasks`);
  
  const statusCounts = {};
  const migrations = [];
  
  tasks.forEach(task => {
    const oldStatus = task.Status || 'Unknown';
    
    // Count current statuses
    statusCounts[oldStatus] = (statusCounts[oldStatus] || 0) + 1;
    
    // Check if migration needed
    const newStatuses = Object.values(TASK_STATUS);
    if (!newStatuses.includes(oldStatus)) {
      const newStatus = normalizeStatus(oldStatus);
      if (newStatus !== oldStatus) {
        migrations.push({
          taskId: task.Task_ID,
          taskName: task.Task_Name?.substring(0, 30) + '...',
          from: oldStatus,
          to: newStatus
        });
      }
    }
  });
  
  Logger.log('\n=== Current Status Distribution ===');
  Object.entries(statusCounts).forEach(([status, count]) => {
    const isNew = Object.values(TASK_STATUS).includes(status);
    Logger.log(`  ${status}: ${count} ${isNew ? '(new)' : '(legacy - will migrate)'}`);
  });
  
  Logger.log('\n=== Migrations Required ===');
  if (migrations.length === 0) {
    Logger.log('No migrations needed - all tasks are using new statuses!');
  } else {
    Logger.log(`${migrations.length} tasks need migration:\n`);
    migrations.forEach(m => {
      Logger.log(`  ${m.taskId}: "${m.from}" → "${m.to}" (${m.taskName})`);
    });
  }
  
  return {
    statusCounts,
    migrationsNeeded: migrations.length,
    migrations
  };
}

/**
 * Get statistics on current task statuses
 */
function getStatusStatistics() {
  const tasks = getSheetData(SHEETS.TASKS_DB);
  
  const byStatus = {};
  const byBucket = {
    'ai_assist': 0,
    'pending_action': 0,
    'review': 0,
    'active': 0,
    'done_pending': 0,
    'on_hold': 0,
    'closed': 0,
    'legacy': 0
  };
  
  tasks.forEach(task => {
    const status = task.Status || 'Unknown';
    byStatus[status] = (byStatus[status] || 0) + 1;
    
    // Map to bucket
    switch(status) {
      case TASK_STATUS.AI_ASSIST:
        byBucket['ai_assist']++;
        break;
      case TASK_STATUS.NOT_ACTIVE:
      case TASK_STATUS.PENDING_ACTION:
        byBucket['pending_action']++;
        break;
      case TASK_STATUS.REVIEW_DATE:
      case TASK_STATUS.REVIEW_SCOPE:
      case TASK_STATUS.REVIEW_ROLE:
        byBucket['review']++;
        break;
      case TASK_STATUS.ON_TIME:
      case TASK_STATUS.SLOW_PROGRESS:
        byBucket['active']++;
        break;
      case TASK_STATUS.COMPLETED:
        byBucket['done_pending']++;
        break;
      case TASK_STATUS.ON_HOLD:
      case TASK_STATUS.SOMEDAY:
        byBucket['on_hold']++;
        break;
      case TASK_STATUS.CLOSED:
        byBucket['closed']++;
        break;
      default:
        byBucket['legacy']++;
    }
  });
  
  Logger.log('=== Task Status Statistics ===\n');
  Logger.log('By Status:');
  Object.entries(byStatus).sort((a, b) => b[1] - a[1]).forEach(([s, c]) => {
    Logger.log(`  ${s}: ${c}`);
  });
  
  Logger.log('\nBy Bucket:');
  Object.entries(byBucket).forEach(([b, c]) => {
    Logger.log(`  ${b}: ${c}`);
  });
  
  return { byStatus, byBucket, total: tasks.length };
}

/**
 * Rollback migration (if needed) - restores original status from Previous_Status column
 * USE WITH CAUTION
 */
function rollbackStatusMigration() {
  Logger.log('=== Status Migration Rollback ===');
  Logger.log('WARNING: This will restore Previous_Status values where available\n');
  
  const tasks = getSheetData(SHEETS.TASKS_DB);
  let rolledBack = 0;
  let skipped = 0;
  
  tasks.forEach(task => {
    if (task.Previous_Status && task.Previous_Status !== task.Status) {
      try {
        updateTask(task.Task_ID, { Status: task.Previous_Status });
        Logger.log(`Rolled back ${task.Task_ID}: ${task.Status} → ${task.Previous_Status}`);
        rolledBack++;
      } catch (e) {
        Logger.log(`Error rolling back ${task.Task_ID}: ${e.toString()}`);
      }
    } else {
      skipped++;
    }
  });
  
  Logger.log(`\nRollback complete: ${rolledBack} rolled back, ${skipped} skipped`);
}

// ============================================
// STEP 3: SET INITIAL CONVERSATION STATES
// ============================================

/**
 * Map old status to initial conversation state
 * Based on what the old status implied about conversation state
 */
function getInitialConversationState(oldStatus, currentStatus) {
  // Map old review statuses to conversation states
  const conversationStateMap = {
    // Old review statuses → conversation states
    'Review_Date': CONVERSATION_STATE.CHANGE_REQUESTED,
    'Review_Date_Boss_Approved': CONVERSATION_STATE.AWAITING_CONFIRMATION,
    'Review_Date_Boss_Rejected': CONVERSATION_STATE.REJECTED,
    'Review_Date_Boss_Proposed': CONVERSATION_STATE.BOSS_PROPOSED,
    'Review_Scope': CONVERSATION_STATE.CHANGE_REQUESTED,
    'Review_Scope_Clarified': CONVERSATION_STATE.RESOLVED,
    'Review_Role': CONVERSATION_STATE.CHANGE_REQUESTED,
    'Review_Stagnation': CONVERSATION_STATE.ACTIVE, // Just slow, no request
    'Review_Update': CONVERSATION_STATE.UPDATE_RECEIVED,
    'Done Pending Review': CONVERSATION_STATE.COMPLETION_PENDING,
    
    // Current lifecycle statuses → default conversation states
    'ai_assist': CONVERSATION_STATE.ACTIVE,
    'not_active': CONVERSATION_STATE.AWAITING_EMPLOYEE,
    'on_time': CONVERSATION_STATE.ACTIVE,
    'slow_progress': CONVERSATION_STATE.ACTIVE,
    'completed': CONVERSATION_STATE.COMPLETION_PENDING,
    'on_hold': CONVERSATION_STATE.ACTIVE,
    'someday': CONVERSATION_STATE.ACTIVE,
    'closed': CONVERSATION_STATE.RESOLVED,
  };
  
  // First check old status
  if (conversationStateMap[oldStatus]) {
    return conversationStateMap[oldStatus];
  }
  
  // Then check current status
  if (conversationStateMap[currentStatus]) {
    return conversationStateMap[currentStatus];
  }
  
  // Default to active
  return CONVERSATION_STATE.ACTIVE;
}

/**
 * Set initial Conversation_State for all tasks that don't have one
 */
function setInitialConversationStates() {
  Logger.log('Setting initial conversation states...');
  
  const tasks = getSheetData(SHEETS.TASKS_DB);
  let setCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  
  tasks.forEach((task, index) => {
    const taskId = task.Task_ID;
    if (!taskId) {
      skippedCount++;
      return;
    }
    
    // Skip if already has a conversation state
    if (task.Conversation_State && task.Conversation_State.trim() !== '') {
      Logger.log(`  - ${taskId}: Already has state "${task.Conversation_State}"`);
      skippedCount++;
      return;
    }
    
    // Determine initial conversation state based on status
    const oldStatus = task.Previous_Status || task.Status;
    const currentStatus = task.Status;
    const conversationState = getInitialConversationState(oldStatus, currentStatus);
    
    try {
      updateTask(taskId, { 
        Conversation_State: conversationState 
      });
      setCount++;
      Logger.log(`  ✓ ${taskId}: Set to "${conversationState}" (was ${oldStatus})`);
    } catch (error) {
      errorCount++;
      Logger.log(`  ✗ ${taskId}: Error - ${error.toString()}`);
    }
  });
  
  Logger.log(`\nConversation States: ${setCount} set, ${skippedCount} skipped, ${errorCount} errors`);
  return { set: setCount, skipped: skippedCount, errors: errorCount };
}

// ============================================
// STEP 4: CREATE INITIAL PARAMETERS SNAPSHOTS
// ============================================

/**
 * Create Initial_Parameters snapshot for tasks that don't have one
 * This snapshot is used to detect changes (e.g., date changed from original)
 */
function createInitialParametersSnapshots() {
  Logger.log('Creating initial parameters snapshots...');
  
  const tasks = getSheetData(SHEETS.TASKS_DB);
  let createdCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  
  tasks.forEach((task, index) => {
    const taskId = task.Task_ID;
    if (!taskId) {
      skippedCount++;
      return;
    }
    
    // Skip if already has initial parameters
    if (task.Initial_Parameters && task.Initial_Parameters.trim() !== '') {
      Logger.log(`  - ${taskId}: Already has Initial_Parameters`);
      skippedCount++;
      return;
    }
    
    // Create snapshot of current parameters
    const snapshot = {
      dueDate: task.Due_Date || null,
      assignee: task.Assignee_Email || null,
      assigneeName: task.Assignee_Name || null,
      taskName: task.Task_Name || '',
      scope: task.Context_Hidden || '',
      projectTag: task.Project_Tag || '',
      createdAt: task.Created_Date || new Date().toISOString()
    };
    
    try {
      updateTask(taskId, { 
        Initial_Parameters: JSON.stringify(snapshot) 
      });
      createdCount++;
      Logger.log(`  ✓ ${taskId}: Created snapshot`);
    } catch (error) {
      errorCount++;
      Logger.log(`  ✗ ${taskId}: Error - ${error.toString()}`);
    }
  });
  
  Logger.log(`\nInitial Parameters: ${createdCount} created, ${skippedCount} skipped, ${errorCount} errors`);
  return { created: createdCount, skipped: skippedCount, errors: errorCount };
}

// ============================================
// PREVIEW / DRY RUN
// ============================================

/**
 * Preview what the full migration would do (DRY RUN - no changes made)
 */
function previewFullMigration() {
  Logger.log('╔════════════════════════════════════════════════════════════╗');
  Logger.log('║       MIGRATION PREVIEW (DRY RUN - NO CHANGES)             ║');
  Logger.log('╚════════════════════════════════════════════════════════════╝\n');
  
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID not found in Script Properties. Please run quickSetup() first.');
  }
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(SHEETS.TASKS_DB);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const tasks = getSheetData(SHEETS.TASKS_DB);
  
  // Check columns
  const newColumns = ['Conversation_State', 'Pending_Changes', 'AI_Summary', 'Initial_Parameters', 'Conversation_History'];
  const missingColumns = newColumns.filter(col => !headers.includes(col));
  
  Logger.log('━━━ COLUMNS ━━━');
  Logger.log(`Existing columns: ${headers.length}`);
  Logger.log(`Missing columns to add: ${missingColumns.length > 0 ? missingColumns.join(', ') : 'None'}`);
  
  // Check statuses
  const statusMigrations = [];
  const conversationStatesToSet = [];
  const paramsToCreate = [];
  
  const newStatuses = Object.values(TASK_STATUS);
  
  tasks.forEach(task => {
    if (!task.Task_ID) return;
    
    // Status migration needed?
    if (!newStatuses.includes(task.Status)) {
      const newStatus = normalizeStatus(task.Status);
      if (newStatus !== task.Status) {
        statusMigrations.push({ id: task.Task_ID, from: task.Status, to: newStatus });
      }
    }
    
    // Conversation state needed?
    if (!task.Conversation_State || task.Conversation_State.trim() === '') {
      const convState = getInitialConversationState(task.Previous_Status || task.Status, task.Status);
      conversationStatesToSet.push({ id: task.Task_ID, state: convState });
    }
    
    // Initial params needed?
    if (!task.Initial_Parameters || task.Initial_Parameters.trim() === '') {
      paramsToCreate.push(task.Task_ID);
    }
  });
  
  Logger.log('\n━━━ STATUS MIGRATIONS ━━━');
  Logger.log(`Tasks needing status migration: ${statusMigrations.length}`);
  statusMigrations.slice(0, 10).forEach(m => {
    Logger.log(`  ${m.id}: "${m.from}" → "${m.to}"`);
  });
  if (statusMigrations.length > 10) {
    Logger.log(`  ... and ${statusMigrations.length - 10} more`);
  }
  
  Logger.log('\n━━━ CONVERSATION STATES ━━━');
  Logger.log(`Tasks needing conversation state: ${conversationStatesToSet.length}`);
  conversationStatesToSet.slice(0, 10).forEach(c => {
    Logger.log(`  ${c.id}: → "${c.state}"`);
  });
  if (conversationStatesToSet.length > 10) {
    Logger.log(`  ... and ${conversationStatesToSet.length - 10} more`);
  }
  
  Logger.log('\n━━━ INITIAL PARAMETERS ━━━');
  Logger.log(`Tasks needing initial parameters snapshot: ${paramsToCreate.length}`);
  
  Logger.log('\n━━━ SUMMARY ━━━');
  Logger.log(`Total tasks: ${tasks.length}`);
  Logger.log(`Columns to add: ${missingColumns.length}`);
  Logger.log(`Statuses to migrate: ${statusMigrations.length}`);
  Logger.log(`Conversation states to set: ${conversationStatesToSet.length}`);
  Logger.log(`Initial params to create: ${paramsToCreate.length}`);
  
  return {
    missingColumns,
    statusMigrations: statusMigrations.length,
    conversationStates: conversationStatesToSet.length,
    initialParams: paramsToCreate.length
  };
}

// ============================================
// VERIFY MIGRATION
// ============================================

/**
 * Verify the migration was successful
 */
function verifyMigration() {
  Logger.log('╔════════════════════════════════════════════════════════════╗');
  Logger.log('║               MIGRATION VERIFICATION                        ║');
  Logger.log('╚════════════════════════════════════════════════════════════╝\n');
  
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID not found in Script Properties. Please run quickSetup() first.');
  }
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(SHEETS.TASKS_DB);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const tasks = getSheetData(SHEETS.TASKS_DB);
  
  // Check columns exist
  const requiredColumns = ['Conversation_State', 'Pending_Changes', 'AI_Summary', 'Initial_Parameters', 'Conversation_History'];
  const missingColumns = requiredColumns.filter(col => !headers.includes(col));
  
  Logger.log('━━━ COLUMNS CHECK ━━━');
  if (missingColumns.length === 0) {
    Logger.log('✓ All required columns exist');
  } else {
    Logger.log(`✗ Missing columns: ${missingColumns.join(', ')}`);
  }
  
  // Check statuses
  const validStatuses = Object.values(TASK_STATUS);
  const invalidStatuses = [];
  const tasksWithoutConvState = [];
  const tasksWithoutInitialParams = [];
  
  tasks.forEach(task => {
    if (!task.Task_ID) return;
    
    if (!validStatuses.includes(task.Status)) {
      invalidStatuses.push({ id: task.Task_ID, status: task.Status });
    }
    
    if (!task.Conversation_State || task.Conversation_State.trim() === '') {
      tasksWithoutConvState.push(task.Task_ID);
    }
    
    if (!task.Initial_Parameters || task.Initial_Parameters.trim() === '') {
      tasksWithoutInitialParams.push(task.Task_ID);
    }
  });
  
  Logger.log('\n━━━ STATUS CHECK ━━━');
  if (invalidStatuses.length === 0) {
    Logger.log('✓ All tasks have valid new statuses');
  } else {
    Logger.log(`✗ ${invalidStatuses.length} tasks have invalid/legacy statuses:`);
    invalidStatuses.slice(0, 5).forEach(t => {
      Logger.log(`    ${t.id}: "${t.status}"`);
    });
  }
  
  Logger.log('\n━━━ CONVERSATION STATE CHECK ━━━');
  if (tasksWithoutConvState.length === 0) {
    Logger.log('✓ All tasks have Conversation_State');
  } else {
    Logger.log(`✗ ${tasksWithoutConvState.length} tasks missing Conversation_State`);
  }
  
  Logger.log('\n━━━ INITIAL PARAMETERS CHECK ━━━');
  if (tasksWithoutInitialParams.length === 0) {
    Logger.log('✓ All tasks have Initial_Parameters');
  } else {
    Logger.log(`✗ ${tasksWithoutInitialParams.length} tasks missing Initial_Parameters`);
  }
  
  // Summary
  const allGood = missingColumns.length === 0 && 
                  invalidStatuses.length === 0 && 
                  tasksWithoutConvState.length === 0 &&
                  tasksWithoutInitialParams.length === 0;
  
  Logger.log('\n━━━ OVERALL STATUS ━━━');
  if (allGood) {
    Logger.log('✓ ✓ ✓ MIGRATION VERIFIED SUCCESSFULLY ✓ ✓ ✓');
  } else {
    Logger.log('✗ ✗ ✗ MIGRATION INCOMPLETE - See issues above ✗ ✗ ✗');
  }
  
  return {
    success: allGood,
    missingColumns: missingColumns.length,
    invalidStatuses: invalidStatuses.length,
    missingConvStates: tasksWithoutConvState.length,
    missingInitialParams: tasksWithoutInitialParams.length
  };
}

