/**
 * Status Migration Script
 * Migrates existing tasks from old status values to new unified status system
 * 
 * Run this function ONCE after deploying the new status system to update
 * all existing tasks in your Tasks_DB sheet.
 */

/**
 * Main migration function - run this from Apps Script editor
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

