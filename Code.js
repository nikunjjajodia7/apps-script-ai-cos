/**
 * Main Entry Point
 * Chief of Staff AI System - Main Code
 */

/**
 * One-time setup for standalone scripts
 * Run this FIRST if your script is not bound to a spreadsheet
 * 
 * @param {string} spreadsheetId - Your Google Spreadsheet ID (from the URL)
 */
function setupStandaloneScript(spreadsheetId) {
  if (!spreadsheetId) {
    throw new Error('Please provide your Spreadsheet ID. Get it from the spreadsheet URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit');
  }
  
  // Store in Script Properties for future use
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  Logger.log('Spreadsheet ID saved. You can now use the system.');
  Logger.log('Next: Make sure SPREADSHEET_ID is also set in your Config sheet.');
}

/**
 * Quick setup function - EDIT THIS with your Spreadsheet ID and run it
 * Replace 'YOUR_SPREADSHEET_ID_HERE' with your actual Spreadsheet ID
 * 
 * YOUR SPREADSHEET ID: 1lyOwPK5Lvg1fKiGcPt_pL8KHyvHsvuLLi7nYA-z80tc
 */
function quickSetup() {
  // Replace 'YOUR_SPREADSHEET_ID_HERE' with your actual Spreadsheet ID
  const spreadsheetId = '1lyOwPK5Lvg1fKiGcPt_pL8KHyvHsvuLLi7nYA-z80tc'; // <-- PUT YOUR ID HERE
  setupStandaloneScript(spreadsheetId);
}

/**
 * Diagnostic function to check Spreadsheet ID and access
 */
function diagnoseSpreadsheetAccess() {
  Logger.log('=== Spreadsheet Access Diagnosis ===');
  
  // Check Script Properties
  const scriptPropsId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  Logger.log('1. Script Properties ID: ' + (scriptPropsId || 'NOT SET'));
  
  // Check if bound to spreadsheet
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      Logger.log('2. Active Spreadsheet ID: ' + active.getId());
      Logger.log('   Status: Script is BOUND to spreadsheet');
    }
  } catch (e) {
    Logger.log('2. Active Spreadsheet: NOT BOUND (standalone script)');
  }
  
  // Try to open by ID
  if (scriptPropsId) {
    try {
      const testSpreadsheet = SpreadsheetApp.openById(scriptPropsId);
      Logger.log('3. Can open spreadsheet by ID: YES');
      Logger.log('   Spreadsheet name: ' + testSpreadsheet.getName());
      Logger.log('   Spreadsheet URL: ' + testSpreadsheet.getUrl());
      
      // Check if Config sheet exists
      const configSheet = testSpreadsheet.getSheetByName('Config');
      if (configSheet) {
        Logger.log('4. Config sheet exists: YES');
      } else {
        Logger.log('4. Config sheet exists: NO - You need to create it');
      }
    } catch (e) {
      Logger.log('3. Can open spreadsheet by ID: NO');
      Logger.log('   Error: ' + e.toString());
      Logger.log('   Possible causes:');
      Logger.log('   - Spreadsheet ID is incorrect');
      Logger.log('   - You don\'t have access to the spreadsheet');
      Logger.log('   - Spreadsheet doesn\'t exist');
      Logger.log('   - Spreadsheet is in a different Google account');
    }
  } else {
    Logger.log('3. Cannot test - no Spreadsheet ID in Script Properties');
  }
  
  Logger.log('=== End Diagnosis ===');
}

/**
 * Initialize system - run this once after setup
 */
function initialize() {
  try {
    const spreadsheet = getSpreadsheet();
    
    // Check if sheets exist, create if not
    const requiredSheets = [SHEETS.TASKS_DB, SHEETS.STAFF_DB, SHEETS.PROJECTS_DB, 
                             SHEETS.KNOWLEDGE_LAKE, SHEETS.CONFIG, SHEETS.ERROR_LOG];
    
    requiredSheets.forEach(sheetName => {
      if (!spreadsheet.getSheetByName(sheetName)) {
        Logger.log(`Sheet ${sheetName} not found. Please run createAllSheets() first.`);
      }
    });
    
    Logger.log('System initialized successfully');
    
  } catch (error) {
    Logger.log('Initialization error: ' + error);
    try {
      logError(ERROR_TYPE.UNKNOWN_ERROR, 'initialize', error.toString(), null, error.stack);
    } catch (e) {
      // If we can't log, just log to console
      Logger.log('Could not log error: ' + e);
    }
  }
}

/**
 * Test connection to all services
 */
function testConnection() {
  try {
    Logger.log('Testing connections...');
    
    // First, verify spreadsheet access
    let spreadsheet;
    let spreadsheetId;
    
    try {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (spreadsheet) {
        Logger.log(`✓ Spreadsheet access: Using active spreadsheet`);
        spreadsheetId = spreadsheet.getId();
      }
    } catch (e) {
      Logger.log('Not bound to spreadsheet, trying Script Properties...');
    }
    
    // If not bound, try Script Properties
    if (!spreadsheet) {
      spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      Logger.log('Spreadsheet ID from Script Properties: ' + (spreadsheetId || 'NOT FOUND'));
      
      if (!spreadsheetId) {
        throw new Error('Spreadsheet ID not found in Script Properties. Please run quickSetup() first.');
      }
      
      try {
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        Logger.log(`✓ Spreadsheet access: Opened by ID: ${spreadsheetId}`);
      } catch (e) {
        throw new Error(`Could not open spreadsheet with ID: ${spreadsheetId}. Error: ${e.toString()}. Make sure:\n1. The Spreadsheet ID is correct\n2. You have access to the spreadsheet\n3. The spreadsheet exists`);
      }
    }
    
    if (!spreadsheet) {
      throw new Error('Could not access spreadsheet. Spreadsheet ID: ' + (spreadsheetId || 'unknown'));
    }
    
    // Test Sheets access
    try {
      // Check if Tasks_DB sheet exists first
      const tasksSheet = spreadsheet.getSheetByName(SHEETS.TASKS_DB);
      if (tasksSheet) {
        const tasks = getSheetData(SHEETS.TASKS_DB);
        Logger.log(`✓ Sheets access: Found ${tasks.length} tasks`);
      } else {
        Logger.log(`⚠ Sheets access: Tasks_DB sheet not found. Run createAllSheets() to create it.`);
      }
    } catch (e) {
      Logger.log(`⚠ Sheets access: ${e.toString()}`);
    }
    
    // Test Gmail access
    try {
      const threads = GmailApp.getInboxThreads(0, 1);
      Logger.log('✓ Gmail access: OK');
    } catch (e) {
      Logger.log(`⚠ Gmail access: ${e.toString()}`);
    }
    
    // Test Calendar access
    try {
      const events = CalendarApp.getDefaultCalendar().getEvents(new Date(), new Date(Date.now() + 86400000));
      Logger.log(`✓ Calendar access: Found ${events.length} events`);
    } catch (e) {
      Logger.log(`⚠ Calendar access: ${e.toString()}`);
    }
    
    // Test Drive access
    try {
      const files = DriveApp.getRootFolder().getFiles();
      Logger.log('✓ Drive access: OK');
    } catch (e) {
      Logger.log(`⚠ Drive access: ${e.toString()}`);
    }
    
    // Test Config
    try {
      const bossEmail = CONFIG.BOSS_EMAIL();
      Logger.log(`✓ Config access: Boss email = ${bossEmail}`);
    } catch (e) {
      Logger.log(`⚠ Config access: ${e.toString()}`);
    }
    
    Logger.log('Connection test completed!');
    return true;
    
  } catch (error) {
    Logger.log('Connection test failed: ' + error);
    Logger.log('Error details: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    try {
      logError(ERROR_TYPE.UNKNOWN_ERROR, 'testConnection', error.toString(), null, error.stack);
    } catch (e) {
      // If we can't log, that's okay
    }
    return false;
  }
}

/**
 * Setup triggers (run this once after deployment)
 * This will create all time-driven triggers automatically
 */
function setupTriggers() {
  try {
    Logger.log('=== Setting up triggers ===');
    
    // Delete existing triggers first
    const existingTriggers = ScriptApp.getProjectTriggers();
    Logger.log(`Found ${existingTriggers.length} existing trigger(s)`);
    
    existingTriggers.forEach(trigger => {
      const functionName = trigger.getHandlerFunction();
      Logger.log(`Deleting trigger: ${functionName}`);
      ScriptApp.deleteTrigger(trigger);
    });
    
    // Note: Drive triggers must be set up manually in the Apps Script editor
    // Go to Triggers > Add Trigger > onVoiceFileAdded (From Drive, On file change)
    // Go to Triggers > Add Trigger > onMeetingFileAdded (From Drive, On file change)
    
    // Time-driven triggers
    Logger.log('Creating trigger: handleSilenceEscalation (every hour)');
    ScriptApp.newTrigger('handleSilenceEscalation')
      .timeBased()
      .everyHours(1)
      .create();
    
    Logger.log('Creating trigger: updateReliabilityScores (daily at 2 AM)');
    ScriptApp.newTrigger('updateReliabilityScores')
      .timeBased()
      .everyDays(1)
      .atHour(2) // 2 AM
      .create();
    
    Logger.log('Creating trigger: checkForReplies (every 15 minutes)');
    ScriptApp.newTrigger('checkForReplies')
      .timeBased()
      .everyMinutes(15)
      .create();
    
    Logger.log('Creating trigger: checkVoiceInbox (every 5 minutes)');
    ScriptApp.newTrigger('checkVoiceInbox')
      .timeBased()
      .everyMinutes(5)
      .create();
    
    Logger.log('Creating trigger: checkMeetingLake (every 15 minutes)');
    ScriptApp.newTrigger('checkMeetingLake')
      .timeBased()
      .everyMinutes(15)
      .create();
    
    // Verify triggers were created
    const newTriggers = ScriptApp.getProjectTriggers();
    Logger.log(`\n=== Triggers created successfully ===`);
    Logger.log(`Total triggers: ${newTriggers.length}`);
    newTriggers.forEach(trigger => {
      Logger.log(`  - ${trigger.getHandlerFunction()} (${trigger.getEventType()})`);
    });
    
    Logger.log('\nNote: Drive triggers must be set up manually in the editor');
    Logger.log('Go to Triggers > Add Trigger > onVoiceFileAdded (From Drive, On file change)');
    Logger.log('Go to Triggers > Add Trigger > onMeetingFileAdded (From Drive, On file change)');
    
  } catch (error) {
    Logger.log('ERROR: Trigger setup failed: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'setupTriggers', error.toString(), null, error.stack);
  }
}

/**
 * List all current triggers (for debugging)
 */
function listTriggers() {
  try {
    Logger.log('=== Current Triggers ===');
    const triggers = ScriptApp.getProjectTriggers();
    
    if (triggers.length === 0) {
      Logger.log('No triggers found');
      return;
    }
    
    triggers.forEach((trigger, index) => {
      Logger.log(`\n${index + 1}. Function: ${trigger.getHandlerFunction()}`);
      Logger.log(`   Event Type: ${trigger.getEventType()}`);
      Logger.log(`   Unique ID: ${trigger.getUniqueId()}`);
      
      if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
        Logger.log(`   Time-based trigger`);
      }
    });
    
    Logger.log(`\nTotal: ${triggers.length} trigger(s)`);
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
  }
}

/**
 * Update reliability scores for all staff
 */
function updateReliabilityScores() {
  try {
    const staff = getSheetData(SHEETS.STAFF_DB);
    
    staff.forEach(member => {
      if (!member.Email) return;
      
      const score = calculateReliabilityScore(member.Email);
      updateStaff(member.Email, { Reliability_Score: score });
    });
    
    Logger.log('Reliability scores updated');
    
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'updateReliabilityScores', error.toString(), null, error.stack);
  }
}

/**
 * Calculate reliability score for a staff member
 */
function calculateReliabilityScore(staffEmail) {
  try {
    const tasks = getTasksByAssignee(staffEmail);
    if (tasks.length === 0) {
      return 100; // No tasks = perfect score (or could be 0, depending on preference)
    }
    
    let completedOnTime = 0;
    let totalCompleted = 0;
    let extensionsRequested = 0;
    let stagnations = 0;
    
    tasks.forEach(task => {
      if (task.Status === TASK_STATUS.DONE) {
        totalCompleted++;
        const dueDate = task.Due_Date ? new Date(task.Due_Date) : null;
        const completedDate = task.Last_Updated ? new Date(task.Last_Updated) : null;
        
        if (dueDate && completedDate && completedDate <= dueDate) {
          completedOnTime++;
        }
      }
      
      // Count extensions requested
      if (task.Interaction_Log && task.Interaction_Log.includes('requested date change')) {
        extensionsRequested++;
      }
      
      // Count stagnations
      if (task.Status === TASK_STATUS.REVIEW_STAGNATION) {
        stagnations++;
      }
    });
    
    // Base score: percentage completed on time
    let score = totalCompleted > 0 ? (completedOnTime / totalCompleted) * 100 : 100;
    
    // Adjustments
    score -= extensionsRequested * 5; // -5 points per extension
    score -= stagnations * 10; // -10 points per stagnation
    
    // Clamp to 0-100
    score = Math.max(0, Math.min(100, score));
    
    return Math.round(score);
    
  } catch (error) {
    logError(ERROR_TYPE.DATA_ERROR, 'calculateReliabilityScore', error.toString(), null, error.stack);
    return 50; // Default score on error
  }
}

/**
 * Update active task counts for all staff
 */
function updateAllTaskCounts() {
  try {
    updateAllActiveTaskCounts();
    Logger.log('Active task counts updated');
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'updateAllTaskCounts', error.toString(), null, error.stack);
  }
}

/**
 * Check Voice_Inbox folder for new files (called by time-driven trigger)
 */
function checkVoiceInbox() {
  try {
    Logger.log('=== Checking Voice_Inbox ===');
    const voiceInboxFolderId = CONFIG.VOICE_INBOX_FOLDER_ID();
    Logger.log('VOICE_INBOX_FOLDER_ID: ' + (voiceInboxFolderId || 'NOT CONFIGURED'));
    
    if (!voiceInboxFolderId) {
      Logger.log('ERROR: VOICE_INBOX_FOLDER_ID not configured in Config sheet');
      Logger.log('Please add VOICE_INBOX_FOLDER_ID to your Config sheet');
      return;
    }
    
    let folder;
    try {
      folder = DriveApp.getFolderById(voiceInboxFolderId);
      Logger.log('Folder found: ' + folder.getName());
    } catch (e) {
      Logger.log('ERROR: Could not access folder with ID: ' + voiceInboxFolderId);
      Logger.log('Error: ' + e.toString());
      Logger.log('Make sure:');
      Logger.log('1. The folder ID is correct');
      Logger.log('2. You have access to the folder');
      Logger.log('3. The folder exists');
      return;
    }
    
    const files = folder.getFiles();
    const now = new Date();
    // Changed to 1 hour for testing - you can change this back to 5 minutes later
    const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
    Logger.log('Checking for files modified after: ' + oneHourAgo);
    
    let fileCount = 0;
    let processedCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      const lastModified = file.getLastUpdated();
      const fileName = file.getName();
      const mimeType = file.getMimeType();
      
      Logger.log(`File ${fileCount}: ${fileName} (Modified: ${lastModified}, Type: ${mimeType})`);
      
      // Skip files that have already been processed
      if (fileName.startsWith('[PROCESSED]') || fileName.startsWith('[UNCLEAR]')) {
        Logger.log(`  -> Skipping (already processed): ${fileName}`);
        continue;
      }
      
      // Process files modified in the last hour (changed from 5 minutes for testing)
      if (lastModified > oneHourAgo) {
        Logger.log(`  -> Processing: ${fileName}`);
        try {
          processVoiceNote(file.getId());
          processedCount++;
          Logger.log(`  -> Successfully processed: ${fileName}`);
        } catch (e) {
          Logger.log(`  -> ERROR processing ${fileName}: ${e.toString()}`);
        }
      } else {
        Logger.log(`  -> Skipping (too old): ${fileName}`);
      }
    }
    
    Logger.log(`=== Summary: Found ${fileCount} files, processed ${processedCount} ===`);
    
    if (fileCount === 0) {
      Logger.log('No files found in Voice_Inbox folder. Make sure:');
      Logger.log('1. Files are uploaded to the correct folder');
      Logger.log('2. Folder ID is correct in Config sheet');
    }
    
  } catch (error) {
    Logger.log('ERROR in checkVoiceInbox: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    try {
      logError(ERROR_TYPE.UNKNOWN_ERROR, 'checkVoiceInbox', error.toString(), null, error.stack);
    } catch (e) {
      // If we can't log, that's okay
    }
  }
}

/**
 * Check Meeting_Lake folder for new files (called by time-driven trigger)
 */
function checkMeetingLake() {
  try {
    Logger.log('=== Checking Meeting_Lake folder ===');
    
    const meetingLakeFolderId = CONFIG.MEETING_LAKE_FOLDER_ID();
    if (!meetingLakeFolderId) {
      Logger.log('ERROR: MEETING_LAKE_FOLDER_ID not configured');
      Logger.log('Please set MEETING_LAKE_FOLDER_ID in your Config sheet');
      return;
    }
    
    const folder = DriveApp.getFolderById(meetingLakeFolderId);
    const files = folder.getFiles();
    const now = new Date();
    // Increased to 1 hour for testing (change back to 5 minutes for production)
    const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
    
    let fileCount = 0;
    let processedCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      const lastModified = file.getLastUpdated();
      const fileName = file.getName();
      
      Logger.log(`File ${fileCount}: ${fileName} (Modified: ${lastModified})`);
      
      // Process files modified in the last hour (for testing)
      if (lastModified > oneHourAgo) {
        Logger.log(`  -> Processing: ${fileName}`);
        try {
          processMeetingAudio(file.getId());
          processedCount++;
        } catch (error) {
          Logger.log(`  -> ERROR processing ${fileName}: ${error.toString()}`);
        }
      } else {
        Logger.log(`  -> Skipping (too old): ${fileName}`);
      }
    }
    
    Logger.log(`=== Summary: Found ${fileCount} file(s), processed ${processedCount} ===`);
    
  } catch (error) {
    Logger.log(`ERROR in checkMeetingLake: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'checkMeetingLake', error.toString(), null, error.stack);
  }
}

/**
 * Create all sheets with proper schema
 * Run this function once to set up the database structure
 */
function createAllSheets() {
  const spreadsheet = getSpreadsheet();
  
  // Create or get Tasks_DB sheet
  let tasksSheet = spreadsheet.getSheetByName('Tasks_DB');
  if (!tasksSheet) {
    tasksSheet = spreadsheet.insertSheet('Tasks_DB');
    tasksSheet.getRange(1, 1, 1, 16).setValues([[
      'Task_ID', 'Task_Name', 'Status', 'Assignee_Email', 'Due_Date', 
      'Proposed_Date', 'Project_Tag', 'Meeting_Action', 'AI_Confidence', 
      'Tone_Detected', 'Context_Hidden', 'Interaction_Log', 'Boss_Reply_Draft',
      'Created_Date', 'Last_Updated', 'Priority'
    ]]);
    tasksSheet.getRange(1, 1, 1, 16).setFontWeight('bold');
    tasksSheet.setFrozenRows(1);
    Logger.log('Created Tasks_DB sheet');
  }
  
  // Create or get Staff_DB sheet
  let staffSheet = spreadsheet.getSheetByName('Staff_DB');
  if (!staffSheet) {
    staffSheet = spreadsheet.insertSheet('Staff_DB');
    staffSheet.getRange(1, 1, 1, 8).setValues([[
      'Name', 'Email', 'Role', 'Reliability_Score', 'Active_Task_Count',
      'Department', 'Manager_Email', 'Last_Updated'
    ]]);
    staffSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    staffSheet.setFrozenRows(1);
    Logger.log('Created Staff_DB sheet');
  }
  
  // Create or get Projects_DB sheet
  let projectsSheet = spreadsheet.getSheetByName('Projects_DB');
  if (!projectsSheet) {
    projectsSheet = spreadsheet.insertSheet('Projects_DB');
    projectsSheet.getRange(1, 1, 1, 8).setValues([[
      'Project_Tag', 'Project_Name', 'Team_Lead_Email', 'Status', 'Priority',
      'Start_Date', 'End_Date', 'Description'
    ]]);
    projectsSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    projectsSheet.setFrozenRows(1);
    Logger.log('Created Projects_DB sheet');
  }
  
  // Create or get Knowledge_Lake sheet
  let knowledgeSheet = spreadsheet.getSheetByName('Knowledge_Lake');
  if (!knowledgeSheet) {
    knowledgeSheet = spreadsheet.insertSheet('Knowledge_Lake');
    knowledgeSheet.getRange(1, 1, 1, 8).setValues([[
      'Info_ID', 'Link', 'Source_Type', 'Summary', 'Created_Date',
      'Meeting_Date', 'Related_Tasks', 'Tags'
    ]]);
    knowledgeSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    knowledgeSheet.setFrozenRows(1);
    Logger.log('Created Knowledge_Lake sheet');
  }
  
  // Create or get Config sheet
  let configSheet = spreadsheet.getSheetByName('Config');
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet('Config');
    configSheet.getRange(1, 1, 1, 4).setValues([[
      'Key', 'Value', 'Description', 'Category'
    ]]);
    configSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    configSheet.setFrozenRows(1);
    
    // Add default configuration
    const defaultConfig = [
      ['BOSS_EMAIL', Session.getActiveUser().getEmail(), 'Boss email address', 'System'],
      ['ESCALATION_FOLLOWUP_HOURS', '24', 'Hours before first follow-up', 'Timing'],
      ['ESCALATION_BOSS_ALERT_HOURS', '48', 'Hours before alerting Boss', 'Timing'],
      ['DEFAULT_MEETING_DURATION_MINUTES', '30', 'Default meeting duration', 'Scheduling'],
      ['FOCUS_TIME_DURATION_MINUTES', '60', 'Focus time block duration', 'Scheduling'],
      ['WORKING_HOURS_START', '09:00', 'Start of working hours', 'Scheduling'],
      ['WORKING_HOURS_END', '17:00', 'End of working hours', 'Scheduling'],
      ['AI_CONFIDENCE_THRESHOLD', '0.6', 'Minimum confidence for auto-processing', 'AI'],
      ['RELIABILITY_UPDATE_INTERVAL_HOURS', '24', 'Reliability score update interval', 'Scoring'],
      ['EMAIL_SIGNATURE', '[Boss\'s Chief of Staff AI]', 'Email signature', 'Email'],
      ['WEEKLY_MEETING_TITLE', 'Weekly Ops', 'Recurring weekly meeting title', 'Scheduling']
    ];
    configSheet.getRange(2, 1, defaultConfig.length, 4).setValues(defaultConfig);
    Logger.log('Created Config sheet with default values');
  }
  
  // Create or get Error_Log sheet
  let errorSheet = spreadsheet.getSheetByName('Error_Log');
  if (!errorSheet) {
    errorSheet = spreadsheet.insertSheet('Error_Log');
    errorSheet.getRange(1, 1, 1, 8).setValues([[
      'Timestamp', 'Error_Type', 'Function_Name', 'Error_Message',
      'Task_ID', 'Stack_Trace', 'Resolved', 'Resolution_Notes'
    ]]);
    errorSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    errorSheet.setFrozenRows(1);
    Logger.log('Created Error_Log sheet');
  }
  
  Logger.log('All sheets created/verified successfully!');
}

/**
 * Add sample data for testing
 */
function addSampleData() {
  const spreadsheet = getSpreadsheet();
  
  // Add sample staff
  const staffSheet = spreadsheet.getSheetByName('Staff_DB');
  if (staffSheet && staffSheet.getLastRow() === 1) {
    staffSheet.appendRow([
      Session.getActiveUser().getName() || 'Boss', 
      Session.getActiveUser().getEmail(), 
      'CEO', 
      '', // Reliability_Score
      '', // Active_Task_Count
      'Executive', 
      '', // Manager_Email
      new Date() // Last_Updated
    ]);
    Logger.log('Added sample staff member (yourself)');
  }
  
  // Add sample project
  const projectsSheet = spreadsheet.getSheetByName('Projects_DB');
  if (projectsSheet && projectsSheet.getLastRow() === 1) {
    projectsSheet.appendRow([
      'General', 
      'General Tasks', 
      Session.getActiveUser().getEmail(), 
      'Active', 
      'Medium',
      '', // Start_Date
      '', // End_Date
      'General task management' // Description
    ]);
    Logger.log('Added sample project');
  }
  
  Logger.log('Sample data added!');
}

/**
 * Diagnostic function to test Voice_Inbox folder access
 * Run this to check if files can be found
 */
function testVoiceInbox() {
  try {
    Logger.log('=== Testing Voice_Inbox Access ===');
    
    const folderId = CONFIG.VOICE_INBOX_FOLDER_ID();
    Logger.log('1. VOICE_INBOX_FOLDER_ID from Config: ' + (folderId || 'NOT SET'));
    
    if (!folderId) {
      Logger.log('ERROR: VOICE_INBOX_FOLDER_ID not set in Config sheet');
      Logger.log('Please add it to your Config sheet');
      return;
    }
    
    Logger.log('2. Attempting to access folder...');
    const folder = DriveApp.getFolderById(folderId);
    Logger.log('   ✓ Folder found: ' + folder.getName());
    Logger.log('   Folder URL: ' + folder.getUrl());
    
    Logger.log('3. Listing all files in folder...');
    const files = folder.getFiles();
    let count = 0;
    const now = new Date();
    const fiveMinutesAgo = new Date(now.getTime() - 5 * 60 * 1000);
    
    while (files.hasNext()) {
      const file = files.next();
      count++;
      const lastModified = file.getLastUpdated();
      const isRecent = lastModified > fiveMinutesAgo;
      Logger.log(`   File ${count}: ${file.getName()}`);
      Logger.log(`      - Modified: ${lastModified}`);
      Logger.log(`      - Type: ${file.getMimeType()}`);
      Logger.log(`      - Recent (< 5 min): ${isRecent ? 'YES' : 'NO'}`);
      if (isRecent) {
        Logger.log(`      -> Would be processed by checkVoiceInbox()`);
      }
    }
    
    Logger.log(`4. Summary: Found ${count} file(s) in folder`);
    
    if (count === 0) {
      Logger.log('   WARNING: No files found in Voice_Inbox folder');
      Logger.log('   Make sure you uploaded a file to the correct folder');
    }
    
    Logger.log('=== End Test ===');
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Manually process a specific file by name
 * Usage: processFileByName('Test recording .m4a')
 * Or modify the function to hardcode the filename for quick testing
 */
function processFileByName(fileName) {
  // If no filename provided, use the test file
  if (!fileName) {
    fileName = 'Test recording .m4a';
  }
  try {
    Logger.log(`=== Processing file: ${fileName} ===`);
    
    const folderId = CONFIG.VOICE_INBOX_FOLDER_ID();
    if (!folderId) {
      Logger.log('ERROR: VOICE_INBOX_FOLDER_ID not set in Config sheet');
      return;
    }
    
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);
    
    if (!files.hasNext()) {
      Logger.log(`ERROR: File "${fileName}" not found in Voice_Inbox folder`);
      Logger.log('Make sure the file name matches exactly (including extension)');
      return;
    }
    
    const file = files.next();
    Logger.log(`Found file: ${file.getName()}`);
    Logger.log(`File ID: ${file.getId()}`);
    Logger.log(`Modified: ${file.getLastUpdated()}`);
    
    Logger.log('Processing file...');
    processVoiceNote(file.getId());
    Logger.log('✓ File processed successfully!');
    Logger.log('Check your Tasks sheet for the new task.');
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Quick test function - processes the test file
 * Just run this function directly
 */
function processTestFile() {
  processFileByName('Test recording .m4a');
}

/**
 * Diagnostic function to inspect Interaction_Log entries
 * Run this first to see what's in the logs
 */
function inspectInteractionLogs() {
  try {
    Logger.log('=== Inspecting Interaction_Log entries ===');
    
    const tasks = getSheetData(SHEETS.TASKS_DB);
    Logger.log(`Found ${tasks.length} tasks`);
    
    tasks.forEach((task, index) => {
      if (!task.Interaction_Log) {
        Logger.log(`Task ${index + 1} (${task.Task_ID}): No Interaction_Log`);
        return;
      }
      
      const log = task.Interaction_Log;
      const logLength = log.length;
      const lineCount = log.split('\n').length;
      
      Logger.log(`\nTask ${index + 1} (${task.Task_ID}):`);
      Logger.log(`  Log length: ${logLength} characters`);
      Logger.log(`  Line count: ${lineCount}`);
      Logger.log(`  Contains "Interaction_Log": ${log.includes('"Interaction_Log"')}`);
      Logger.log(`  Contains "Last_Updated": ${log.includes('"Last_Updated"')}`);
      Logger.log(`  Contains nested JSON: ${log.includes('"Interaction_Log"') && log.includes('"Last_Updated"')}`);
      Logger.log(`  First 200 chars: ${log.substring(0, 200)}...`);
      
      // Check if it looks corrupted (very long or has nested JSON)
      if (logLength > 500 || (log.includes('"Interaction_Log"') && log.includes('"Last_Updated"'))) {
        Logger.log(`  -> This log looks CORRUPTED`);
      }
    });
    
    Logger.log('\n=== Inspection complete ===');
    
  } catch (error) {
    Logger.log('ERROR in inspectInteractionLogs: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Clean up corrupted Interaction_Log entries
 * This removes the recursive JSON nesting that occurred due to the bug
 */
function cleanupInteractionLogs() {
  try {
    Logger.log('=== Cleaning up Interaction_Log entries ===');
    
    const tasks = getSheetData(SHEETS.TASKS_DB);
    let cleanedCount = 0;
    
    tasks.forEach(task => {
      if (!task.Interaction_Log) return;
      
      const log = task.Interaction_Log;
      const logLength = log.length;
      
      // Check if log is corrupted (either has nested JSON or is suspiciously long)
      const hasNestedJSON = log.includes('"Interaction_Log"') && log.includes('"Last_Updated"');
      const isSuspiciouslyLong = logLength > 500; // Normal logs should be < 500 chars
      
      if (hasNestedJSON || isSuspiciouslyLong) {
        Logger.log(`Found corrupted log for task: ${task.Task_ID} (${logLength} chars)`);
        
        // Extract just the first meaningful log entry
        const lines = log.split('\n');
        const firstRealEntry = lines.find(line => 
          line.includes(' - ') && 
          !line.includes('"Interaction_Log"') &&
          !line.includes('"Last_Updated"') &&
          !line.includes('Task updated:')
        );
        
        if (firstRealEntry) {
          const cleanLog = firstRealEntry.trim();
          updateTask(task.Task_ID, { Interaction_Log: cleanLog });
          cleanedCount++;
          Logger.log(`  -> Cleaned: "${cleanLog.substring(0, 50)}..."`);
        } else {
          // If no clean entry found, create a simple summary
          const taskName = task.Task_Name || 'Unknown task';
          const cleanLog = `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')} - Task created: ${taskName} (log cleaned)`;
          updateTask(task.Task_ID, { Interaction_Log: cleanLog });
          cleanedCount++;
          Logger.log(`  -> Replaced with summary: "${cleanLog}"`);
        }
      }
    });
    
    Logger.log(`=== Cleanup complete: ${cleanedCount} logs cleaned ===`);
    
    if (cleanedCount === 0) {
      Logger.log('No corrupted logs found. All logs appear clean!');
    }
    
  } catch (error) {
    Logger.log('ERROR in cleanupInteractionLogs: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Clean up verbose JSON logs (simplifies "Task updated" entries)
 * This removes the JSON dumps from update logs, keeping only essential info
 */
function cleanupVerboseLogs() {
  try {
    Logger.log('=== Cleaning up verbose Interaction_Log entries ===');
    
    const tasks = getSheetData(SHEETS.TASKS_DB);
    let cleanedCount = 0;
    
    tasks.forEach(task => {
      if (!task.Interaction_Log) return;
      
      const log = task.Interaction_Log;
      const lines = log.split('\n');
      const cleanedLines = [];
      
      lines.forEach(line => {
        // If line contains "Task updated:" with JSON, simplify it
        if (line.includes('Task updated:') && line.includes('{')) {
          // Extract just the timestamp and action, remove JSON
          const timestamp = line.split(' - ')[0];
          cleanedLines.push(`${timestamp} - Task updated`);
        } else {
          // Keep other lines as-is
          cleanedLines.push(line);
        }
      });
      
      const newLog = cleanedLines.join('\n');
      
      // Only update if we actually changed something
      if (newLog !== log) {
        updateTask(task.Task_ID, { Interaction_Log: newLog });
        cleanedCount++;
        Logger.log(`Cleaned verbose log for task: ${task.Task_ID}`);
        Logger.log(`  Before: ${log.length} chars, After: ${newLog.length} chars`);
      }
    });
    
    Logger.log(`=== Cleanup complete: ${cleanedCount} logs cleaned ===`);
    
  } catch (error) {
    Logger.log('ERROR in cleanupVerboseLogs: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Clean a specific task's log by Task_ID
 * Usage: cleanTaskLog('TASK-20251221022155')
 */
function cleanTaskLog(taskId) {
  try {
    Logger.log(`=== Cleaning log for task: ${taskId} ===`);
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    if (!task.Interaction_Log) {
      Logger.log(`Task ${taskId} has no Interaction_Log`);
      return;
    }
    
    const log = task.Interaction_Log;
    const lines = log.split('\n');
    const cleanedLines = [];
    
    lines.forEach(line => {
      // Keep only essential log entries
      if (line.includes('Task created:')) {
        cleanedLines.push(line);
      } else if (line.includes('Task updated:')) {
        // Simplify update entries
        const timestamp = line.split(' - ')[0];
        cleanedLines.push(`${timestamp} - Task updated`);
      } else if (line.includes('Assignment') || line.includes('Email sent')) {
        cleanedLines.push(line);
      }
      // Skip verbose JSON entries
    });
    
    const newLog = cleanedLines.join('\n');
    updateTask(taskId, { Interaction_Log: newLog });
    
    Logger.log(`✓ Cleaned log for task: ${taskId}`);
    Logger.log(`  Before: ${log.length} chars, ${log.split('\n').length} lines`);
    Logger.log(`  After: ${newLog.length} chars, ${newLog.split('\n').length} lines`);
    
  } catch (error) {
    Logger.log('ERROR in cleanTaskLog: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Clean all logs - removes verbose JSON from all tasks
 */
function cleanAllLogs() {
  cleanupVerboseLogs();
}

/**
 * List all tasks with assignees (helper function)
 */
function listTasksWithAssignees() {
  try {
    Logger.log('=== Tasks with Assignees ===');
    const tasks = getSheetData(SHEETS.TASKS_DB);
    const tasksWithAssignees = tasks.filter(t => t.Assignee_Email && t.Assignee_Email.trim() !== '');
    
    if (tasksWithAssignees.length === 0) {
      Logger.log('No tasks with assignees found');
      Logger.log('Please create a task and set Assignee_Email in the Tasks sheet');
      return;
    }
    
    Logger.log(`Found ${tasksWithAssignees.length} task(s) with assignees:`);
    tasksWithAssignees.forEach((task, index) => {
      Logger.log(`${index + 1}. ${task.Task_ID} - ${task.Task_Name}`);
      Logger.log(`   Assignee: ${task.Assignee_Email}`);
      Logger.log(`   Status: ${task.Status}`);
    });
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
  }
}

/**
 * Test function to send task assignment email
 * Usage: testSendAssignmentEmail() - will use the most recent task with an assignee
 * Or modify the function to hardcode a task ID
 */
function testSendAssignmentEmail() {
  try {
    Logger.log('=== Testing Task Assignment Email ===');
    
    // Find a task with an assignee
    Logger.log('Finding task with assignee...');
    const tasks = getSheetData(SHEETS.TASKS_DB);
    
    if (tasks.length === 0) {
      Logger.log('ERROR: No tasks found in database');
      Logger.log('Please create a task first');
      return;
    }
    
    Logger.log(`Found ${tasks.length} total task(s)`);
    
    // Find a task with an assignee
    const taskWithAssignee = tasks.find(t => t.Assignee_Email && t.Assignee_Email.trim() !== '');
    
    if (!taskWithAssignee) {
      Logger.log('ERROR: No tasks with assignees found');
      Logger.log('Available tasks:');
      tasks.forEach((t, i) => {
        Logger.log(`  ${i + 1}. ${t.Task_ID} - ${t.Task_Name} (Assignee: ${t.Assignee_Email || 'NONE'})`);
      });
      Logger.log('\nPlease set Assignee_Email for a task in the Tasks sheet');
      return;
    }
    
    const taskId = taskWithAssignee.Task_ID;
    Logger.log(`✓ Found task: ${taskId} - ${taskWithAssignee.Task_Name}`);
    Logger.log(`  Assignee: ${taskWithAssignee.Assignee_Email}`);
    Logger.log(`  Status: ${taskWithAssignee.Status}`);
    
    // Verify task exists
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found when retrieving`);
      return;
    }
    
    if (!task.Assignee_Email || task.Assignee_Email.trim() === '') {
      Logger.log('ERROR: Task has no assignee email');
      Logger.log('Please set Assignee_Email in the Tasks sheet');
      return;
    }
    
    // Update status to Assigned if needed
    if (task.Status !== TASK_STATUS.ASSIGNED) {
      Logger.log(`Updating status from "${task.Status}" to "Assigned"...`);
      updateTask(taskId, { Status: TASK_STATUS.ASSIGNED });
      // Re-fetch task to get updated status
      const updatedTask = getTask(taskId);
      Logger.log(`Status updated to: ${updatedTask.Status}`);
    }
    
    // Send email
    Logger.log(`\nSending assignment email to ${task.Assignee_Email}...`);
    Logger.log(`Task ID being passed: ${taskId}`);
    
    sendTaskAssignmentEmail(taskId);
    
    Logger.log('✓ Email sent successfully!');
    Logger.log(`Check inbox for: ${task.Assignee_Email}`);
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Send email for a specific task ID
 * Modify the task ID below to test a specific task
 */
function testEmailAssignmentSpecific() {
  // CHANGE THIS to your task ID
  const taskId = 'TASK-20251221023047'; // Replace with your actual task ID
  testSendAssignmentEmailForTask(taskId);
}

/**
 * Send email for a specific task (internal function)
 */
function testSendAssignmentEmailForTask(taskId) {
  try {
    if (!taskId) {
      Logger.log('ERROR: No task ID provided');
      Logger.log('Run listTasksWithAssignees() to see available tasks');
      return;
    }
    
    Logger.log(`=== Sending email for task: ${taskId} ===`);
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      Logger.log('Run listTasksWithAssignees() to see available tasks');
      return;
    }
    
    Logger.log(`Task: ${task.Task_Name}`);
    Logger.log(`Assignee: ${task.Assignee_Email || 'NOT SET'}`);
    
    if (!task.Assignee_Email || task.Assignee_Email.trim() === '') {
      Logger.log('ERROR: Task has no assignee email');
      return;
    }
    
    // Update status if needed
    if (task.Status !== TASK_STATUS.ASSIGNED) {
      updateTask(taskId, { Status: TASK_STATUS.ASSIGNED });
    }
    
    Logger.log('Sending email...');
    sendTaskAssignmentEmail(taskId);
    
    Logger.log('✓ Email sent!');
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
  }
}

/**
 * Set MEETING_LAKE_FOLDER_ID in Config sheet
 * Run this function to configure the Meeting Lake folder ID
 */
function setMeetingLakeFolderId() {
  const folderId = '180zM9CKnspyUUD7xOt1MJDruUWM3WK4T';
  setConfigValue('MEETING_LAKE_FOLDER_ID', folderId, 'Google Drive folder for meeting recordings', 'System');
  Logger.log('MEETING_LAKE_FOLDER_ID set to: ' + folderId);
  Logger.log('You can verify this in your Config sheet.');
  
  // Verify the folder exists
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log('✓ Folder verified: ' + folder.getName());
    Logger.log('  Folder URL: ' + folder.getUrl());
  } catch (e) {
    Logger.log('⚠ Warning: Could not verify folder access. Error: ' + e.toString());
    Logger.log('  Make sure the folder ID is correct and you have access to it.');
  }
}

