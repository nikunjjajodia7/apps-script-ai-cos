/**
 * Sheets Helper Utilities
 * Functions for reading and writing to Google Sheets
 */

/**
 * Get the spreadsheet (works for both bound and standalone scripts)
 */
function getSpreadsheet() {
  // Try to get active spreadsheet first (if script is bound to a sheet)
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (spreadsheet) {
      return spreadsheet;
    }
  } catch (e) {
    // Not bound, continue
  }
  
  // If null, it's a standalone script - use spreadsheet ID from Script Properties first
  // (avoid circular dependency with Config)
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      return spreadsheet;
    } catch (e) {
      throw new Error('Could not open spreadsheet with ID from Script Properties: ' + spreadsheetId + '. Error: ' + e.toString());
    }
  }
  
  // Fallback: try to get from Config (but this might cause circular dependency)
  try {
    const configSpreadsheetId = CONFIG.SPREADSHEET_ID();
    if (configSpreadsheetId) {
      spreadsheet = SpreadsheetApp.openById(configSpreadsheetId);
      return spreadsheet;
    }
  } catch (e) {
    // Config not available yet, that's okay
  }
  
  throw new Error('SPREADSHEET_ID not found. Please run quickSetup() or setupStandaloneScript() first with your Spreadsheet ID.');
}

/**
 * Get a sheet by name
 */
function getSheet(sheetName) {
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in spreadsheet. Make sure you've run createAllSheets() first.`);
  }
  return sheet;
}

/**
 * Get all rows from a sheet as objects
 */
function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];
  
  const headers = data[0];
  const rows = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((header, index) => {
      row[header] = data[i][index];
    });
    rows.push(row);
  }
  
  return rows;
}

/**
 * Find a row by a specific column value
 */
function findRowByValue(sheetName, columnName, value) {
  const rows = getSheetData(sheetName);
  return rows.find(row => row[columnName] === value);
}

/**
 * Find all rows matching a condition
 */
function findRowsByCondition(sheetName, conditionFn) {
  const rows = getSheetData(sheetName);
  return rows.filter(conditionFn);
}

/**
 * Add a new row to a sheet
 */
function addRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => rowData[header] || '');
  sheet.appendRow(row);
  return sheet.getLastRow();
}

/**
 * Update a row by matching a column value
 */
function updateRowByValue(sheetName, matchColumn, matchValue, updates) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const matchColumnIndex = headers.indexOf(matchColumn);
  
  if (matchColumnIndex === -1) {
    throw new Error(`Column "${matchColumn}" not found`);
  }
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][matchColumnIndex] === matchValue) {
      // Update the row
      Object.keys(updates).forEach(key => {
        const columnIndex = headers.indexOf(key);
        if (columnIndex !== -1) {
          sheet.getRange(i + 1, columnIndex + 1).setValue(updates[key]);
        }
      });
      return i + 1; // Return row number (1-indexed)
    }
  }
  
  return null; // Row not found
}

/**
 * Get task by Task_ID
 */
function getTask(taskId) {
  return findRowByValue(SHEETS.TASKS_DB, 'Task_ID', taskId);
}

/**
 * Create a new task
 */
function createTask(taskData) {
  // Generate Task_ID if not provided
  if (!taskData.Task_ID) {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    taskData.Task_ID = `TASK-${timestamp}`;
  }
  
  // Set defaults
  if (!taskData.Status) {
    taskData.Status = TASK_STATUS.DRAFT;
  }
  if (!taskData.Created_Date) {
    taskData.Created_Date = new Date();
  }
  if (!taskData.Last_Updated) {
    taskData.Last_Updated = new Date();
  }
  
  const rowNum = addRow(SHEETS.TASKS_DB, taskData);
  logInteraction(taskData.Task_ID, `Task created: ${taskData.Task_Name}`);
  return taskData.Task_ID;
}

/**
 * Update a task
 */
function updateTask(taskId, updates) {
  updates.Last_Updated = new Date();
  const rowNum = updateRowByValue(SHEETS.TASKS_DB, 'Task_ID', taskId, updates);
  if (rowNum) {
    // Only log if we're not updating Interaction_Log itself (to avoid recursive loop)
    if (!updates.hasOwnProperty('Interaction_Log')) {
      logInteraction(taskId, `Task updated: ${JSON.stringify(updates)}`);
    }
  }
  return rowNum !== null;
}

/**
 * Get staff member by email
 */
function getStaff(email) {
  return findRowByValue(SHEETS.STAFF_DB, 'Email', email);
}

/**
 * Update staff member
 */
function updateStaff(email, updates) {
  updates.Last_Updated = new Date();
  return updateRowByValue(SHEETS.STAFF_DB, 'Email', email, updates);
}

/**
 * Get project by tag
 */
function getProject(projectTag) {
  return findRowByValue(SHEETS.PROJECTS_DB, 'Project_Tag', projectTag);
}

/**
 * Find staff email by name (fuzzy matching)
 * Tries multiple matching strategies for better accuracy
 */
function findStaffEmailByName(name) {
  if (!name) return null;
  
  try {
    const staff = getSheetData(SHEETS.STAFF_DB);
    if (!staff || staff.length === 0) {
      Logger.log('STAFF_DB is empty, cannot match name');
      return null;
    }
    
    const searchName = name.trim().toLowerCase();
    
    // Strategy 1: Exact match (case insensitive)
    let matched = staff.find(s => 
      s.Name && s.Name.toLowerCase() === searchName
    );
    if (matched) {
      Logger.log(`Exact match found: "${name}" -> ${matched.Email}`);
      return matched.Email;
    }
    
    // Strategy 2: Contains match (name contains search term or vice versa)
    matched = staff.find(s => {
      if (!s.Name) return false;
      const staffName = s.Name.toLowerCase();
      return staffName.includes(searchName) || searchName.includes(staffName);
    });
    if (matched) {
      Logger.log(`Contains match found: "${name}" -> ${matched.Email}`);
      return matched.Email;
    }
    
    // Strategy 3: First name match
    const firstName = searchName.split(' ')[0];
    if (firstName.length >= 3) {
      matched = staff.find(s => {
        if (!s.Name) return false;
        const staffFirstName = s.Name.toLowerCase().split(' ')[0];
        return staffFirstName === firstName;
      });
      if (matched) {
        Logger.log(`First name match found: "${name}" -> ${matched.Email}`);
        return matched.Email;
      }
    }
    
    // Strategy 4: Last name match
    const nameParts = searchName.split(' ');
    if (nameParts.length > 1) {
      const lastName = nameParts[nameParts.length - 1];
      if (lastName.length >= 3) {
        matched = staff.find(s => {
          if (!s.Name) return false;
          const staffNameParts = s.Name.toLowerCase().split(' ');
          if (staffNameParts.length > 1) {
            const staffLastName = staffNameParts[staffNameParts.length - 1];
            return staffLastName === lastName;
          }
          return false;
        });
        if (matched) {
          Logger.log(`Last name match found: "${name}" -> ${matched.Email}`);
          return matched.Email;
        }
      }
    }
    
    Logger.log(`No match found for name: "${name}"`);
    return null;
  } catch (error) {
    Logger.log(`Error finding staff by name: ${error.toString()}`);
    return null;
  }
}

/**
 * Find project tag by name (fuzzy matching)
 * Searches in Project_Name and Project_Tag fields
 */
function findProjectTagByName(searchText) {
  if (!searchText) return null;
  
  try {
    const projects = getSheetData(SHEETS.PROJECTS_DB);
    if (!projects || projects.length === 0) {
      Logger.log('PROJECTS_DB is empty, cannot match project');
      return null;
    }
    
    const searchLower = searchText.toLowerCase();
    
    // Strategy 1: Exact match in Project_Name
    let matched = projects.find(p => 
      p.Project_Name && p.Project_Name.toLowerCase() === searchLower
    );
    if (matched && matched.Project_Tag) {
      Logger.log(`Exact project match found: "${searchText}" -> ${matched.Project_Tag}`);
      return matched.Project_Tag;
    }
    
    // Strategy 2: Contains match in Project_Name
    matched = projects.find(p => {
      if (!p.Project_Name) return false;
      const projectName = p.Project_Name.toLowerCase();
      return projectName.includes(searchLower) || searchLower.includes(projectName);
    });
    if (matched && matched.Project_Tag) {
      Logger.log(`Contains project match found: "${searchText}" -> ${matched.Project_Tag}`);
      return matched.Project_Tag;
    }
    
    // Strategy 3: Match in Project_Tag itself
    matched = projects.find(p => 
      p.Project_Tag && p.Project_Tag.toLowerCase() === searchLower
    );
    if (matched && matched.Project_Tag) {
      Logger.log(`Project tag match found: "${searchText}" -> ${matched.Project_Tag}`);
      return matched.Project_Tag;
    }
    
    // Strategy 4: Word-by-word matching (for multi-word project names)
    const searchWords = searchLower.split(/\s+/).filter(w => w.length >= 3);
    if (searchWords.length > 0) {
      matched = projects.find(p => {
        if (!p.Project_Name) return false;
        const projectName = p.Project_Name.toLowerCase();
        return searchWords.some(word => projectName.includes(word));
      });
      if (matched && matched.Project_Tag) {
        Logger.log(`Word-based project match found: "${searchText}" -> ${matched.Project_Tag}`);
        return matched.Project_Tag;
      }
    }
    
    Logger.log(`No project match found for: "${searchText}"`);
    return null;
  } catch (error) {
    Logger.log(`Error finding project by name: ${error.toString()}`);
    return null;
  }
}

/**
 * Log interaction to task's Interaction_Log
 * This function directly updates the sheet to avoid recursive loops
 */
function logInteraction(taskId, message) {
  try {
    const task = getTask(taskId);
    if (!task) return;
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const logEntry = `${timestamp} - ${message}`;
    
    const currentLog = task.Interaction_Log || '';
    const newLog = currentLog ? `${currentLog}\n${logEntry}` : logEntry;
    
    // Directly update the sheet without calling updateTask to avoid recursive loop
    const updates = {
      Interaction_Log: newLog,
      Last_Updated: new Date()
    };
    updateRowByValue(SHEETS.TASKS_DB, 'Task_ID', taskId, updates);
  } catch (error) {
    Logger.log('Error in logInteraction: ' + error.toString());
    // Don't throw - logging failures shouldn't break the system
  }
}

/**
 * Log error to Error_Log sheet
 */
function logError(errorType, functionName, errorMessage, taskId = null, stackTrace = null) {
  try {
    const errorData = {
      Timestamp: new Date(),
      Error_Type: errorType,
      Function_Name: functionName,
      Error_Message: errorMessage,
      Task_ID: taskId || '',
      Stack_Trace: stackTrace || '',
      Resolved: false,
      Resolution_Notes: '',
    };
    
    addRow(SHEETS.ERROR_LOG, errorData);
  } catch (e) {
    Logger.log('Failed to log error: ' + e);
  }
}

/**
 * Get tasks by status
 */
function getTasksByStatus(status) {
  return findRowsByCondition(SHEETS.TASKS_DB, task => task.Status === status);
}

/**
 * Get tasks assigned to an email
 */
function getTasksByAssignee(email) {
  return findRowsByCondition(SHEETS.TASKS_DB, task => task.Assignee_Email === email);
}

/**
 * Get active task count for a staff member
 */
function getActiveTaskCount(email) {
  const tasks = getTasksByAssignee(email);
  return tasks.filter(task => 
    task.Status === TASK_STATUS.ACTIVE || 
    task.Status === TASK_STATUS.ASSIGNED
  ).length;
}

/**
 * Update active task count for all staff
 */
function updateAllActiveTaskCounts() {
  const staff = getSheetData(SHEETS.STAFF_DB);
  staff.forEach(member => {
    if (member.Email) {
      const count = getActiveTaskCount(member.Email);
      updateStaff(member.Email, { Active_Task_Count: count });
    }
  });
}

