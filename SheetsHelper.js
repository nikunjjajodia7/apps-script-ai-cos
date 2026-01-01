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
 * Delete a row by matching a column value
 */
function deleteRowByValue(sheetName, matchColumn, matchValue) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      // Only headers or empty sheet
      return false;
    }
    
    const headers = data[0];
    const matchColumnIndex = headers.indexOf(matchColumn);
    
    if (matchColumnIndex === -1) {
      Logger.log(`Column "${matchColumn}" not found in sheet "${sheetName}"`);
      return false;
    }
    
    // Search from bottom to top to avoid index shifting issues
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][matchColumnIndex] === matchValue) {
        sheet.deleteRow(i + 1); // Delete the row (1-indexed)
        Logger.log(`Deleted row ${i + 1} with ${matchColumn}="${matchValue}"`);
        return true;
      }
    }
    
    Logger.log(`Row with ${matchColumn}="${matchValue}" not found`);
    return false; // Row not found
  } catch (error) {
    Logger.log(`Error in deleteRowByValue: ${error.toString()}`);
    throw error;
  }
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
  
  // Set defaults - new status system
  if (!taskData.Status) {
    taskData.Status = TASK_STATUS.AI_ASSIST;
  }
  if (!taskData.Created_Date) {
    taskData.Created_Date = new Date();
  }
  if (!taskData.Last_Updated) {
    taskData.Last_Updated = new Date();
  }
  
  // Store initial parameters snapshot for change detection
  const initialSnapshot = {
    dueDate: taskData.Due_Date || null,
    assignee: taskData.Assignee_Email || null,
    assigneeName: taskData.Assignee_Name || null,
    taskName: taskData.Task_Name || null,
    scope: taskData.Context_Hidden || null,
    projectTag: taskData.Project_Tag || null,
    createdAt: new Date().toISOString()
  };
  taskData.Initial_Parameters = JSON.stringify(initialSnapshot);
  
  const rowNum = addRow(SHEETS.TASKS_DB, taskData);
  logInteraction(taskData.Task_ID, `Task created: ${taskData.Task_Name}`);
  
  // Auto-link staff to project if both are present
  if (taskData.Assignee_Email && taskData.Project_Tag) {
    try {
      addStaffToProject(taskData.Assignee_Email, taskData.Project_Tag);
    } catch (error) {
      Logger.log(`Error linking staff to project: ${error.toString()}`);
      // Don't fail task creation if linking fails
    }
  }
  
  // Execute workflows for task_created trigger
  try {
    executeWorkflow('task_created', {
      taskId: taskData.Task_ID,
      task: taskData,
      status: taskData.Status
    });
  } catch (error) {
    Logger.log(`Error executing workflow for task_created: ${error.toString()}`);
    // Don't fail task creation if workflow fails
  }
  
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
    
    // Auto-link staff to project if both are being updated or one is new
    const task = getTask(taskId);
    if (task) {
      const assigneeEmail = updates.Assignee_Email !== undefined ? updates.Assignee_Email : task.Assignee_Email;
      const projectTag = updates.Project_Tag !== undefined ? updates.Project_Tag : task.Project_Tag;
      
      if (assigneeEmail && projectTag) {
        try {
          addStaffToProject(assigneeEmail, projectTag);
        } catch (error) {
          Logger.log(`Error linking staff to project during task update: ${error.toString()}`);
          // Don't fail task update if linking fails
        }
      }
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
 * Create a new staff member in Staff_DB
 * @param {string} name - Full name of staff member
 * @param {string} email - Email address (required, must be unique)
 * @param {object} additionalData - Optional additional fields (Role, Department, Manager_Email)
 * @returns {boolean} - True if created successfully, false if email already exists
 */
function createStaff(name, email, additionalData = {}) {
  if (!name || !email) {
    Logger.log('Error: Name and email are required to create staff member');
    return false;
  }
  
  // Check if staff with this email already exists
  const existingStaff = getStaff(email);
  if (existingStaff) {
    Logger.log(`Staff with email ${email} already exists: ${existingStaff.Name}`);
    return false;
  }
  
  // Check if staff with this name already exists (but different email)
  const existingByName = findStaffEmailByName(name);
  if (existingByName && existingByName !== email) {
    Logger.log(`Staff with name "${name}" already exists with different email: ${existingByName}`);
    // Still create, but log a warning
  }
  
  try {
    const staffData = {
      Name: name.trim(),
      Email: email.trim().toLowerCase(),
      Role: additionalData.Role || 'Team Member',
      Reliability_Score: '',
      Active_Task_Count: 0,
      Department: additionalData.Department || '',
      Manager_Email: additionalData.Manager_Email || '',
      Last_Updated: new Date()
    };
    
    const rowNumber = addRow(SHEETS.STAFF_DB, staffData);
    Logger.log(`Created new staff member: ${name} (${email}) at row ${rowNumber}`);
    return true;
  } catch (error) {
    Logger.log(`Error creating staff member: ${error.toString()}`);
    return false;
  }
}

/**
 * Get project by tag
 */
function getProject(projectTag) {
  return findRowByValue(SHEETS.PROJECTS_DB, 'Project_Tag', projectTag);
}

/**
 * Add a staff member to a project (bidirectional linking)
 * @param {string} staffEmail - Email of the staff member
 * @param {string} projectTag - Project tag to add
 * @returns {boolean} - True if successful
 */
function addStaffToProject(staffEmail, projectTag) {
  if (!staffEmail || !projectTag) {
    Logger.log('Error: Staff email and project tag are required');
    return false;
  }
  
  const staff = getStaff(staffEmail);
  const project = getProject(projectTag);
  
  if (!staff) {
    Logger.log(`Error: Staff member with email ${staffEmail} not found`);
    return false;
  }
  
  if (!project) {
    Logger.log(`Error: Project with tag ${projectTag} not found`);
    return false;
  }
  
  // Add project to staff's Project_Tags
  const staffProjects = (staff.Project_Tags || '').split(',').map(tag => tag.trim()).filter(tag => tag);
  if (!staffProjects.includes(projectTag)) {
    staffProjects.push(projectTag);
    updateStaff(staffEmail, { Project_Tags: staffProjects.join(', ') });
    Logger.log(`Added project ${projectTag} to staff ${staffEmail}`);
  }
  
  // Add staff to project's Team_Members
  const projectMembers = (project.Team_Members || '').split(',').map(email => email.trim().toLowerCase()).filter(email => email);
  const staffEmailLower = staffEmail.toLowerCase();
  if (!projectMembers.includes(staffEmailLower)) {
    projectMembers.push(staffEmailLower);
    updateRowByValue(SHEETS.PROJECTS_DB, 'Project_Tag', projectTag, { Team_Members: projectMembers.join(', ') });
    Logger.log(`Added staff ${staffEmail} to project ${projectTag}`);
  }
  
  return true;
}

/**
 * Remove a staff member from a project (bidirectional unlinking)
 * @param {string} staffEmail - Email of the staff member
 * @param {string} projectTag - Project tag to remove
 * @returns {boolean} - True if successful
 */
function removeStaffFromProject(staffEmail, projectTag) {
  if (!staffEmail || !projectTag) {
    Logger.log('Error: Staff email and project tag are required');
    return false;
  }
  
  const staff = getStaff(staffEmail);
  const project = getProject(projectTag);
  
  if (!staff || !project) {
    return false;
  }
  
  // Remove project from staff's Project_Tags
  const staffProjects = (staff.Project_Tags || '').split(',').map(tag => tag.trim()).filter(tag => tag && tag !== projectTag);
  updateStaff(staffEmail, { Project_Tags: staffProjects.join(', ') });
  
  // Remove staff from project's Team_Members
  const projectMembers = (project.Team_Members || '').split(',').map(email => email.trim().toLowerCase()).filter(email => email && email !== staffEmail.toLowerCase());
  updateRowByValue(SHEETS.PROJECTS_DB, 'Project_Tag', projectTag, { Team_Members: projectMembers.join(', ') });
  
  Logger.log(`Removed staff ${staffEmail} from project ${projectTag}`);
  return true;
}

/**
 * Get all projects a staff member is working on
 * @param {string} staffEmail - Email of the staff member
 * @returns {Array<string>} - Array of project tags
 */
function getStaffProjects(staffEmail) {
  const staff = getStaff(staffEmail);
  if (!staff || !staff.Project_Tags) {
    return [];
  }
  return staff.Project_Tags.split(',').map(tag => tag.trim()).filter(tag => tag);
}

/**
 * Get all staff members working on a project
 * @param {string} projectTag - Project tag
 * @returns {Array<object>} - Array of staff objects
 */
function getProjectStaff(projectTag) {
  const project = getProject(projectTag);
  if (!project || !project.Team_Members) {
    return [];
  }
  
  const memberEmails = project.Team_Members.split(',').map(email => email.trim().toLowerCase()).filter(email => email);
  return memberEmails.map(email => getStaff(email)).filter(staff => staff !== null && staff !== undefined);
}

/**
 * Sync staff-project relationships from tasks
 * This function analyzes all tasks and updates staff-project links based on task assignments
 * @returns {number} - Number of relationships updated
 */
function syncStaffProjectRelationships() {
  const tasks = getSheetData(SHEETS.TASKS_DB);
  let updatedCount = 0;
  
  tasks.forEach(task => {
    if (task.Assignee_Email && task.Project_Tag) {
      const staff = getStaff(task.Assignee_Email);
      const project = getProject(task.Project_Tag);
      
      if (staff && project) {
        // Check if relationship already exists
        const staffProjects = getStaffProjects(task.Assignee_Email);
        if (!staffProjects.includes(task.Project_Tag)) {
          if (addStaffToProject(task.Assignee_Email, task.Project_Tag)) {
            updatedCount++;
          }
        }
      }
    }
  });
  
  Logger.log(`Synced ${updatedCount} staff-project relationships from tasks`);
  return updatedCount;
}

/**
 * Calculate Levenshtein distance between two strings
 * Used for fuzzy string matching
 */
function levenshteinDistance(str1, str2) {
  const len1 = str1.length;
  const len2 = str2.length;
  const matrix = [];
  
  // Initialize matrix
  for (let i = 0; i <= len1; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= len2; j++) {
    matrix[0][j] = j;
  }
  
  // Fill matrix
  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      if (str1[i - 1] === str2[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,     // deletion
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j - 1] + 1  // substitution
        );
      }
    }
  }
  
  return matrix[len1][len2];
}

/**
 * Calculate similarity score between two strings (0-1)
 * Higher score = more similar
 */
function stringSimilarity(str1, str2) {
  const maxLen = Math.max(str1.length, str2.length);
  if (maxLen === 0) return 1.0;
  const distance = levenshteinDistance(str1.toLowerCase(), str2.toLowerCase());
  return 1 - (distance / maxLen);
}

/**
 * Simple phonetic matching - checks if names sound similar
 * Handles common variations like: anaya/anaaya, john/jon, michael/mike
 */
function phoneticSimilarity(name1, name2) {
  const n1 = name1.toLowerCase().replace(/[^a-z]/g, '');
  const n2 = name2.toLowerCase().replace(/[^a-z]/g, '');
  
  // Exact match after normalization
  if (n1 === n2) return 1.0;
  
  // Check if one contains the other (for nicknames)
  if (n1.includes(n2) || n2.includes(n1)) {
    const shorter = Math.min(n1.length, n2.length);
    const longer = Math.max(n1.length, n2.length);
    return shorter / longer; // Proportional similarity
  }
  
  // Check for common phonetic variations
  const variations = {
    'aa': 'a', 'ee': 'e', 'ii': 'i', 'oo': 'o', 'uu': 'u',
    'ph': 'f', 'ck': 'k', 'qu': 'kw'
  };
  
  let n1Norm = n1;
  let n2Norm = n2;
  for (const [pattern, replacement] of Object.entries(variations)) {
    n1Norm = n1Norm.replace(new RegExp(pattern, 'g'), replacement);
    n2Norm = n2Norm.replace(new RegExp(pattern, 'g'), replacement);
  }
  
  // Remove double letters
  n1Norm = n1Norm.replace(/(.)\1+/g, '$1');
  n2Norm = n2Norm.replace(/(.)\1+/g, '$1');
  
  if (n1Norm === n2Norm) return 0.9;
  
  // Use Levenshtein on normalized strings
  return stringSimilarity(n1Norm, n2Norm);
}

/**
 * Find staff email by name (enhanced fuzzy matching with phonetic support)
 * Tries multiple matching strategies including sound-alike matching
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
    
    // Strategy 5: Phonetic/sound-alike matching
    // Calculate similarity scores for all staff
    const similarities = staff.map(s => {
      if (!s.Name) return { staff: s, score: 0 };
      const staffName = s.Name.toLowerCase();
      const similarity = Math.max(
        phoneticSimilarity(searchName, staffName),
        phoneticSimilarity(firstName, staffName.split(' ')[0]),
        stringSimilarity(searchName, staffName)
      );
      return { staff: s, score: similarity };
    });
    
    // Sort by similarity score (highest first)
    similarities.sort((a, b) => b.score - a.score);
    
    // If best match has similarity > 0.7, use it
    if (similarities.length > 0 && similarities[0].score > 0.7) {
      const bestMatch = similarities[0];
      Logger.log(`Phonetic match found: "${name}" -> "${bestMatch.staff.Name}" (similarity: ${bestMatch.score.toFixed(2)})`);
      return bestMatch.staff.Email;
    }
    
    // Strategy 6: Partial phonetic match (for cases like "anaya" vs "anaaya")
    // Check if removing one character from either name makes them match
    for (let i = 0; i < searchName.length; i++) {
      const searchVariation = searchName.slice(0, i) + searchName.slice(i + 1);
      matched = staff.find(s => {
        if (!s.Name) return false;
        const staffName = s.Name.toLowerCase();
        if (staffName === searchVariation || searchVariation === staffName) return true;
        // Try removing character from staff name
        for (let j = 0; j < staffName.length; j++) {
          const staffVariation = staffName.slice(0, j) + staffName.slice(j + 1);
          if (staffVariation === searchName || searchName === staffVariation) return true;
        }
        return false;
      });
      if (matched) {
        Logger.log(`Partial phonetic match found: "${name}" -> "${matched.Name}"`);
        return matched.Email;
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
 * Get all staff members with the same name (for email selection)
 * Returns array of staff objects with same name but different emails
 */
function getStaffByName(name) {
  if (!name) return [];
  
  try {
    const staff = getSheetData(SHEETS.STAFF_DB);
    if (!staff || staff.length === 0) {
      return [];
    }
    
    const searchName = name.trim().toLowerCase();
    
    // Find all staff with matching name (exact or fuzzy)
    const matches = [];
    
    staff.forEach(s => {
      if (!s.Name) return;
      
      const staffName = s.Name.toLowerCase();
      
      // Exact match
      if (staffName === searchName) {
        matches.push(s);
      }
      // First name match
      else if (staffName.split(' ')[0] === searchName.split(' ')[0] && 
               searchName.split(' ')[0].length >= 3) {
        matches.push(s);
      }
      // Phonetic match
      else {
        const similarity = Math.max(
          phoneticSimilarity(searchName, staffName),
          phoneticSimilarity(searchName.split(' ')[0], staffName.split(' ')[0])
        );
        if (similarity > 0.7) {
          matches.push(s);
        }
      }
    });
    
    return matches;
  } catch (error) {
    Logger.log(`Error finding staff by name: ${error.toString()}`);
    return [];
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
    
    let currentLog = task.Interaction_Log || '';
    
    // Check size BEFORE adding new entry (Google Sheets limit is 50,000 characters)
    const maxLogSize = 45000; // Leave buffer before 50k limit
    const estimatedNewSize = currentLog.length + logEntry.length + 1; // +1 for newline
    
    // If adding this entry would exceed limit, truncate first
    if (estimatedNewSize > maxLogSize) {
      // Extract important data before truncating (Thread IDs, Message IDs)
      const threadIds = [];
      const messageIds = [];
      
      // Extract Thread IDs
      const threadIdMatches = currentLog.match(/Thread ID:\s*([a-zA-Z0-9_-]+)/gi);
      if (threadIdMatches) {
        threadIdMatches.forEach(match => {
          const threadId = match.replace(/Thread ID:\s*/i, '').trim();
          if (threadId && !threadIds.includes(threadId)) {
            threadIds.push(threadId);
          }
        });
      }
      
      // Extract Message IDs
      const messageIdMatches = currentLog.match(/Message ID:\s*([a-zA-Z0-9_-]+)/gi);
      if (messageIdMatches) {
        messageIdMatches.forEach(match => {
          const messageId = match.replace(/Message ID:\s*/i, '').trim();
          if (messageId && !messageIds.includes(messageId)) {
            messageIds.push(messageId);
          }
        });
      }
      
      // Split log into lines and keep only most recent entries
      const logLines = currentLog.split('\n');
      const maxLines = 150; // Keep last 150 entries
      
      // Keep most recent lines
      let truncatedLines = logLines.slice(-maxLines);
      
      // Reconstruct log with preserved Thread IDs and Message IDs
      let truncatedLog = truncatedLines.join('\n');
      
      // Add preserved Thread IDs if not already in truncated log
      threadIds.forEach(threadId => {
        if (!truncatedLog.includes(`Thread ID: ${threadId}`)) {
          truncatedLog = `${truncatedLog}\n${timestamp} - Thread ID: ${threadId} [preserved from truncated log]`;
        }
      });
      
      // Add preserved Message IDs if not already in truncated log
      messageIds.forEach(messageId => {
        if (!truncatedLog.includes(`Message ID: ${messageId}`)) {
          truncatedLog = `${truncatedLog}\n${timestamp} - Message ID: ${messageId} [preserved from truncated log]`;
        }
      });
      
      // Add truncation notice
      const truncationNotice = `${timestamp} - [Log truncated: kept last ${truncatedLines.length} entries, preserved ${threadIds.length} thread ID(s) and ${messageIds.length} message ID(s)]`;
      truncatedLog = `${truncatedLog}\n${truncationNotice}`;
      
      currentLog = truncatedLog;
      Logger.log(`Interaction_Log truncated for task ${taskId}: kept last ${truncatedLines.length} entries`);
    }
    
    const newLog = currentLog ? `${currentLog}\n${logEntry}` : logEntry;
    
    // Final safety check
    if (newLog.length > maxLogSize) {
      // Emergency truncation - keep only last 100 lines
      const logLines = newLog.split('\n');
      const emergencyLog = logLines.slice(-100).join('\n');
      const emergencyNotice = `${timestamp} - [Emergency truncation: kept last 100 entries]`;
      const finalLog = `${emergencyLog}\n${emergencyNotice}\n${logEntry}`;
      
      const updates = {
        Interaction_Log: finalLog,
        Last_Updated: new Date()
      };
      updateRowByValue(SHEETS.TASKS_DB, 'Task_ID', taskId, updates);
      Logger.log(`Emergency truncation applied for task ${taskId}`);
      return;
    }
    
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
 * Get processed message IDs from dedicated field
 * Returns Set for O(1) lookup
 */
function getProcessedMessageIds(taskId) {
  const task = getTask(taskId);
  if (!task || !task.Processed_Message_IDs) {
    return new Set();
  }
  
  try {
    const ids = JSON.parse(task.Processed_Message_IDs);
    return new Set(ids);
  } catch (e) {
    Logger.log(`Error parsing Processed_Message_IDs for ${taskId}: ${e.toString()}`);
    return new Set();
  }
}

/**
 * Mark message as processed
 * Keeps only last 500 message IDs to prevent field bloat
 */
function markMessageAsProcessed(taskId, messageId) {
  const processedIds = getProcessedMessageIds(taskId);
  if (processedIds.has(messageId)) {
    return false; // Already processed
  }
  
  processedIds.add(messageId);
  const idsArray = Array.from(processedIds);
  
  // Keep only last 500 message IDs (prevent field from getting too large)
  const trimmedIds = idsArray.slice(-500);
  
  updateTask(taskId, {
    Processed_Message_IDs: JSON.stringify(trimmedIds),
    Last_Updated: new Date()
  });
  
  return true; // Newly processed
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
    task.Status === TASK_STATUS.ON_TIME || 
    task.Status === TASK_STATUS.NOT_ACTIVE ||
    task.Status === TASK_STATUS.SLOW_PROGRESS
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

// ============================================
// PROMPT MANAGEMENT FUNCTIONS
// ============================================

/**
 * Get prompt sheet name based on type
 * @param {string} promptType - 'voice', 'email', or 'mom'
 * @returns {string} - Sheet name
 */
function getPromptSheetName(promptType) {
  switch (promptType.toLowerCase()) {
    case 'voice':
      return SHEETS.VOICE_PROMPTS;
    case 'email':
      return SHEETS.EMAIL_PROMPTS;
    case 'mom':
      return SHEETS.MOM_PROMPTS;
    default:
      throw new Error(`Unknown prompt type: ${promptType}. Use 'voice', 'email', or 'mom'.`);
  }
}

/**
 * Get a prompt from the appropriate prompts sheet
 * @param {string} promptName - The name/identifier of the prompt (e.g., 'parseVoiceCommand', 'classifyReplyType')
 * @param {string} promptType - Type of prompt: 'voice', 'email', or 'mom'
 * @returns {object|null} - Prompt object with Name, Content, Description, Last_Updated or null if not found
 */
function getPrompt(promptName, promptType) {
  try {
    const sheetName = getPromptSheetName(promptType);
    
    // Check if sheet exists
    let sheet;
    try {
      sheet = getSheet(sheetName);
    } catch (e) {
      Logger.log(`Prompt sheet ${sheetName} not found. Run createAllSheets() first.`);
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      // Only headers or empty
      return null;
    }
    
    const headers = data[0];
    const nameIndex = headers.indexOf('Name');
    const contentIndex = headers.indexOf('Content');
    const descriptionIndex = headers.indexOf('Description');
    const lastUpdatedIndex = headers.indexOf('Last_Updated');
    
    if (nameIndex === -1 || contentIndex === -1) {
      Logger.log(`Prompt sheet ${sheetName} is missing required columns (Name, Content)`);
      return null;
    }
    
    // Search for the prompt by name
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIndex] === promptName) {
        return {
          Name: data[i][nameIndex],
          Content: data[i][contentIndex],
          Description: descriptionIndex >= 0 ? data[i][descriptionIndex] : '',
          Last_Updated: lastUpdatedIndex >= 0 ? data[i][lastUpdatedIndex] : null
        };
      }
    }
    
    // Prompt not found
    return null;
  } catch (error) {
    Logger.log(`Error getting prompt ${promptName} from ${promptType}: ${error.toString()}`);
    return null;
  }
}

/**
 * Save a prompt to the appropriate prompts sheet
 * Creates or updates the prompt based on whether it already exists
 * @param {string} promptName - The name/identifier of the prompt
 * @param {string} promptType - Type of prompt: 'voice', 'email', or 'mom'
 * @param {string} content - The prompt content/template
 * @param {string} description - Optional description of what the prompt does
 * @returns {boolean} - True if saved successfully
 */
function savePrompt(promptName, promptType, content, description = '') {
  try {
    const sheetName = getPromptSheetName(promptType);
    
    // Check if sheet exists, create if not
    let sheet;
    try {
      sheet = getSheet(sheetName);
    } catch (e) {
      Logger.log(`Prompt sheet ${sheetName} not found. Creating it...`);
      const spreadsheet = getSpreadsheet();
      sheet = spreadsheet.insertSheet(sheetName);
      // Add headers
      sheet.appendRow(['Name', 'Content', 'Description', 'Last_Updated']);
      sheet.setFrozenRows(1);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIndex = headers.indexOf('Name');
    const contentIndex = headers.indexOf('Content');
    const descriptionIndex = headers.indexOf('Description');
    const lastUpdatedIndex = headers.indexOf('Last_Updated');
    
    // Ensure required columns exist
    if (nameIndex === -1) {
      Logger.log(`Adding Name column to ${sheetName}`);
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Name');
    }
    if (contentIndex === -1) {
      Logger.log(`Adding Content column to ${sheetName}`);
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Content');
    }
    
    // Refresh data after potential column additions
    const refreshedData = sheet.getDataRange().getValues();
    const refreshedHeaders = refreshedData[0];
    const finalNameIndex = refreshedHeaders.indexOf('Name');
    const finalContentIndex = refreshedHeaders.indexOf('Content');
    const finalDescriptionIndex = refreshedHeaders.indexOf('Description');
    const finalLastUpdatedIndex = refreshedHeaders.indexOf('Last_Updated');
    
    // Check if prompt already exists
    let existingRow = -1;
    for (let i = 1; i < refreshedData.length; i++) {
      if (refreshedData[i][finalNameIndex] === promptName) {
        existingRow = i + 1; // 1-indexed for sheet operations
        break;
      }
    }
    
    const timestamp = new Date();
    
    if (existingRow > 0) {
      // Update existing prompt
      sheet.getRange(existingRow, finalContentIndex + 1).setValue(content);
      if (finalDescriptionIndex >= 0 && description) {
        sheet.getRange(existingRow, finalDescriptionIndex + 1).setValue(description);
      }
      if (finalLastUpdatedIndex >= 0) {
        sheet.getRange(existingRow, finalLastUpdatedIndex + 1).setValue(timestamp);
      }
      Logger.log(`Updated prompt: ${promptName} in ${sheetName}`);
    } else {
      // Add new prompt
      const newRow = [];
      for (let i = 0; i < refreshedHeaders.length; i++) {
        if (i === finalNameIndex) {
          newRow.push(promptName);
        } else if (i === finalContentIndex) {
          newRow.push(content);
        } else if (i === finalDescriptionIndex) {
          newRow.push(description);
        } else if (i === finalLastUpdatedIndex) {
          newRow.push(timestamp);
        } else {
          newRow.push('');
        }
      }
      sheet.appendRow(newRow);
      Logger.log(`Created prompt: ${promptName} in ${sheetName}`);
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error saving prompt ${promptName} to ${promptType}: ${error.toString()}`);
    return false;
  }
}

/**
 * List all prompts of a given type
 * @param {string} promptType - Type of prompt: 'voice', 'email', or 'mom'
 * @returns {Array<object>} - Array of prompt objects
 */
function listPrompts(promptType) {
  try {
    const sheetName = getPromptSheetName(promptType);
    return getSheetData(sheetName);
  } catch (error) {
    Logger.log(`Error listing prompts for ${promptType}: ${error.toString()}`);
    return [];
  }
}

/**
 * Delete a prompt from the appropriate prompts sheet
 * @param {string} promptName - The name/identifier of the prompt
 * @param {string} promptType - Type of prompt: 'voice', 'email', or 'mom'
 * @returns {boolean} - True if deleted successfully
 */
function deletePrompt(promptName, promptType) {
  try {
    const sheetName = getPromptSheetName(promptType);
    return deleteRowByValue(sheetName, 'Name', promptName);
  } catch (error) {
    Logger.log(`Error deleting prompt ${promptName} from ${promptType}: ${error.toString()}`);
    return false;
  }
}

