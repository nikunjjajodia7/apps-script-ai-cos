/**
 * Dashboard Actions
 * Unified API endpoint for Lovable dashboard and other clients
 * Handles all actions, data reading, and file uploads
 */

/**
 * Handle POST requests
 * - Dashboard actions (approve, assign, modify, etc.)
 * - File uploads (voice recordings and meeting recordings)
 * 
 * Called from Lovable UI or any other client via the Apps Script web app URL
 */
function doPost(e) {
  try {
    const contentType = e.postData.type || '';
    
    // Handle file uploads (check for upload action or file data)
    let postData;
    try {
      postData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      // If not JSON, might be file upload
      if (contentType.indexOf('multipart') !== -1 || contentType.indexOf('application/octet-stream') !== -1) {
        return handleFileUpload(e);
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Invalid JSON in request body',
        message: parseError.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Check if this is a file upload request
    if (postData.action === 'upload_recording' || postData.fileData) {
      return handleFileUpload(e);
    }
    
    // Handle JSON actions
    const action = postData.action;
    const taskId = postData.taskId;
    const data = postData.data || {};
    
    Logger.log(`Dashboard action: ${action} for task ${taskId}`);
    
    let result = { success: false, message: 'Unknown action' };
    
    switch (action) {
      // Frontend API actions
      case 'create_task':
        result = handleCreateTask(postData);
        break;
      case 'update_task':
        result = handleUpdateTask(postData);
        break;
      // Category B1: Review_AI_Assist actions
      case 'approve_interpretation':
        result = handleApproveInterpretation(taskId, data);
        break;
      case 'modify_task':
        result = handleModifyTask(taskId, data);
        break;
      case 'assign_task':
        result = handleAssignTask(taskId, data);
        break;
      case 'rewrite_task':
        result = handleRewriteTask(taskId, data);
        break;
      case 'reject_task':
        result = handleRejectTask(taskId);
        break;
      
      // Category C1: Review_Date actions
      case 'approve_new_date':
        result = handleApproveNewDate(taskId);
        break;
      case 'reject_date_change':
        result = handleRejectDateChange(taskId);
        break;
      case 'negotiate_date':
        result = handleNegotiateDate(taskId, data);
        break;
      case 'force_meeting_date':
        result = handleForceMeeting(taskId);
        break;
      
      // Category C2: Review_Scope actions
      case 'provide_clarification':
        result = handleProvideClarification(taskId, data);
        break;
      case 'reduce_scope':
        result = handleReduceScope(taskId, data);
        break;
      case 'change_owner':
        result = handleChangeOwner(taskId, data);
        break;
      case 'increase_priority':
        result = handleIncreasePriority(taskId);
        break;
      case 'cancel_task_scope':
        result = handleCancelTask(taskId);
        break;
      
      // Category C3: Review_Role actions
      case 'accept_reassign':
        result = handleAcceptReassign(taskId, data);
        break;
      case 'override_role':
        result = handleOverrideRole(taskId, data);
        break;
      case 'redirect_task':
        result = handleRedirectTask(taskId, data);
        break;
      case 'assign_to_self':
        result = handleAssignToSelf(taskId);
        break;
      case 'cancel_task_role':
        result = handleCancelTask(taskId);
        break;
      
      // Category A1: Completion Review actions
      case 'approve_done':
        result = handleApproveDone(taskId);
        break;
      case 'reopen_task':
        result = handleReopenTask(taskId, data);
        break;
      case 'request_proof':
        result = handleRequestProof(taskId, data);
        break;
      
      // Category A2: Stagnation actions
      case 'force_meeting_stagnation':
        result = handleForceMeeting(taskId);
        break;
      case 'send_hard_nudge':
        result = handleSendHardNudge(taskId);
        break;
      case 'reassign_stagnation':
        result = handleReassign(taskId, data);
        break;
      case 'kill_task':
        result = handleCancelTask(taskId);
        break;
      
      // Category A3: Significant Update actions
      case 'acknowledge_update':
        result = handleAcknowledgeUpdate(taskId);
        break;
      case 'clarify_update':
        result = handleClarifyUpdate(taskId, data);
        break;
      case 'convert_to_meeting':
        result = handleConvertToMeeting(taskId, data);
        break;
      case 'add_to_weekly':
        result = handleAddToWeekly(taskId);
        break;
      case 'schedule_focus_time':
      case 'carve_personal_time':
        result = handleScheduleFocusTime(taskId);
        break;
      case 'mark_handled':
        result = handleMarkHandled(taskId);
        break;
      
      default:
        result = { success: false, message: `Unknown action: ${action}` };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'doPost', error.toString(), null, error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString(),
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Category B1: Review_AI_Assist handlers
function handleApproveInterpretation(taskId, data) {
  const task = getTask(taskId);
  if (!task) {
    return { success: false, message: 'Task not found' };
  }
  
  // Move to next status
  const newStatus = task.Assignee_Email ? TASK_STATUS.ASSIGNED : TASK_STATUS.NEW;
  updateTask(taskId, { Status: newStatus });
  
  if (newStatus === TASK_STATUS.ASSIGNED) {
    sendTaskAssignmentEmail(taskId);
  }
  
  return { success: true, message: 'Task approved and processed' };
}

function handleModifyTask(taskId, data) {
  const updates = {};
  if (data.taskName) updates.Task_Name = data.taskName;
  if (data.assigneeEmail) updates.Assignee_Email = data.assigneeEmail;
  if (data.dueDate) updates.Due_Date = data.dueDate;
  if (data.context) updates.Context_Hidden = data.context;
  
  updates.Status = updates.Assignee_Email ? TASK_STATUS.ASSIGNED : TASK_STATUS.NEW;
  updateTask(taskId, updates);
  
  if (updates.Status === TASK_STATUS.ASSIGNED) {
    sendTaskAssignmentEmail(taskId);
  }
  
  return { success: true, message: 'Task modified' };
}

function handleAssignTask(taskId, data) {
  if (!data.assigneeEmail) {
    return { success: false, message: 'Assignee email required' };
  }
  
  updateTask(taskId, {
    Assignee_Email: data.assigneeEmail,
    Status: TASK_STATUS.ASSIGNED,
  });
  
  sendTaskAssignmentEmail(taskId);
  return { success: true, message: 'Task assigned' };
}

function handleRewriteTask(taskId, data) {
  if (!data.newTaskName) {
    return { success: false, message: 'New task name required' };
  }
  
  updateTask(taskId, {
    Task_Name: data.newTaskName,
    Status: TASK_STATUS.DRAFT,
  });
  
  return { success: true, message: 'Task rewritten' };
}

function handleRejectTask(taskId) {
  updateTask(taskId, { Status: TASK_STATUS.CANCELLED });
  return { success: true, message: 'Task rejected' };
}

// Category C1: Review_Date handlers
function handleApproveNewDate(taskId) {
  const task = getTask(taskId);
  if (!task || !task.Proposed_Date) {
    return { success: false, message: 'No proposed date found' };
  }
  
  updateTask(taskId, {
    Due_Date: task.Proposed_Date,
    Proposed_Date: '',
    Status: TASK_STATUS.ACTIVE,
  });
  
  // Send confirmation email to assignee
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  const emailBody = `Hello ${assigneeName},\n\nYour request for a new due date has been approved. The new due date is ${task.Proposed_Date}.\n\nThank you,\n${CONFIG.EMAIL_SIGNATURE()}`;
  
  GmailApp.sendEmail(
    task.Assignee_Email,
    `Date Change Approved: ${task.Task_Name}`,
    emailBody
  );
  
  return { success: true, message: 'Date change approved' };
}

function handleRejectDateChange(taskId) {
  const task = getTask(taskId);
  if (!task) {
    return { success: false, message: 'Task not found' };
  }
  
  updateTask(taskId, {
    Proposed_Date: '',
    Status: TASK_STATUS.ACTIVE,
  });
  
  // Send email to assignee
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  const emailBody = `Hello ${assigneeName},\n\nThe original due date of ${task.Due_Date} stands. Please complete the task by this date.\n\nThank you,\n${CONFIG.EMAIL_SIGNATURE()}`;
  
  GmailApp.sendEmail(
    task.Assignee_Email,
    `Date Change Request: ${task.Task_Name}`,
    emailBody
  );
  
  return { success: true, message: 'Date change rejected' };
}

function handleNegotiateDate(taskId, data) {
  if (!data.newDate) {
    return { success: false, message: 'New date required' };
  }
  
  const task = getTask(taskId);
  updateTask(taskId, {
    Proposed_Date: data.newDate,
  });
  
  // Send negotiation email (could use AI to draft)
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  const emailBody = `Hello ${assigneeName},\n\nWe'd like to propose a compromise due date: ${data.newDate}. Please let us know if this works for you.\n\nThank you,\n${CONFIG.EMAIL_SIGNATURE()}`;
  
  GmailApp.sendEmail(
    task.Assignee_Email,
    `Date Negotiation: ${task.Task_Name}`,
    emailBody
  );
  
  return { success: true, message: 'Negotiation email sent' };
}

function handleForceMeeting(taskId) {
  const task = getTask(taskId);
  if (!task || !task.Assignee_Email) {
    return { success: false, message: 'Task has no assignee' };
  }
  
  updateTask(taskId, { Meeting_Action: MEETING_ACTION.ONE_ON_ONE });
  scheduleOneOnOne(taskId);
  
  return { success: true, message: 'Meeting scheduled' };
}

// Category C2: Review_Scope handlers
function handleProvideClarification(taskId, data) {
  if (!data.clarification) {
    return { success: false, message: 'Clarification text required' };
  }
  
  const task = getTask(taskId);
  const currentContext = task.Context_Hidden || '';
  const newContext = currentContext + '\n\nClarification: ' + data.clarification;
  
  updateTask(taskId, {
    Context_Hidden: newContext,
    Status: TASK_STATUS.ACTIVE,
  });
  
  // Send clarification email
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  const emailBody = `Hello ${assigneeName},\n\nHere's additional clarification for the task:\n\n${data.clarification}\n\nThank you,\n${CONFIG.EMAIL_SIGNATURE()}`;
  
  GmailApp.sendEmail(
    task.Assignee_Email,
    `Clarification: ${task.Task_Name}`,
    emailBody
  );
  
  return { success: true, message: 'Clarification sent' };
}

function handleReduceScope(taskId, data) {
  if (!data.newScope) {
    return { success: false, message: 'New scope required' };
  }
  
  updateTask(taskId, {
    Task_Name: data.newScope,
    Status: TASK_STATUS.ACTIVE,
  });
  
  return { success: true, message: 'Scope reduced' };
}

function handleChangeOwner(taskId, data) {
  if (!data.newAssigneeEmail) {
    return { success: false, message: 'New assignee email required' };
  }
  
  const task = getTask(taskId);
  const oldAssignee = task.Assignee_Email;
  
  updateTask(taskId, {
    Assignee_Email: data.newAssigneeEmail,
    Status: TASK_STATUS.ASSIGNED,
  });
  
  sendTaskAssignmentEmail(taskId);
  
  return { success: true, message: 'Owner changed' };
}

function handleIncreasePriority(taskId) {
  updateTask(taskId, { Priority: 'High' });
  return { success: true, message: 'Priority increased' };
}

function handleCancelTask(taskId) {
  updateTask(taskId, { Status: TASK_STATUS.CANCELLED });
  return { success: true, message: 'Task cancelled' };
}

// Category C3: Review_Role handlers
function handleAcceptReassign(taskId, data) {
  if (!data.newAssigneeEmail) {
    return { success: false, message: 'New assignee email required' };
  }
  
  updateTask(taskId, {
    Assignee_Email: data.newAssigneeEmail,
    Status: TASK_STATUS.ASSIGNED,
  });
  
  sendTaskAssignmentEmail(taskId);
  return { success: true, message: 'Task reassigned' };
}

function handleOverrideRole(taskId, data) {
  const task = getTask(taskId);
  if (!task || !task.Assignee_Email) {
    return { success: false, message: 'Task has no assignee' };
  }
  
  // Send firm but respectful email
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  const emailBody = `Hello ${assigneeName},\n\nThis task is indeed your responsibility. We expect you to proceed with it as assigned. If you have specific concerns, please let us know.\n\nThank you,\n${CONFIG.EMAIL_SIGNATURE()}`;
  
  GmailApp.sendEmail(
    task.Assignee_Email,
    `Task Assignment: ${task.Task_Name}`,
    emailBody
  );
  
  updateTask(taskId, { Status: TASK_STATUS.ACTIVE });
  
  return { success: true, message: 'Override email sent' };
}

function handleRedirectTask(taskId, data) {
  if (!data.newTeamEmail) {
    return { success: false, message: 'New team email required' };
  }
  
  updateTask(taskId, {
    Assignee_Email: data.newTeamEmail,
    Status: TASK_STATUS.ASSIGNED,
  });
  
  sendTaskAssignmentEmail(taskId);
  return { success: true, message: 'Task redirected' };
}

function handleAssignToSelf(taskId) {
  const bossEmail = CONFIG.BOSS_EMAIL();
  updateTask(taskId, {
    Assignee_Email: bossEmail,
    Status: TASK_STATUS.ACTIVE,
  });
  
  return { success: true, message: 'Task assigned to Boss' };
}

// Category A1: Completion Review handlers
function handleApproveDone(taskId) {
  updateTask(taskId, { Status: TASK_STATUS.DONE });
  return { success: true, message: 'Task marked as done' };
}

function handleReopenTask(taskId, data) {
  updateTask(taskId, {
    Status: TASK_STATUS.REOPENED,
    Context_Hidden: (getTask(taskId).Context_Hidden || '') + '\n\nReopened: ' + (data.reason || ''),
  });
  
  return { success: true, message: 'Task reopened' };
}

function handleRequestProof(taskId, data) {
  const task = getTask(taskId);
  if (!task || !task.Assignee_Email) {
    return { success: false, message: 'Task has no assignee' };
  }
  
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  const emailBody = `Hello ${assigneeName},\n\nBefore we can mark this task as complete, we need evidence or a demo of the result. Please provide the requested proof.\n\nThank you,\n${CONFIG.EMAIL_SIGNATURE()}`;
  
  GmailApp.sendEmail(
    task.Assignee_Email,
    `Proof Requested: ${task.Task_Name}`,
    emailBody
  );
  
  return { success: true, message: 'Proof request sent' };
}

// Category A2: Stagnation handlers
function handleSendHardNudge(taskId) {
  sendEscalationEmail(taskId);
  return { success: true, message: 'Hard nudge sent' };
}

function handleReassign(taskId, data) {
  if (!data.newAssigneeEmail) {
    return { success: false, message: 'New assignee email required' };
  }
  
  updateTask(taskId, {
    Assignee_Email: data.newAssigneeEmail,
    Status: TASK_STATUS.ASSIGNED,
  });
  
  sendTaskAssignmentEmail(taskId);
  return { success: true, message: 'Task reassigned' };
}

// Category A3: Significant Update handlers
function handleAcknowledgeUpdate(taskId) {
  updateTask(taskId, { Status: TASK_STATUS.ACTIVE });
  return { success: true, message: 'Update acknowledged' };
}

function handleClarifyUpdate(taskId, data) {
  if (!data.clarification) {
    return { success: false, message: 'Clarification required' };
  }
  
  logInteraction(taskId, `Boss requested clarification: ${data.clarification}`);
  return { success: true, message: 'Clarification logged' };
}

function handleConvertToMeeting(taskId, data) {
  updateTask(taskId, { Meeting_Action: MEETING_ACTION.ONE_ON_ONE });
  scheduleOneOnOne(taskId);
  return { success: true, message: 'Meeting scheduled' };
}

function handleAddToWeekly(taskId) {
  updateTask(taskId, { Meeting_Action: MEETING_ACTION.WEEKLY });
  addToWeeklyAgenda(taskId);
  return { success: true, message: 'Added to weekly agenda' };
}

function handleScheduleFocusTime(taskId) {
  updateTask(taskId, { Meeting_Action: MEETING_ACTION.SELF });
  scheduleFocusTime(taskId);
  return { success: true, message: 'Focus time scheduled' };
}

function handleMarkHandled(taskId) {
  updateTask(taskId, { Status: TASK_STATUS.ACTIVE });
  return { success: true, message: 'Marked as handled' };
}

// Frontend API handlers
function handleCreateTask(postData) {
  try {
    const { taskName, assigneeEmail, dueDate, priority, description } = postData;
    
    if (!taskName) {
      return { success: false, error: 'taskName is required' };
    }
    
    // Create task data object
    const taskData = {
      Task_Name: taskName,
      Status: assigneeEmail ? TASK_STATUS.ASSIGNED : TASK_STATUS.NEW,
      Assignee_Email: assigneeEmail || '',
      Priority: priority || 'Medium',
      Context_Hidden: description || '',
      Created_By: 'Manual'
    };
    
    // Parse due date if provided
    if (dueDate) {
      try {
        taskData.Due_Date = new Date(dueDate);
      } catch (e) {
        Logger.log('Invalid due date format: ' + dueDate);
      }
    }
    
    // Create the task
    const taskId = createTask(taskData);
    
    // Send assignment email if assignee is provided
    if (assigneeEmail && taskData.Status === TASK_STATUS.ASSIGNED) {
      try {
        sendTaskAssignmentEmail(taskId);
      } catch (emailError) {
        Logger.log('Could not send assignment email: ' + emailError.toString());
        // Don't fail the request if email fails
      }
    }
    
    return {
      success: true,
      data: { taskId: taskId }
    };
    
  } catch (error) {
    Logger.log('Error in handleCreateTask: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    return {
      success: false,
      error: error.toString()
    };
  }
}

function handleUpdateTask(postData) {
  try {
    const { taskId, ...fieldsToUpdate } = postData;
    
    if (!taskId) {
      return { success: false, error: 'taskId is required' };
    }
    
    // Check if task exists
    const existingTask = getTask(taskId);
    if (!existingTask) {
      return { success: false, error: 'Task not found' };
    }
    
    // Map frontend field names to sheet column names
    const updates = {};
    if (fieldsToUpdate.taskName !== undefined) updates.Task_Name = fieldsToUpdate.taskName;
    if (fieldsToUpdate.assigneeEmail !== undefined) updates.Assignee_Email = fieldsToUpdate.assigneeEmail;
    if (fieldsToUpdate.status !== undefined) updates.Status = fieldsToUpdate.status;
    if (fieldsToUpdate.priority !== undefined) updates.Priority = fieldsToUpdate.priority;
    if (fieldsToUpdate.dueDate !== undefined) {
      try {
        updates.Due_Date = new Date(fieldsToUpdate.dueDate);
      } catch (e) {
        Logger.log('Invalid due date format: ' + fieldsToUpdate.dueDate);
      }
    }
    if (fieldsToUpdate.description !== undefined) updates.Context_Hidden = fieldsToUpdate.description;
    if (fieldsToUpdate.projectId !== undefined) updates.Project_Tag = fieldsToUpdate.projectId;
    
    // Update the task
    const updated = updateTask(taskId, updates);
    
    if (!updated) {
      return { success: false, error: 'Failed to update task' };
    }
    
    // If assignee changed and task is now assigned, send email
    if (updates.Assignee_Email && updates.Status === TASK_STATUS.ASSIGNED) {
      try {
        sendTaskAssignmentEmail(taskId);
      } catch (emailError) {
        Logger.log('Could not send assignment email: ' + emailError.toString());
        // Don't fail the request if email fails
      }
    }
    
    return { success: true };
    
  } catch (error) {
    Logger.log('Error in handleUpdateTask: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Handle GET requests (for health checks, data reading, and web app deployment verification)
 */
function doGet(e) {
  try {
    const action = e.parameter.action || 'health';
    
    if (action === 'health') {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        service: 'Chief of Staff AI API',
        timestamp: new Date().toISOString(),
        version: '1.0.0'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'info') {
      return ContentService.createTextOutput(JSON.stringify({
        service: 'Chief of Staff AI API',
        availableActions: [
          'approve_interpretation', 'modify_task', 'assign_task', 'rewrite_task', 'reject_task',
          'approve_new_date', 'reject_date_change', 'negotiate_date', 'force_meeting_date',
          'provide_clarification', 'reduce_scope', 'change_owner', 'increase_priority', 'cancel_task_scope',
          'accept_reassign', 'override_role', 'redirect_task', 'assign_to_self', 'cancel_task_role',
          'approve_done', 'reopen_task', 'request_proof',
          'force_meeting_stagnation', 'send_hard_nudge', 'reassign_stagnation', 'kill_task',
          'acknowledge_update', 'clarify_update', 'convert_to_meeting', 'add_to_weekly', 'mark_handled'
        ],
        usage: 'POST JSON with { action, taskId, data } to this URL'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // get_tasks action (for frontend compatibility)
    if (action === 'get_tasks') {
      const status = e.parameter.status;
      const tasks = getSheetData(SHEETS.TASKS_DB);
      let filteredTasks = status 
        ? tasks.filter(t => t.Status === status)
        : tasks;
      
      // Format tasks to match expected frontend format
      const formattedTasks = filteredTasks.map(task => ({
        taskId: task.Task_ID || '',
        taskName: task.Task_Name || '',
        assigneeEmail: task.Assignee_Email || '',
        status: task.Status || '',
        priority: task.Priority || '',
        dueDate: task.Due_Date || '',
        description: task.Context_Hidden || '',
        projectId: task.Project_Tag || '',
        createdAt: task.Created_Date || ''
      }));
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: formattedTasks
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Data reading endpoints (legacy support)
    if (action === 'tasks') {
      const status = e.parameter.status;
      const tasks = getSheetData(SHEETS.TASKS_DB);
      const filteredTasks = status 
        ? tasks.filter(t => t.Status === status)
        : tasks;
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: filteredTasks,
        count: filteredTasks.length
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'staff') {
      const staff = getSheetData(SHEETS.STAFF_DB);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: staff,
        count: staff.length
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'projects') {
      const projects = getSheetData(SHEETS.PROJECTS_DB);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: projects,
        count: projects.length
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'task') {
      const taskId = e.parameter.taskId;
      if (!taskId) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'taskId parameter required'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const task = getTask(taskId);
      if (!task) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'Task not found'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: task
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      error: 'Unknown action',
      availableActions: ['health', 'info', 'tasks', 'staff', 'projects', 'task']
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle file upload (voice recordings)
 * Accepts file as base64 string in JSON or as blob data
 */
function handleFileUpload(e) {
  try {
    Logger.log('=== Handling file upload ===');
    
    let postData;
    let type, fileData, fileName, mimeType;
    
    // Try to parse as JSON first (for base64 encoded files)
    try {
      postData = JSON.parse(e.postData.contents);
      type = postData.type || 'task';
      fileData = postData.fileData; // base64 string
      fileName = postData.fileName;
      mimeType = postData.mimeType || 'audio/webm';
    } catch (parseError) {
      // If not JSON, try to get from parameters or raw blob
      type = e.parameter.type || 'task';
      fileData = e.postData.contents; // raw blob data
      fileName = e.parameter.fileName;
      mimeType = e.parameter.mimeType || 'audio/webm';
    }
    
    if (type !== 'task' && type !== 'meeting') {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Type must be "task" or "meeting"'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get folder ID based on type
    // Tasks go to VOICE_INBOX_FOLDER_ID, meetings go to MEETING_LAKE_FOLDER_ID
    const folderId = type === 'task' 
      ? CONFIG.VOICE_INBOX_FOLDER_ID()
      : CONFIG.MEETING_LAKE_FOLDER_ID();
    
    if (!folderId) {
      const configKey = type === 'task' 
        ? 'VOICE_INBOX_FOLDER_ID'
        : 'MEETING_LAKE_FOLDER_ID';
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: `Folder ID not configured. Please set ${configKey} in Config sheet.`
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (!fileData || fileData.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'No file data received'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Create blob from data
    let fileBlob;
    if (typeof fileData === 'string' && fileData.indexOf('data:') === 0) {
      // Base64 data URL
      const base64Data = fileData.split(',')[1];
      const bytes = Utilities.base64Decode(base64Data);
      fileBlob = Utilities.newBlob(bytes, mimeType);
    } else if (typeof fileData === 'string') {
      // Plain base64 string
      const bytes = Utilities.base64Decode(fileData);
      fileBlob = Utilities.newBlob(bytes, mimeType);
    } else {
      // Raw blob data
      fileBlob = Utilities.newBlob(fileData, mimeType);
    }
    
    // Generate file name with timestamp if not provided
    if (!fileName) {
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0] + '_' + 
                        new Date().toISOString().replace(/[:.]/g, '-').split('T')[1].split('.')[0];
      const extension = mimeType.includes('webm') ? '.webm' : 
                       mimeType.includes('m4a') ? '.m4a' : 
                       mimeType.includes('mp3') ? '.mp3' : '.webm';
      fileName = type === 'meeting' 
        ? `Meeting - ${timestamp}${extension}`
        : `Voice Command - ${timestamp}${extension}`;
    }
    
    // Get the target folder
    const folder = DriveApp.getFolderById(folderId);
    
    // Create file in Drive
    const file = folder.createFile(fileBlob.setName(fileName));
    
    Logger.log(`File uploaded: ${file.getName()} (ID: ${file.getId()})`);
    
    // If it's a task recording, trigger processing immediately (optional)
    if (type === 'task') {
      // The checkVoiceInbox trigger will process it, but we can also process immediately
      try {
        Utilities.sleep(1000); // Small delay to ensure file is fully saved
        processVoiceNote(file.getId());
        Logger.log('Voice note processed immediately');
      } catch (processError) {
        Logger.log('Note: Could not process immediately, will be processed by trigger: ' + processError.toString());
        // That's okay - the trigger will handle it
      }
    }
    
    // Return response matching frontend expected format
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: {
        fileId: file.getId(),
        fileName: file.getName()
      }
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error in handleFileUpload: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Deploy this script as a web app
 * Run this function once, then follow the deployment steps
 * 
 * Instructions:
 * 1. Run this function: deployWebApp()
 * 2. Go to Deploy > New deployment
 * 3. Select type: Web app
 * 4. Execute as: Me (your account)
 * 5. Who has access: Anyone (or Anyone with Google account if you want authentication)
 * 6. Click Deploy
 * 7. Copy the Web app URL
 * 8. Use that URL as APPS_SCRIPT_API_URL in your Lovable environment variables
 * 
 * This single endpoint handles:
 * - GET requests: health checks, reading tasks/staff/projects
 * - POST requests: actions (approve, assign, etc.)
 * - POST requests with files: voice/meeting recordings
 */
function deployWebApp() {
  Logger.log('=== Web App Deployment Instructions ===');
  Logger.log('');
  Logger.log('To deploy this script as a web app:');
  Logger.log('1. Click "Deploy" in the Apps Script editor');
  Logger.log('2. Click "New deployment"');
  Logger.log('3. Click the gear icon (⚙️) next to "Select type"');
  Logger.log('4. Choose "Web app"');
  Logger.log('5. Set "Execute as": Me');
  Logger.log('6. Set "Who has access": Anyone (or "Anyone with Google account" for auth)');
  Logger.log('7. Click "Deploy"');
  Logger.log('8. Copy the Web app URL');
  Logger.log('9. Add it to your Lovable environment variables as APPS_SCRIPT_API_URL');
  Logger.log('');
  Logger.log('This single endpoint handles everything:');
  Logger.log('  GET ?action=health - Health check');
  Logger.log('  GET ?action=tasks - Get all tasks');
  Logger.log('  GET ?action=tasks&status=Active - Get filtered tasks');
  Logger.log('  GET ?action=staff - Get all staff');
  Logger.log('  GET ?action=projects - Get all projects');
  Logger.log('  GET ?action=task&taskId=TASK-123 - Get single task');
  Logger.log('  POST (JSON) - Dashboard actions');
  Logger.log('  POST (FormData with file) - Upload recordings');
  Logger.log('');
  Logger.log('After deployment, test with:');
  Logger.log('  GET: [YOUR_WEB_APP_URL]?action=health');
  Logger.log('  GET: [YOUR_WEB_APP_URL]?action=tasks');
  Logger.log('  POST: [YOUR_WEB_APP_URL] with JSON body');
  Logger.log('');
  Logger.log('=== End Instructions ===');
  
  // Return a simple message
  return 'Check the execution log for deployment instructions';
}

