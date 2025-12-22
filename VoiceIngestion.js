/**
 * Voice Task Ingestion Module
 * Processes voice commands from Voice_Inbox folder
 */

/**
 * Drive trigger handler for new files in Voice_Inbox
 */
function onVoiceFileAdded(e) {
  try {
    const file = e.source;
    const fileId = file.getId();
    
    // Check if file is in Voice_Inbox folder
    const voiceInboxFolderId = CONFIG.VOICE_INBOX_FOLDER_ID();
    if (!voiceInboxFolderId) {
      Logger.log('VOICE_INBOX_FOLDER_ID not configured');
      return;
    }
    
    const folders = file.getParents();
    let isInVoiceInbox = false;
    
    folders.forEach(folder => {
      if (folder.getId() === voiceInboxFolderId) {
        isInVoiceInbox = true;
      }
    });
    
    if (isInVoiceInbox) {
      // Process the voice file
      processVoiceNote(fileId);
    }
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'onVoiceFileAdded', error.toString(), null, error.stack);
  }
}

/**
 * Main function to process a voice note
 */
function processVoiceNote(fileId) {
  try {
    Logger.log(`=== Processing voice note: ${fileId} ===`);
    
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    Logger.log(`File name: ${fileName}`);
    
    // Check file type (should be audio)
    const mimeType = file.getMimeType();
    Logger.log(`File MIME type: ${mimeType}`);
    
    if (!mimeType.startsWith('audio/') && mimeType !== 'video/mp4' && mimeType !== 'application/octet-stream') {
      Logger.log(`WARNING: File ${fileName} may not be an audio file (${mimeType})`);
      Logger.log('Will attempt to process anyway...');
      // Don't return - try to process it anyway
    }
    
    // Transcribe audio using Gemini
    let transcript;
    try {
      Logger.log('Starting audio transcription with Gemini...');
      transcript = transcribeAudio(fileId);
      Logger.log(`Transcript received (${transcript.length} characters): ${transcript.substring(0, 100)}...`);
    } catch (error) {
      Logger.log('ERROR: Transcription failed: ' + error.toString());
      Logger.log('Creating task with low confidence...');
      try {
        logError(ERROR_TYPE.API_ERROR, 'processVoiceNote', `Transcription failed: ${error.toString()}`, null, error.stack);
      } catch (e) {
        // Ignore logging errors
      }
      // Create task with low confidence
      const taskId = createTaskFromVoice({
        task_name: `Voice command from ${fileName} (transcription failed)`,
        assignee_email: null,
        assignee_name: null,
        due_date: null,
        context: `Original file: ${fileName}. Transcription failed: ${error.toString()}`,
        tone: 'normal',
        confidence: 0.2,
        ambiguities: ['Transcription failed']
      });
      Logger.log('Created task with low confidence. Task ID: ' + taskId);
      return;
    }
    
    // Parse voice command
    Logger.log('Parsing voice command...');
    const parsedData = parseVoiceCommand(transcript);
    Logger.log(`Parsed data: ${JSON.stringify(parsedData)}`);
    
    // Create task from parsed data
    Logger.log('Creating task from parsed data...');
    const taskId = createTaskFromVoice(parsedData);
    Logger.log(`Task created with ID: ${taskId}`);
    
    // Boss confirmation email disabled - no longer sending emails when tasks are created
    // Logger.log('Sending confirmation email to Boss...');
    // const summary = formatTaskSummary(parsedData, taskId);
    // try {
    //   sendBossConfirmation(taskId, summary);
    //   Logger.log('Confirmation email sent');
    // } catch (e) {
    //   Logger.log('Warning: Could not send confirmation email: ' + e.toString());
    // }
    
    Logger.log(`=== Voice note processed successfully. Task ID: ${taskId} ===`);
    
  } catch (error) {
    Logger.log('ERROR in processVoiceNote: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    try {
      logError(ERROR_TYPE.UNKNOWN_ERROR, 'processVoiceNote', error.toString(), null, error.stack);
    } catch (e) {
      // If we can't log, that's okay
    }
  }
}

/**
 * Create task from parsed voice command data
 */
function createTaskFromVoice(parsedData) {
  try {
    // Resolve assignee email if name was provided
    let assigneeEmail = parsedData.assignee_email;
    
    if (!assigneeEmail && parsedData.assignee_name) {
      // Try to find staff by name
      const staff = getSheetData(SHEETS.STAFF_DB);
      const matchedStaff = staff.find(s => 
        s.Name && s.Name.toLowerCase().includes(parsedData.assignee_name.toLowerCase())
      );
      if (matchedStaff) {
        assigneeEmail = matchedStaff.Email;
      }
    }
    
    // Determine status based on confidence
    let status = TASK_STATUS.DRAFT;
    if (parsedData.confidence < CONFIG.AI_CONFIDENCE_THRESHOLD() || parsedData.ambiguities.length > 0) {
      status = TASK_STATUS.REVIEW_AI_ASSIST;
    } else if (assigneeEmail) {
      status = TASK_STATUS.NEW;
    }
    
    // Prepare task data
    const taskData = {
      Task_Name: parsedData.task_name || 'Untitled Task',
      Status: status,
      Assignee_Email: assigneeEmail || '',
      Due_Date: parsedData.due_date || '',
      Project_Tag: parsedData.project_tag || '',
      AI_Confidence: parsedData.confidence || 0.5,
      Tone_Detected: parsedData.tone || 'normal',
      Context_Hidden: parsedData.context || '',
      Created_By: 'Voice',
      Priority: parsedData.priority || (parsedData.tone === 'urgent' ? 'High' : 'Medium'),
    };
    
    // Add due time to context if specified
    if (parsedData.due_time) {
      taskData.Context_Hidden = (taskData.Context_Hidden || '') + 
        (taskData.Context_Hidden ? '\n' : '') + `Due time: ${parsedData.due_time}`;
    }
    
    // Add ambiguities to context if any
    if (parsedData.ambiguities && parsedData.ambiguities.length > 0) {
      taskData.Context_Hidden = (taskData.Context_Hidden || '') + 
        '\n\nAmbiguities: ' + parsedData.ambiguities.join(', ');
    }
    
    // Create task
    const taskId = createTask(taskData);
    
    // If assignee is known and confidence is high, auto-assign
    if (assigneeEmail && status === TASK_STATUS.NEW) {
      updateTask(taskId, { Status: TASK_STATUS.ASSIGNED });
      // Trigger email assignment (will be handled by email module)
      Utilities.sleep(1000); // Small delay to ensure task is saved
      try {
        sendTaskAssignmentEmail(taskId);
      } catch (error) {
        Logger.log(`Failed to send assignment email: ${error}`);
      }
    }
    
    return taskId;
  } catch (error) {
    logError(ERROR_TYPE.DATA_ERROR, 'createTaskFromVoice', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Format task summary for Boss confirmation email
 */
function formatTaskSummary(parsedData, taskId) {
  let summary = `Task: ${parsedData.task_name || 'Untitled Task'}\n`;
  
  if (parsedData.assignee_name || parsedData.assignee_email) {
    summary += `Assigned to: ${parsedData.assignee_name || parsedData.assignee_email}\n`;
  }
  
  if (parsedData.due_date) {
    summary += `Due date: ${parsedData.due_date}\n`;
  }
  
  if (parsedData.context) {
    summary += `Context: ${parsedData.context}\n`;
  }
  
  summary += `Confidence: ${(parsedData.confidence * 100).toFixed(0)}%\n`;
  
  if (parsedData.ambiguities && parsedData.ambiguities.length > 0) {
    summary += `\nNote: Some elements were ambiguous. Please review in dashboard.`;
  }
  
  return summary;
}

/**
 * Manual trigger function for testing
 */
function testProcessVoiceNote() {
  // Get a test file from Voice_Inbox
  const voiceInboxFolderId = CONFIG.VOICE_INBOX_FOLDER_ID();
  if (!voiceInboxFolderId) {
    Logger.log('VOICE_INBOX_FOLDER_ID not configured');
    return;
  }
  
  const folder = DriveApp.getFolderById(voiceInboxFolderId);
  const files = folder.getFiles();
  
  if (files.hasNext()) {
    const file = files.next();
    processVoiceNote(file.getId());
  } else {
    Logger.log('No files found in Voice_Inbox');
  }
}

