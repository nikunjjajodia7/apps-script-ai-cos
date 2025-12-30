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
  let file = null;
  let fileName = '';
  let voiceInboxFolderId = null;
  
  try {
    Logger.log(`=== Processing voice note: ${fileId} ===`);
    
    file = DriveApp.getFileById(fileId);
    fileName = file.getName();
    Logger.log(`File name: ${fileName}`);
    
    // Check if file has already been processed (marked with prefix)
    if (fileName.startsWith('[PROCESSED]') || fileName.startsWith('[UNCLEAR]')) {
      Logger.log(`File ${fileName} has already been processed, skipping`);
      return;
    }
    
    // Get Voice_Inbox folder ID for later cleanup
    voiceInboxFolderId = CONFIG.VOICE_INBOX_FOLDER_ID();
    
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
    let transcriptionFailed = false;
    try {
      Logger.log('Starting audio transcription with Gemini...');
      transcript = transcribeAudio(fileId);
      Logger.log(`Transcript received (${transcript.length} characters): ${transcript.substring(0, 100)}...`);
    } catch (error) {
      Logger.log('ERROR: Transcription failed: ' + error.toString());
      Logger.log('Creating task with low confidence...');
      transcriptionFailed = true;
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
        ambiguities: ['Transcription failed'],
        raw_transcript: `[TRANSCRIPTION FAILED] Error: ${error.toString()}`
      });
      Logger.log('Created task with low confidence. Task ID: ' + taskId);
      
      // Mark file as unclear so it won't be processed again
      markFileAsUnclear(file, fileName);
      return;
    }
    
    // Parse voice command
    Logger.log('Parsing voice command...');
    const parsedData = parseVoiceCommand(transcript);
    Logger.log(`Parsed data: ${JSON.stringify(parsedData)}`);
    
    // Add the raw transcript to parsed data for storage
    parsedData.raw_transcript = transcript;
    
    // Create task from parsed data
    Logger.log('Creating task from parsed data...');
    const taskId = createTaskFromVoice(parsedData);
    Logger.log(`Task created with ID: ${taskId}`);
    
    // Check confidence level to determine file handling
    const confidence = parsedData.confidence || 0.5;
    const confidenceThreshold = CONFIG.AI_CONFIDENCE_THRESHOLD();
    const hasAmbiguities = parsedData.ambiguities && parsedData.ambiguities.length > 0;
    const isClear = confidence >= confidenceThreshold && !hasAmbiguities;
    
    if (isClear) {
      // Voice was clear - delete the file from Voice_Inbox
      Logger.log(`Voice note was clear (confidence: ${confidence}). Deleting file from Voice_Inbox...`);
      try {
        file.setTrashed(true);
        Logger.log(`File ${fileName} deleted successfully`);
      } catch (deleteError) {
        Logger.log(`Warning: Could not delete file: ${deleteError.toString()}`);
        // Fallback: rename to mark as processed
        try {
          file.setName(`[PROCESSED] ${fileName}`);
          Logger.log(`File renamed to mark as processed`);
        } catch (renameError) {
          Logger.log(`Warning: Could not rename file: ${renameError.toString()}`);
        }
      }
    } else {
      // Voice was unclear - mark file so it won't be processed again
      Logger.log(`Voice note was unclear (confidence: ${confidence}, ambiguities: ${hasAmbiguities}). Marking file to prevent reprocessing...`);
      markFileAsUnclear(file, fileName);
    }
    
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
 * Mark file as unclear so it won't be processed again
 */
function markFileAsUnclear(file, originalFileName) {
  try {
    if (!file) return;
    
    // Check if already marked
    if (originalFileName.startsWith('[UNCLEAR]')) {
      return;
    }
    
    // Rename file to mark it as unclear
    const newName = `[UNCLEAR] ${originalFileName}`;
    file.setName(newName);
    Logger.log(`File marked as unclear: ${newName}`);
  } catch (error) {
    Logger.log(`Warning: Could not mark file as unclear: ${error.toString()}`);
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
      // Try to find staff by name (improved matching)
      assigneeEmail = findStaffEmailByName(parsedData.assignee_name);
      if (assigneeEmail) {
        Logger.log(`Matched name "${parsedData.assignee_name}" to email: ${assigneeEmail}`);
      }
    }
    
    // Resolve project tag if project name was provided
    let projectTag = parsedData.project_tag;
    
    if (!projectTag) {
      // Try to find project by name in task name or context
      const searchText = `${parsedData.task_name || ''} ${parsedData.context || ''}`.toLowerCase();
      projectTag = findProjectTagByName(searchText);
      if (projectTag) {
        Logger.log(`Matched project from context to tag: ${projectTag}`);
      }
    } else {
      // If project_tag was provided, verify it exists in PROJECTS_DB
      const project = getProject(projectTag);
      if (!project) {
        Logger.log(`Project tag "${projectTag}" not found in PROJECTS_DB, trying to match by name...`);
        projectTag = findProjectTagByName(projectTag);
      }
    }
    
    // Determine status based on confidence - using new status system
    let status = TASK_STATUS.AI_ASSIST;  // Default to AI assist for review
    if (parsedData.confidence >= CONFIG.AI_CONFIDENCE_THRESHOLD() && 
        (!parsedData.ambiguities || parsedData.ambiguities.length === 0)) {
      // High confidence, no ambiguities - can proceed
      status = assigneeEmail ? TASK_STATUS.NOT_ACTIVE : TASK_STATUS.AI_ASSIST;
    }
    
    // Prepare task data (no priority field - using status system)
    const taskData = {
      Task_Name: parsedData.task_name || 'Untitled Task',
      Status: status,
      Assignee_Email: assigneeEmail || '',
      Due_Date: parsedData.due_date || '',
      Project_Tag: projectTag || '',
      AI_Confidence: parsedData.confidence || 0.5,
      Tone_Detected: parsedData.tone || 'normal',
      Context_Hidden: parsedData.context || '',
      Created_By: 'Voice',
      Voice_Transcript: parsedData.raw_transcript || '',
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
    
    // Add name matching debug info if available
    if (parsedData.name_heard_as) {
      taskData.Context_Hidden = (taskData.Context_Hidden || '') + 
        '\n\nVoice Recognition: Heard "' + parsedData.name_heard_as + '"' +
        (parsedData.assignee_name ? ' → Matched to "' + parsedData.assignee_name + '"' : ' → No match found');
    }
    
    // Add due date text for reference
    if (parsedData.due_date_text) {
      taskData.Context_Hidden = (taskData.Context_Hidden || '') + 
        '\n\nDue Date Spoken: "' + parsedData.due_date_text + '"' +
        (parsedData.due_date ? ' → Interpreted as ' + parsedData.due_date : '');
    }
    
    // Create task
    const taskId = createTask(taskData);
    
    // If assignee is known and status is not_active, send assignment email
    if (assigneeEmail && status === TASK_STATUS.NOT_ACTIVE) {
      // Task is already set to NOT_ACTIVE, trigger email assignment
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

